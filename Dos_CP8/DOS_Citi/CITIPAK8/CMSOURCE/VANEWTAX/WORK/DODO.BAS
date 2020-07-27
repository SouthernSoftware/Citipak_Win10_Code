DECLARE SUB OpenDCCustFile (NumOfDcRecs%, DCFile%)
DECLARE SUB OpenDCCustIdxFile (NumOfDCIdxRecs%, DCIdxFile%)
DEFINT A-Z

DEFSNG A-Z
DEFINT A-Z
SUB PostPayments (Operator)

  PayTName$ = "DCPYT" + QPTrim$(STR$(Operator)) + ".DAT"

  OpenDCCustIdxFile NumOfDCIdxRecs, DCIdxFile
  OpenDCCustFile NumOfDcRecs, DCFile
  DCEditRecLen = LEN(EditPaymentRec(1))
  DCEdFile = FREEFILE

  OPEN PayTName$ FOR RANDOM ACCESS READ WRITE SHARED AS DCEdFile LEN = DCEditRecLen
  NumOfDcRecs = LOF(DCEdFile) \ DCEditRecLen
  
  REDIM DCTransRec(1) AS DCTransRecType
  DCTransRecLen = LEN(DCTransRec(1))
  DCTransFile = FREEFILE
  OPEN "DCTrans.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCTransFile LEN = DCTransRecLen
  NumOfTransRecs = LOF(DCTransFile) \ DCTransRecLen
  NextTransRec = NumOfTransRecs + 1
  
  DO
    Cnt = Cnt + 1
    GET DCEdFile, Cnt, EditPaymentRec(1)
    
    IF EditPaymentRec(1).Amount <> 0 THEN
      
      GOSUB OldVehPost
      GOSUB UpdateVehRecord
      
      GET DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
      help$ = "Posting: " + LEFT$(DCCustRec(1).BillName, 30)
      PrintHelp help$
      
      ' Post Charge First to Offset Payment of Decal
      DCTransRec(1).CustomerNumber = EditPaymentRec(1).CustNumber
      DCTransRec(1).TransDate = EditPaymentRec(1).TranDate
      DCTransRec(1).TransAmount = EditPaymentRec(1).Amount
      DCTransRec(1).TransType = 1               ' Type 1 = Charge
      DCTransRec(1).TransDesc = "Decal Purchase " + EditPaymentRec(1).Sticker
      DCTransRec(1).CashAmount = EditPaymentRec(1).CashAmt
      DCTransRec(1).ChkAmount = EditPaymentRec(1).CheckAmt
      DCTransRec(1).BalanceAfterTrans = DCTransRec(1).BalanceAfterTrans + EditPaymentRec(1).Amount
      DCTransRec(1).MakeModel = EditPaymentRec(1).MakeModel
      DCTransRec(1).StateTag = EditPaymentRec(1).StateTag
      DCTransRec(1).Sticker = EditPaymentRec(1).Sticker
      DCTransRec(1).ExpireDate = EditPaymentRec(1).ExpDate
      DCTransRec(1).DecalCat = EditPaymentRec(1).DecalCat
      DCTransRec(1).ExtraRoom = ""
      DCTransRec(1).NextTrans = 0
      PUT DCTransFile, NextTransRec, DCTransRec(1)
      
      
      GET DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
      DCCustRec(1).AcctBal = DCCustRec(1).AcctBal + EditPaymentRec(1).Amount
      PUT DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
      
      IF DCCustRec(1).FirstTrans = 0 THEN
        DCCustRec(1).FirstTrans = NextTransRec
        DCCustRec(1).LastTrans = NextTransRec
        PUT DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
      ELSE
        Prev! = DCCustRec(1).LastTrans
        DCCustRec(1).LastTrans = NextTransRec
        PUT DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
        GET DCTransFile, Prev!, DCTransRec(1)
        DCTransRec(1).NextTrans = NextTransRec
        PUT DCTransFile, Prev!, DCTransRec(1)
      END IF
      NextTransRec = NextTransRec + 1
      
      ' Post Transaction Record First
      DCTransRec(1).CustomerNumber = EditPaymentRec(1).CustNumber
      DCTransRec(1).TransDate = EditPaymentRec(1).TranDate
      DCTransRec(1).TransAmount = EditPaymentRec(1).Amount
      DCTransRec(1).TransType = 2               ' Type 2 = Payment
      DCTransRec(1).TransDesc = EditPaymentRec(1).Desc
      DCTransRec(1).CashAmount = EditPaymentRec(1).CashAmt
      DCTransRec(1).ChkAmount = EditPaymentRec(1).CheckAmt
      DCTransRec(1).BalanceAfterTrans = DCTransRec(1).BalanceAfterTrans - EditPaymentRec(1).Amount
      DCTransRec(1).ExtraRoom = ""
      DCTransRec(1).NextTrans = 0
      PUT DCTransFile, NextTransRec, DCTransRec(1)
      
      GET DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
      DCCustRec(1).AcctBal = DCCustRec(1).AcctBal - EditPaymentRec(1).Amount
      DCCustRec(1).LICENSE = EditPaymentRec(1).Sticker
      PUT DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
      
      IF DCCustRec(1).FirstTrans = 0 THEN
        DCCustRec(1).FirstTrans = NextTransRec
        DCCustRec(1).LastTrans = NextTransRec
        PUT DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
      ELSE
        Prev! = DCCustRec(1).LastTrans
        DCCustRec(1).LastTrans = NextTransRec
        PUT DCFile, VAL(EditPaymentRec(1).CustNumber), DCCustRec(1)
        GET DCTransFile, Prev!, DCTransRec(1)
        DCTransRec(1).NextTrans = NextTransRec
        PUT DCTransFile, Prev!, DCTransRec(1)
      END IF
      NextTransRec = NextTransRec + 1
    END IF
    
  LOOP UNTIL Cnt = NumOfDcRecs
  CLOSE
  KILL PayTName$
  ' Show All Posted
  PRINT CHR$(7);
'  SaveScrn TempScrn()
'  DisplayDCScrn "DCPOSTED"
'  WaitForAction
'  RestScrn TempScrn()
  LOCATE , , 1
  CLOSE
  
  EXIT SUB

OldVehPost:
  IF EditPaymentRec(1).Owner = "Y" THEN RETURN
  REDIM DCOVRec(1) AS DCOldVehType
  DCOVRecLen = LEN(DCOVRec(1))
  DCOVFile = FREEFILE
  OPEN "DCOLDVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCOVFile LEN = DCOVRecLen
  NumOfOVRecs! = LOF(DCOVFile) \ DCOVRecLen
  NextOVRec! = NumOfOVRecs! + 1
  DCOVRec(1).Make = LTRIM$(EditPaymentRec(1).OldMake)
  DCOVRec(1).Year = LTRIM$(EditPaymentRec(1).OldDesc)
  DCOVRec(1).CustRec = VAL(EditPaymentRec(1).CustNumber)
  DCOVRec(1).MoreRoom = ""
  PUT DCOVFile, NextOVRec!, DCOVRec(1)
  CLOSE DCOVFile
  RETURN
  
UpdateVehRecord:
  DCVehRecLen = LEN(DCVRec(1))
  DCVFile = FREEFILE
  OPEN "DCVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS DCVFile LEN = DCVehRecLen
  NumOfVRecs! = LOF(DCVFile) \ DCVehRecLen
  VehRecord! = EditPaymentRec(1).VehRecord
  IF VehRecord! <= 0 OR VehRecord! > NumOfVRecs! THEN CLOSE DCVFile: RETURN
  GET DCVFile, VehRecord!, DCVRec(1)
  DCVRec(1).ExpireDate = EditPaymentRec(1).ExpDate
  DCVRec(1).Sticker = LTRIM$(EditPaymentRec(1).Sticker)
  DCVRec(1).Valid = "Y"
  DCVRec(1).Fee = EditPaymentRec(1).Amount
  PUT DCVFile, VehRecord!, DCVRec(1)
  CLOSE DCVFile
  RETURN
  
END SUB

