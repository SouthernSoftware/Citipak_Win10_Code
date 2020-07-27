DECLARE SUB GateCity ()
DEFINT A-Z

'$INCLUDE: 'pro7.bi'
'$INCLUDE: 'ap.bi'
'$INCLUDE: 'cp65ap.bi'

'--Procedures in APAUX.BAS
DECLARE SUB OpenVendorFile (VendorFile%, NumVRecs%)
DECLARE SUB OpenAPLedgerFile (APLedgerFile%, NumTrans&, RecLen%)
DECLARE SUB OpenAPDistFile (APDistFile%, NumDistRecs&, RecLen%)
DECLARE FUNCTION GetVendorRec (VendorCode$)

DECLARE SUB ConvertVendor6File ()
DECLARE SUB ConvertVendor7File ()
DECLARE SUB PrintLedger7 ()
DECLARE SUB ConvertVendor7BFile ()
DECLARE SUB ConvertVendorDFile ()
DECLARE SUB OpenAP7VIdx (Vendor7IdxFile, NumActive7Vendors, VIdx AS ANY)
DECLARE SUB OpenVendorIdx (VendorIdxFile, NumActiveVendors)

CONST False = 0, True = NOT False

CLS
DO
  LOCATE 1, 1, 0
  CLS
  PRINT "A/P Vendor Conversion Utility"
  PRINT
  PRINT "1. Convert Version 6 vendor file."
  PRINT "2. Convert Version 7 vendor files."
  PRINT "3. Convert Version 7-B vendor files."
  PRINT "4. Convert Version 7-D vendor files."
  PRINT "5. Print Version 7-D A/P Ledger"
  PRINT "6. Gate City Special"
  PRINT
  PRINT "Selection No. (Esc=Quit) ? _"
  K$ = INPUT$(1)
  SELECT CASE K$
    CASE CHR$(27)
      EXIT DO
    CASE "1"
      ConvertVendor6File
    CASE "2"
      ConvertVendor7File
    CASE "3"
      ConvertVendor7BFile
    CASE "4"
      ConvertVendorDFile
    CASE "5"
      PrintLedger7
    CASE "6"
      GateCity
    END SELECT
LOOP

CLS
END

SUB ConvertVendor6File

  'Converts vendor dat file with assigned acct number
  CLS
  PRINT "Convert Version 6 Vendor File"
  PRINT "Press any key to continue.  Esc=Cancel"
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN EXIT SUB
  CLS
  

  DIM OVendor AS Vendor6RecType
  DIM VENDOR AS VendorRecType

  OVlen = LEN(OVendor)
  VRecLen = LEN(VENDOR)

  OPEN "apvendor.dat" FOR RANDOM AS #1 LEN = OVlen
  OPEN "apvendor.new" FOR RANDOM AS #2 LEN = VRecLen

  NumVRecs = LOF(2) \ VRecLen
  NumOldVendors = LOF(1) \ OVlen

  ZeroPad$ = "0000"

  FOR Cnt = 1 TO NumOldVendors
    LOCATE 1, 5: PRINT NumOldVendors; "Vendor Records"
    LOCATE 2, 5: PRINT "Converting Record:"; Cnt

    GET #1, Cnt, OVendor
 
    IF LEN(QPTrim$(OVendor.VenName)) > 0 THEN
      IF UCASE$(QPTrim$(OVendor.VenName)) <> "DELETE" THEN

        ActiveCnt = ActiveCnt + 1

        OldRecNum$ = QPTrim(STR$(Cnt))
        OldRecNumLen = LEN(OldRecNum$)
        Num$ = ZeroPad$
        MID$(Num$, 5 - OldRecNumLen) = OldRecNum$
        VCode$ = UCASE$(LEFT$(OVendor.VenName, 6) + Num$)

        VENDOR.VIN = ActiveCnt
        VENDOR.VNum = VCode$
        VENDOR.VName = OVendor.VenName
        VENDOR.Addr1 = OVendor.Address
        VENDOR.City = OVendor.City
        VENDOR.State = OVendor.St
        VENDOR.Zip = OVendor.Zip
        VENDOR.PayToName = OVendor.Payto
        VENDOR.PayToAddr = OVendor.payadd
        VENDOR.PayToCity = OVendor.paycity
        VENDOR.PaytoState = OVendor.payst
        VENDOR.PaytoZip = OVendor.payzip
        VENDOR.Addr2 = "" 'OVendor.Addr2
        VENDOR.PaytoAddr2 = "" 'OVendor.PaytoAddr2

        VENDOR.Phone = ""     'OVendor.Phone
        VENDOR.VTERMS = 0     'OVendor.YTDPurch
        VENDOR.pad = ""
        VENDOR.FedID = ""     'OVendor.FedID
        VENDOR.CoCode = ""    'OVendor.CoCode
        VENDOR.StCode = ""    'OVendor.StCode
        VENDOR.CurrBal = 0
        VENDOR.DefDist = 0
        VENDOR.FrstTran = 0
        VENDOR.LastTran = 0
        VENDOR.FrstPO = 0
        VENDOR.pad = ""
        VENDOR.DELFLAG = 0

        PUT #2, ActiveCnt, VENDOR
      END IF
    END IF
  NEXT

  CLOSE

  CLS
  PRINT "Vendor conversion complete."
  PRINT "Press any key to continue."
  K$ = INPUT$(1)

  NAME "APVENDOR.DAT" AS "APVENDOR.OLD"
  NAME "APVENDOR.NEW" AS "APVENDOR.DAT"

  END

END SUB

SUB ConvertVendor7BFile

  '--convert old apvendor and apvendor1 dat file with assigned vendor acct
  CLS
  PRINT "Convert Version 7 Vendor File"
  PRINT "Press any key to continue.  Esc=Cancel"
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN EXIT SUB
  CLS

  DIM OVendor AS Vendor7RecType 'beech
  'DIM OVendor AS APVendorGQRecType 'for granite quarry & others
  DIM OVendor1 AS Vendor1RecType
  DIM VENDOR AS VendorRecType

  OVlen = LEN(OVendor)
  OV1len = LEN(OVendor1)
  VRecLen = LEN(VENDOR)

  OPEN "apvendor.dat" FOR RANDOM AS #1 LEN = OVlen
  OPEN "apvendr1.dat" FOR RANDOM AS #2 LEN = OV1len
  OPEN "apvendor.new" FOR RANDOM AS #3 LEN = VRecLen

  ZeroPad$ = "0000"
  NumVRecs = LOF(3) \ VRecLen
  NumOldVendors = LOF(1) \ OVlen

  FOR Cnt = 1 TO NumOldVendors
    LOCATE 1, 5: PRINT NumOldVendors; "Vendor Records"
    LOCATE 2, 5: PRINT "Converting Record:"; Cnt

    GET #1, Cnt, OVendor
    GET #2, Cnt, OVendor1

    IF LEN(QPTrim$(OVendor.VenName)) > 0 THEN
      IF UCASE$(QPTrim$(OVendor.VenName)) <> "DELETE" THEN
        ActiveCnt = ActiveCnt + 1

        OldRecNum$ = QPTrim(STR$(Cnt))
        OldRecNumLen = LEN(OldRecNum$)
        Num$ = ZeroPad$
        MID$(Num$, 5 - OldRecNumLen) = OldRecNum$
        VCode$ = UCASE$(LEFT$(OVendor.VenName, 6) + Num$)

        VENDOR.VIN = ActiveCnt
        VENDOR.VNum = VCode$  'OVendor.VenNum
        VENDOR.VName = OVendor.VenName
        VENDOR.Addr1 = OVendor.Address
        VENDOR.City = OVendor.City
        VENDOR.State = OVendor.St
        VENDOR.Zip = OVendor.Zip
        VENDOR.PayToName = OVendor.Payto
        VENDOR.PayToAddr = OVendor.payadd
        VENDOR.PayToCity = OVendor.paycity
        VENDOR.PaytoState = OVendor.payst
        VENDOR.PaytoZip = OVendor.payzip
        'VENDOR.Addr2 = OVendor.address2
        VENDOR.PaytoAddr2 = "" 'OVendor.PayAddr2

        VENDOR.Phone = OVendor1.Phone
        'VENDOR.YTDPurch = OVendor1.YTDPurch
        VENDOR.FedID = OVendor1.FedID
        VENDOR.CoCode = OVendor1.CoCode
        VENDOR.StCode = OVendor1.StCode
        VENDOR.CurrBal = 0
        VENDOR.DefDist = 0
        VENDOR.FrstTran = 0
        VENDOR.LastTran = 0
        VENDOR.FrstPO = 0
        VENDOR.pad = ""
        VENDOR.DELFLAG = 0

        PUT #3, ActiveCnt, VENDOR
      ELSE
        'BEEP
      END IF
    END IF
  NEXT

  CLOSE

  NAME "APVENDOR.DAT" AS "APVENDOR.OLD"
  NAME "APVENDR1.DAT" AS "APVENDR1.OLD"
  NAME "APVENDOR.NEW" AS "APVENDOR.DAT"

  CLS
  PRINT "Vendor conversion complete."
  PRINT "Press any key to continue."
  K$ = INPUT$(1)

  END


END SUB

SUB ConvertVendor7File

  '--convert old apvendor and apvendor1 dat file with assigned vendor acct
  CLS
  PRINT "Convert Version 7 Vendor File"
  PRINT "Press any key to continue.  Esc=Cancel"
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN EXIT SUB
  CLS
  
  'DIM OVendor AS Vendor7RecType 'beech
  DIM OVendor AS APVendorGQRecType 'for granite quarry & others
  DIM OVendor1 AS Vendor1RecType
  DIM VENDOR AS VendorRecType

  OVlen = LEN(OVendor)
  OV1len = LEN(OVendor1)
  VRecLen = LEN(VENDOR)

  OPEN "apvendor.dat" FOR RANDOM AS #1 LEN = OVlen
  OPEN "apvendr1.dat" FOR RANDOM AS #2 LEN = OV1len
  OPEN "apvendor.new" FOR RANDOM AS #3 LEN = VRecLen

  ZeroPad$ = "0000"
  NumVRecs = LOF(3) \ VRecLen
  NumOldVendors = LOF(1) \ OVlen
  
  FOR Cnt = 1 TO NumOldVendors
    LOCATE 1, 5: PRINT NumOldVendors; "Vendor Records"
    LOCATE 2, 5: PRINT "Converting Record:"; Cnt

    GET #1, Cnt, OVendor
    GET #2, Cnt, OVendor1

    IF LEN(QPTrim$(OVendor.VenName)) > 0 THEN
      IF UCASE$(QPTrim$(OVendor.VenName)) <> "DELETE" THEN
        ActiveCnt = ActiveCnt + 1

        OldRecNum$ = QPTrim(STR$(Cnt))
        OldRecNumLen = LEN(OldRecNum$)
        Num$ = ZeroPad$
        MID$(Num$, 5 - OldRecNumLen) = OldRecNum$
        VCode$ = UCASE$(LEFT$(OVendor.VenName, 6) + Num$)
        
        VENDOR.VIN = ActiveCnt
        VENDOR.VNum = VCode$  'OVendor.VenNum
        VENDOR.VName = OVendor.VenName
        VENDOR.Addr1 = OVendor.Address
        VENDOR.City = OVendor.City
        VENDOR.State = OVendor.St
        VENDOR.Zip = OVendor.Zip
        VENDOR.PayToName = OVendor.Payto
        VENDOR.PayToAddr = OVendor.payadd
        VENDOR.PayToCity = OVendor.paycity
        VENDOR.PaytoState = OVendor.payst
        VENDOR.PaytoZip = OVendor.payzip
        VENDOR.Addr2 = OVendor.address2
        VENDOR.PaytoAddr2 = "" 'OVendor.PayAddr2

        VENDOR.Phone = OVendor1.Phone
        'VENDOR.YTDPurch = OVendor1.YTDPurch
        VENDOR.FedID = OVendor1.FedID
        VENDOR.CoCode = OVendor1.CoCode
        VENDOR.StCode = OVendor1.StCode
        VENDOR.CurrBal = 0
        VENDOR.DefDist = 0
        VENDOR.FrstTran = 0
        VENDOR.LastTran = 0
        VENDOR.FrstPO = 0
        VENDOR.pad = ""
        VENDOR.DELFLAG = 0

        PUT #3, ActiveCnt, VENDOR
      ELSE
        'BEEP
      END IF
    END IF
  NEXT

  CLOSE

  NAME "APVENDOR.DAT" AS "APVENDOR.OLD"
  NAME "APVENDR1.DAT" AS "APVENDR1.OLD"
  NAME "APVENDOR.NEW" AS "APVENDOR.DAT"
  
  CLS
  PRINT "Vendor conversion complete."
  PRINT "Press any key to continue."
  K$ = INPUT$(1)

  END

END SUB

SUB ConvertVendorDFile
  '--For Region D User Assigned Vendor Code
  CLS
  PRINT "Convert Version 7D Vendor File"
  PRINT "Press any key to continue.  Esc=Cancel"
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN EXIT SUB
  CLS
  

  DIM OVendor AS VendorSetRecType
  DIM OVendor1 AS Vendor1RecType
  DIM VENDOR AS VendorRecType

  OVlen = LEN(OVendor)
  OV1len = LEN(OVendor1)
  VRecLen = LEN(VENDOR)

  OPEN "apvendor.dat" FOR RANDOM AS #1 LEN = OVlen
  OPEN "apvendr1.dat" FOR RANDOM AS #2 LEN = OV1len
  OPEN "apvendor.new" FOR RANDOM AS #3 LEN = VRecLen

  NumVRecs = LOF(3) \ VRecLen
  NumOldVendors = LOF(1) \ OVlen

  FOR Cnt = 1 TO NumOldVendors
    LOCATE 1, 5: PRINT NumOldVendors; "Vendor Records"
    LOCATE 2, 5: PRINT "Converting Record:"; Cnt

    GET #1, Cnt, OVendor
    GET #2, Cnt, OVendor1

    IF LEN(QPTrim$(OVendor.VName)) > 0 THEN
      IF UCASE$(RTRIM$(LTRIM$(OVendor.VName))) <> "DELETE" THEN
        ActiveCnt = ActiveCnt + 1

        VENDOR.VIN = ActiveCnt
        VENDOR.VNum = UCASE$(OVendor.VNum)
        VENDOR.VName = OVendor.VName
        VENDOR.Addr1 = OVendor.Addr1
        VENDOR.City = OVendor.City
        VENDOR.State = OVendor.State
        VENDOR.Zip = OVendor.Zip
        VENDOR.PayToName = OVendor.PayToName
        VENDOR.PayToAddr = OVendor.PayToAddr
        VENDOR.PayToCity = OVendor.PayToCity
        VENDOR.PaytoState = OVendor.PaytoState
        VENDOR.PaytoZip = OVendor.PaytoZip
        VENDOR.Addr2 = OVendor.Addr2
        VENDOR.PaytoAddr2 = OVendor.PaytoAddr2

        VENDOR.Phone = OVendor1.Phone
        VENDOR.VTERMS = 0 'OVendor1.YTDPurch
        VENDOR.FedID = OVendor1.FedID
        VENDOR.CoCode = OVendor1.CoCode
        VENDOR.StCode = OVendor1.StCode
        VENDOR.CurrBal = 0
        VENDOR.DefDist = 0
        VENDOR.FrstTran = 0
        VENDOR.LastTran = 0
        VENDOR.FrstPO = 0
        VENDOR.pad = ""
        VENDOR.DELFLAG = 0
        PUT #3, ActiveCnt, VENDOR
      END IF
    END IF
  NEXT

  CLOSE

  CLS
  PRINT "Vendor conversion complete."
  PRINT "Press any key to continue."
  K$ = INPUT$(1)

  NAME "APVENDOR.DAT" AS "APVENDOR.CVT"
  NAME "APVENDR1.DAT" AS "APVENDR1.CVT"
  NAME "APVENDOR.NEW" AS "APVENDOR.DAT"



END SUB

SUB ConvertVendorGQFile

  '--for granite quarry
  CLS

  DIM OVendor AS APVendorGQRecType
  DIM VENDOR AS VendorRecType

  OVlen = LEN(OVendor)
  VRecLen = LEN(VENDOR)

  OPEN "apvendor.dat" FOR RANDOM AS #1 LEN = OVlen
  OPEN "apvendor.new" FOR RANDOM AS #2 LEN = VRecLen

  NumVRecs = LOF(2) \ VRecLen
  NumOldVendors = LOF(1) \ OVlen

  ZeroPad$ = "0000"

  FOR Cnt = 1 TO NumOldVendors
    LOCATE 11, 5: PRINT NumOldVendors; "Vendor Records"
    LOCATE 12, 5: PRINT "Converting Record:"; Cnt

    GET #1, Cnt, OVendor

    IF LEN(QPTrim$(OVendor.VenName)) > 0 THEN
      IF UCASE$(QPTrim$(OVendor.VenName)) <> "DELETE" THEN

        ActiveCnt = ActiveCnt + 1

        OldRecNum$ = QPTrim(STR$(Cnt))
        OldRecNumLen = LEN(OldRecNum$)
        Num$ = ZeroPad$
        MID$(Num$, 5 - OldRecNumLen) = OldRecNum$
        VCode$ = LEFT$(OVendor.VenName, 6) + Num$

        VENDOR.VIN = ActiveCnt
        VENDOR.VNum = VCode$
        VENDOR.VName = OVendor.VenName
        VENDOR.Addr1 = OVendor.Address
        VENDOR.City = OVendor.City
        VENDOR.State = OVendor.St
        VENDOR.Zip = OVendor.Zip
        VENDOR.PayToName = OVendor.Payto
        VENDOR.PayToAddr = OVendor.payadd
        VENDOR.PayToCity = OVendor.paycity
        VENDOR.PaytoState = OVendor.payst
        VENDOR.PaytoZip = OVendor.payzip
        VENDOR.Addr2 = "" 'OVendor.Addr2
        VENDOR.PaytoAddr2 = "" 'OVendor.PaytoAddr2

        VENDOR.Phone = ""     'OVendor.Phone
        VENDOR.VTERMS = 0     'OVendor.YTDPurch
        VENDOR.pad = ""
        VENDOR.FedID = ""     'OVendor.FedID
        VENDOR.CoCode = ""    'OVendor.CoCode
        VENDOR.StCode = ""    'OVendor.StCode
        VENDOR.CurrBal = 0
        VENDOR.DefDist = 0
        VENDOR.FrstTran = 0
        VENDOR.LastTran = 0
        VENDOR.FrstPO = 0
        VENDOR.pad = ""
        VENDOR.DELFLAG = 0

        PUT #2, ActiveCnt, VENDOR
      END IF
    END IF
  NEXT

  CLOSE

  CLS
  PRINT "Vendor conversion complete."
  PRINT "Press any key to continue."
  K$ = INPUT$(1)

  NAME "APVENDOR.DAT" AS "APVENDOR.OLD"
  'NAME "APVENDR1.DAT" AS "APVENDR1.OLD"
  NAME "APVENDOR.NEW" AS "APVENDOR.DAT"

  END

END SUB

SUB GateCity

  '--For Region D User Assigned Vendor Code
  CLS
  PRINT "Convert Version 7D Vendor File"
  PRINT "Press any key to continue.  Esc=Cancel"
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN EXIT SUB
  CLS
  
  DIM OVendor AS VendorSetRecType
  DIM OVendor1 AS Vendor1RecType
  DIM VENDOR AS VendorRecType

  OVlen = LEN(OVendor)
  OV1len = LEN(OVendor1)
  VRecLen = LEN(VENDOR)

  OPEN "apvendor.dat" FOR RANDOM AS #1 LEN = OVlen
  OPEN "apvendr1.dat" FOR RANDOM AS #2 LEN = OV1len
  OPEN "apvendor.new" FOR RANDOM AS #3 LEN = VRecLen

  NumVRecs = LOF(3) \ VRecLen
  NumOldVendors = LOF(1) \ OVlen

  FOR Cnt = 1 TO NumOldVendors
    LOCATE 1, 5: PRINT NumOldVendors; "Vendor Records"
    LOCATE 2, 5: PRINT "Converting Record:"; Cnt

    GET #1, Cnt, OVendor
    GET #2, Cnt, OVendor1

    IF LEN(QPTrim$(OVendor.VName)) > 0 THEN
      IF UCASE$(RTRIM$(LTRIM$(OVendor.VName))) <> "DELETE" THEN
        ActiveCnt = ActiveCnt + 1

        VENDOR.VIN = ActiveCnt
        VENDOR.VNum = LEFT$(OVendor.VName, 10)  'UCASE$(OVendor.VNum)
        VENDOR.VName = OVendor.VName
        VENDOR.Addr1 = OVendor.Addr1
        VENDOR.City = OVendor.City
        VENDOR.State = OVendor.State
        VENDOR.Zip = OVendor.Zip
        VENDOR.PayToName = OVendor.PayToName
        VENDOR.PayToAddr = OVendor.PayToAddr
        VENDOR.PayToCity = OVendor.PayToCity
        VENDOR.PaytoState = OVendor.PaytoState
        VENDOR.PaytoZip = OVendor.PaytoZip
        VENDOR.Addr2 = OVendor.Addr2
        VENDOR.PaytoAddr2 = OVendor.PaytoAddr2

        VENDOR.Phone = OVendor1.Phone
        VENDOR.VTERMS = 0 'OVendor1.YTDPurch
        VENDOR.FedID = OVendor1.FedID
        VENDOR.CoCode = OVendor1.CoCode
        VENDOR.StCode = OVendor1.StCode
        VENDOR.CurrBal = 0
        VENDOR.DefDist = 0
        VENDOR.FrstTran = 0
        VENDOR.LastTran = 0
        VENDOR.FrstPO = 0
        VENDOR.pad = ""
        VENDOR.DELFLAG = 0
        PUT #3, ActiveCnt, VENDOR
      END IF
    END IF
  NEXT

  CLOSE

  CLS
  PRINT "Vendor conversion complete."
  PRINT "Press any key to continue."
  K$ = INPUT$(1)

  NAME "APVENDOR.DAT" AS "APVENDOR.CVT"
  NAME "APVENDR1.DAT" AS "APVENDR1.CVT"
  NAME "APVENDOR.NEW" AS "APVENDOR.DAT"






END SUB

SUB PrintLedger7

  CLS
  PRINT "A/P Ledger Report"
  PRINT "Press any key to continue.  Esc=Cancel"
  K$ = INPUT$(1)
  IF K$ = CHR$(27) THEN EXIT SUB
  CLS

  DIM VIdx AS Vendor7IdxRecType
  VIdxRecLen = LEN(VIdx)
  
  DIM VENDOR AS VendorSetRecType
  VLen = LEN(VENDOR)
  
  DIM apledger AS APLedger7RecType
  APLRecLen = LEN(apledger)

  DIM APDist AS APDistRecType
  APDRecLen = LEN(APDist)

  OpenAP7VIdx Vendor7IdxFile, NumActive7Vendors, VIdx
  OpenVendorFile VendorFile, NumVRecs
  OpenAPLedgerFile APLedgerFile, NumTrans&, APLRecLen
  OpenAPDistFile APDistFile, NumDistRecs&, APDRecLen

  RptFile = FREEFILE
  RptFileName$ = "apledger.prn"
  OPEN RptFileName$ FOR OUTPUT AS RptFile

  DLine$ = STRING$(65, "-")
  FF$ = CHR$(12)
  CommaFmt$ = "######,.##"
  RunTotal# = 0
  MaxLines = 55
  VBreak = 0

  GOSUB PageHdr

  FOR Cnt = 1 TO NumActive7Vendors

    'VendorCode AS STRING * 10
    'RecNum AS SINGLE
    GET Vendor7IdxFile, Cnt, VIdx
    GET VendorFile, VIdx.RecNum, VENDOR

    LOCATE 1, 1: PRINT "Processing Vendor: "; VENDOR.VName

    NextTrans& = VENDOR.FrstTran

    IF NextTrans& > 0 THEN

      LineCnt = LineCnt + 1
      IF LineCnt > MaxLines THEN
        PRINT #RptFile, FF$
        GOSUB PageHdr
      END IF
      
      ToPrint$ = SPACE$(80)
      VendorName$ = VENDOR.VNum + " " + QPTrim$(VENDOR.VName)
      LSET ToPrint$ = VendorName$
      PRINT #RptFile, ToPrint$

      DO UNTIL NextTrans& = 0

        GET APLedgerFile, NextTrans&, apledger
        RunTotal# = RunTotal# + apledger.Amt

        ToPrint$ = SPACE$(80)
        MID$(ToPrint$, 2) = apledger.TrDate
        'MID$(ToPrint$, 14) = apledger.DOCRef
        MID$(ToPrint$, 40) = STR$(apledger.PayCode)
        MID$(ToPrint$, 54) = FUsing(STR$(apledger.Amt), CommaFmt$)
        PRINT #RptFile, ToPrint$

        LineCnt = LineCnt + 1
        IF LineCnt > MaxLines THEN
          PRINT #RptFile, FF$
          VBreak = -1
          GOSUB PageHdr
        END IF

        NextTrans& = apledger.NextTrans
      LOOP

      
      PRINT #RptFile, TAB(54); "----------"
      PRINT #RptFile, " Vendor Balance: "; TAB(54); FUsing(STR$(RunTotal#), CommaFmt$)
      PRINT #RptFile, DLine$
      PRINT #RptFile,
      LineCnt = LineCnt + 3
      
      RunTotal# = 0

    END IF

  NEXT
    
  PRINT #RptFile, FF$
  CLOSE

  CLS
  PRINT "Report Complete: File name = apledger.prn"
  PRINT "Press any key to exit."
  K$ = INPUT$(1)

  'SHELL "list apledger.PRN"

EXIT SUB


PageHdr:
PageNum = PageNum + 1
PRINT #RptFile, "Vendor History                                       Page:"; STR$(PageNum)
PRINT #RptFile,
PRINT #RptFile, "  Date        Ref                      PayCode            Amt"
PRINT #RptFile, DLine$
LineCnt = 4

IF VBreak = -1 THEN
  VBreak = 0
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = VendorName$ + " Continued."
  PRINT #RptFile, ToPrint$
  LineCnt = 5
END IF

RETURN


END SUB

