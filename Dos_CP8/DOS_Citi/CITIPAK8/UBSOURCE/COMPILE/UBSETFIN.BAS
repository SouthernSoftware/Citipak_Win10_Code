DEFINT A-Z
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE SUB WaitForAction ()
DECLARE SUB BlockClear ()
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FClose (Handle%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB UBLog (Text$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB ReIndexSystem (PromptFlag%)

  '$INCLUDE: 'defcnf.bi'
  '$INCLUDE: 'formedit.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubowner.BI'

  CONST False = 0, True = NOT False

  CrLf$ = CHR$(13) + CHR$(10)

SUB DeleteCust

  UBLog " IN: DELETE Account."

  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  DO
    BlockClear
    LookUp RecNo&, "Customer to Delete", 4, True, False
    IF RecNo& > 0 THEN
      GOSUB CheckCust
      IF CustOK THEN
        EXIT DO
      END IF
    ELSE
      EXIT DO
    END IF
  LOOP
  
  IF RecNo& > 0 THEN
    DisplayUBScrn "DELACUST"
    DO
      GOSUB DDisplayCustInfo
      Ok = MsgBox%("UB", "DODELBTN")
      SELECT CASE Ok
      CASE 1
        OKFlag = False
        ExitFlag = True
      CASE 2
        ExitFlag = True
        Ok = MsgBox%("UB", "OK2DEL")
        SELECT CASE Ok
        CASE 1
          OKFlag = False
        CASE 2
          OKFlag = True
        END SELECT
      CASE 3
        ShowCustHistory RecNo&
      END SELECT
    LOOP UNTIL ExitFlag
  END IF
  
  IF NOT OKFlag THEN
    GOTO ExitDeleteCust
  END IF
  
  GOSUB DeleteCustomer
  ReIndexSystem False
  '******************************
  
ExitDeleteCust:
  UBLog "OUT: DELETE Account." + CrLf$
  EXIT SUB
  
CheckCust:
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, RecNo&, UBCustRec(1)
  CLOSE UBCust
  TotalBalance# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
  IF TotalBalance# <> 0 THEN
    UBLog "NODELETE:" + STR$(RecNo&) + " BAL:" + STR$(TotalBalance#)
    WhatError = 1
    GOSUB NODeleteError
    CustOK = False
  ELSEIF UBCustRec(1).DepositAmt <> 0 THEN
    UBLog "NODELETE:" + STR$(RecNo&) + " DEP:" + STR$(UBCustRec(1).DepositAmt)
    WhatError = 2
    GOSUB NODeleteError
    CustOK = False
  ELSEIF UBCustRec(1).Status <> "I" THEN
    UBLog "NODELETE:" + STR$(RecNo&) + " NOT INACTIVE"
    WhatError = 3
    GOSUB NODeleteError
    'DisplayUBScrn "ERRSCRN1"
    CustOK = False
  ELSE
    CustOK = True
  END IF
  
RETURN
  
DDisplayCustInfo:
  Zip$ = UBCustRec(1).ZIPCODE
  Zip$ = LEFT$(Zip$, 5) + "-" + MID$(Zip$, 6)
  QPrintRC UBCustRec(1).BOOK, 5, 26, -1
  QPrintRC UBCustRec(1).SEQNUMB, 5, 29, -1
  QPrintRC UBCustRec(1).Status, 5, 47, -1
  QPrintRC Num2Date(UBCustRec(1).OPENDATE), 5, 63, -1
  QPrintRC UBCustRec(1).SEARCH, 7, 26, -1
  QPrintRC STR$(RecNo&), 7, 62, -1
  QPrintRC UBCustRec(1).CUSTNAME, 8, 26, -1
  QPrintRC UBCustRec(1).ADDR1, 9, 26, -1
  QPrintRC UBCustRec(1).ADDR2, 10, 26, -1
  QPrintRC UBCustRec(1).SERVADDR, 11, 26, -1
  QPrintRC UBCustRec(1).CITY, 12, 26, -1
  QPrintRC UBCustRec(1).STATE, 12, 54, -1
  QPrintRC Zip$, 12, 63, -1
RETURN
  
DeleteCustomer:
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, RecNo&, UBCustRec(1)
  TN$ = QPTrim$(UBCustRec(1).CUSTNAME)
  IF LEN(TN$) = 0 THEN
    TN$ = "NO NAME"
  END IF
    
  UBCustRec(1).Status = ""
  UBCustRec(1).BOOK = ""
  UBCustRec(1).SEQNUMB = ""
  UBCustRec(1).DelFlag = True
  PUT UBCust, RecNo&, UBCustRec(1)
  CLOSE UBCust
  UBLog "DELETED ACCOUNT:" + STR$(RecNo&) + "  " + TN$
RETURN

NODeleteError:
  CursorOff
  DisplayUBScrn "ERRSCRN1"
  SELECT CASE WhatError
  CASE 1 'has balance
    QPrintRC "This account HAS A BALANCE", 10, 27, -1
    QPrintRC "CAN NOT DELETE THIS ACCOUNT!", 12, 26, -1
  CASE 2 'has deposit
    QPrintRC "This account HAS A DEPOSIT", 10, 27, -1
    QPrintRC "CAN NOT DELETE THIS ACCOUNT!", 12, 26, -1
  CASE 3 'not inactive
    QPrintRC "This account IS NOT INACTIVE", 10, 26, -1
    QPrintRC "CAN NOT DELETE THIS ACCOUNT!", 12, 26, -1
  END SELECT
  WaitForAction
  WhatError = 0

RETURN
END SUB

SUB SetCustFinal
  
  UBLog " IN: Set Customer to Final."
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBOwnerRec(1) AS UBOwnerRecType
  UBOwnerRecLen = LEN(UBOwnerRec(1))
  
  BlockClear
  LookUp RecNo&, "Customer to Final", 2, True, True
  
  IF RecNo& > 0 THEN
    GOSUB GetCustInfo
    DisplayUBScrn "FINACUST"
    GOSUB DisplayCustInfo
    Ok = MsgBox%("UB", "DOFINBTN")
    SELECT CASE Ok
    CASE 1
      UBLog "ABORTED: SET FINAL:" + STR$(RecNo&)
      OKFlag = False
    CASE 2
      OKFlag = True
    END SELECT
  END IF
  
  IF NOT OKFlag THEN
    GOTO ExitSetFinal
  END IF
  
  '******************************
  BlockClear
  DisplayUBScrn "UPDATDSK"
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  GET UBCust, RecNo&, UBCustRec(1)

  LSET UBCustRec(2) = UBCustRec(1)
  
  UBCustRec(2).Status = "I"
  UBCustRec(2).OPENDATE = -32767
  UBCustRec(2).SEARCH = ""
  UBCustRec(2).CUSTNAME = "VACANT"
  UBCustRec(2).ADDR1 = ""
  UBCustRec(2).ADDR2 = ""

  UBCustRec(2).HPHONE = ""
  UBCustRec(2).WPHONE = ""
  UBCustRec(2).SOSEC = ""
  UBCustRec(2).DRVLIC = ""
  'UBCustRec(2).CUSTTYPE = ""

  UBCustRec(2).CustPIN = 0
  UBCustRec(2).LastTrans = 0
  UBCustRec(2).CurrBalance = 0
  UBCustRec(2).PrevBalance = 0
  FOR Cnt = 1 TO 15
    UBCustRec(2).CurrRevAmts(Cnt) = 0
  NEXT

  FOR Cnt = 1 TO 7
    UBCustRec(2).LocMeters(Cnt).AvgUse = 0
    UBCustRec(2).LocMeters(Cnt).UseCnt = 0
  NEXT

  UBCustRec(2).EPPFlag = "N"
  UBCustRec(2).USEDRAFT = ""
  UBCustRec(2).AcctType = ""
  UBCustRec(2).BANKNAME = ""
  UBCustRec(2).BANKLOC = ""
  UBCustRec(2).TRANSIT = ""
  UBCustRec(2).BANKACCT = ""
  UBCustRec(2).DepositAmt = 0
  UBCustRec(2).DelFlag = 0
  UBCustRec(2).PreNoteFlag = 0
  UBCustRec(2).WOLastTrans = 0
  UBCustRec(2).EstFlag = ""
  UBCustRec(2).MessageRec = 0

  NextCustRecs& = (LOF(UBCust) \ UBCustRecLen) + 1
  UBCustRec(1).Status = "F"
  UBCustRec(1).BOOK = ""
  UBCustRec(1).SEQNUMB = ""
  UBCustRec(1).OldRec = NextCustRecs&

  PUT UBCust, RecNo&, UBCustRec(1)
  PUT UBCust, NextCustRecs&, UBCustRec(2)
  CLOSE

'121598 Added coping of owner to new "F" customer record
  UBFile = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBFile LEN = UBOwnerRecLen
  GET UBFile, RecNo&, UBOwnerRec(1)
  PUT UBFile, NextCustRecs&, UBOwnerRec(1)
  CLOSE

  UBLog "SET FINAL:" + STR$(RecNo&)
  UBLog "  CREATED:" + STR$(NextCustRecs&)
  UBLog "SET FINAL:  REINDEXING"

  ReIndexSystem False

ExitSetFinal:
  UBLog "OUT: SET Customer to Final." + CrLf$
EXIT SUB
  
  
GetCustInfo:
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, RecNo&, UBCustRec(1)
  CLOSE UBCust
RETURN
  
DisplayCustInfo:
  Zip$ = UBCustRec(1).ZIPCODE
  Zip$ = LEFT$(Zip$, 5) + "-" + MID$(Zip$, 6)
  QPrintRC UBCustRec(1).BOOK, 5, 26, -1
  QPrintRC UBCustRec(1).SEQNUMB, 5, 29, -1
  QPrintRC UBCustRec(1).Status, 5, 47, -1
  QPrintRC Num2Date(UBCustRec(1).OPENDATE), 5, 63, -1
  QPrintRC UBCustRec(1).SEARCH, 7, 26, -1
  QPrintRC UBCustRec(1).CUSTNAME, 8, 26, -1
  QPrintRC STR$(RecNo&), 7, 63, -1
  QPrintRC UBCustRec(1).ADDR1, 9, 26, -1
  QPrintRC UBCustRec(1).ADDR2, 10, 26, -1
  QPrintRC UBCustRec(1).SERVADDR, 11, 26, -1
  QPrintRC UBCustRec(1).CITY, 12, 26, -1
  QPrintRC UBCustRec(1).STATE, 12, 54, -1
  QPrintRC Zip$, 12, 63, -1
RETURN

END SUB

