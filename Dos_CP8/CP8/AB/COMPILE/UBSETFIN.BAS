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
'  CursorOff
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

