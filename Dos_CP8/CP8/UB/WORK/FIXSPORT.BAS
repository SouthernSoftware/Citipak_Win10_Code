DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  TranDate = Date2Num("05-29-2002")
  CALL KillFile("NUBTRANS.DAT")

  CLS
  LOCATE 3, 1, 0
  PRINT "Removing Transactions. . ."

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen

  NUBFile = FREEFILE
  OPEN "NUBTRANS.DAT" FOR RANDOM SHARED AS NUBFile LEN = UBTranRecLen
  
  CUBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CUBFile LEN = UBCustRecLen

  TNumOfRecs& = LOF(UBFile) / UBTranRecLen

  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; TNumOfRecs&;
    GET UBFile, Cnt&, UBTranRec(1)
    IF UBTranRec(1).TransType = TranPenaltyCharge THEN
      SELECT CASE UBTranRec(1).TransDate
      CASE TranDate
        GET CUBFile, UBTranRec(1).CustAcctNo, UBCustRec(1)
        FOR RCnt = 1 TO 15
          UBCustRec(1).CurrRevAmts(RCnt) = uRound#(UBCustRec(1).CurrRevAmts(RCnt) - UBTranRec(1).RevAmt(RCnt))
          UBCustRec(1).CurrBalance = uRound#(UBCustRec(1).CurrBalance - UBTranRec(1).RevAmt(RCnt))
        NEXT
        Removed = Removed + 1
        PUT CUBFile, UBTranRec(1).CustAcctNo, UBCustRec(1)
      CASE ELSE
        PUT NUBFile, , UBTranRec(1)
      END SELECT
    ELSE
      PUT NUBFile, , UBTranRec(1)
    END IF
  NEXT
  CLOSE

LOCATE 9, 1
PRINT "Removed:"; Removed

NAME "UBTRANS.DAT" AS "OUBTRANS.OLD"
NAME "NUBTRANS.DAT" AS "UBTRANS.DAT"

CALL KillFile("FIXSPORT.EXE")


'      IF UBTranRec(1).TransAmt > 0 THEN
'        PUT NUBFile, , UBTranRec(1)
'      ELSE
'        Removed = Removed + 1
'      END IF

DEFSNG A-Z
DEFINT A-Z
'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

