DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
DECLARE SUB KillFile (FileName$)

  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  Date1 = Date2Num%("08-15-2003")
  
  CLS
  LOCATE 3, 1, 0
  PRINT "Removing Transactions. . ."

  REDIM UBCust(1)    AS NewUBCustRecType
  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))
  CustLen = LEN(UBCust(1))

  KillFile "UBTRANS.bak"
  NAME "UBTRANS.DAT" AS "UBTRANS.bak"

  UBFile1 = FREEFILE
  OPEN "UBTRANS.bak" FOR RANDOM SHARED AS UBFile1 LEN = UBTranRecLen

  UBFile2 = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile2 LEN = UBTranRecLen

  UBFile3 = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile3 LEN = CustLen

  TNumOfRecs& = LOF(UBFile1) / UBTranRecLen

  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; " of"; TNumOfRecs&;
    GET UBFile1, Cnt&, UBTranRec(1)
    IF (UBTranRec(1).TransDate = Date1) THEN 'AND (UBTranRec(1).OperatorNumber = 4) THEN
      SELECT CASE UBTranRec(1).TransType
      CASE TranPenaltyCharge
        GET UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
        FOR RCnt = 1 TO 10
          UBCust(1).CurrRevAmts(RCnt) = uRound(UBCust(1).CurrRevAmts(RCnt) - UBTranRec(1).RevAmt(RCnt))
        NEXT
        UBCust(1).CurrBalance = uRound#(UBCust(1).CurrBalance - UBTranRec(1).TransAmt)
        PUT UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
        Removed = Removed + 1
      CASE ELSE
        PUT UBFile2, , UBTranRec(1)
      END SELECT
    ELSE
      PUT UBFile2, , UBTranRec(1)
    END IF
  NEXT

  CLOSE

LOCATE 9, 1
'PRINT " Copied:"; TTCnt&;
PRINT "Removed:"; Removed
END
KILL "UBTRANS.bak"
KILL "striptrn.exe"

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

