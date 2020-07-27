DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
DECLARE SUB KillFile (FileName$)

  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  Date1 = Date2Num%("09-30-2009")
  
  CLS
  LOCATE 3, 1, 0
  PRINT "Removing Transactions. . ."

  REDIM UBCust(1)    AS NewUBCustRecType
  REDIM UBTranRec(1) AS UBTransRecType
  DIM Fixed AS LONG
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

  FOR Cnt& = TNumOfRecs& TO (TNumOfRecs& - 504) STEP -1
    Fixed = Fixed + 1
    LOCATE 5, 1
    PRINT "Correcting Customer Balances"; Fixed; " of 505";
    GET UBFile1, Cnt&, UBTranRec(1)
    IF (UBTranRec(1).TransDate = Date1) THEN 'AND (UBTranRec(1).OperatorNumber = 4) THEN
      SELECT CASE UBTranRec(1).TransType
      CASE TranUtilityBill
        GET UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
        FOR RCnt = 1 TO 15
          UBCust(1).CurrRevAmts(RCnt) = uRound(UBCust(1).CurrRevAmts(RCnt) - UBTranRec(1).RevAmt(RCnt))
        NEXT
        UBCust(1).CurrBalance = uRound#(UBCust(1).CurrBalance - UBTranRec(1).TransAmt)
        PUT UBFile3, UBTranRec(1).CustAcctNo, UBCust(1)
      END SELECT
    END IF
  NEXT

  FOR Cnt& = 1 TO (TNumOfRecs& - 504)
    LOCATE 8, 1
    PRINT "Copying Transactions:"; Cnt&; " of "; TNumOfRecs&;
    GET UBFile1, Cnt&, UBTranRec(1)
    PUT UBFile2, , UBTranRec(1)
  NEXT
 

LOCATE 12, 1
'PRINT " Copied:"; TTCnt&;

PRINT "Corrected:"; Fixed

END
KILL "UBTRANS.bak"
KILL "fixanson.exe"

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

