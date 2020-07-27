DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)

  GoodDate = Date2Num("04-28-1999")
  BadDate = Date2Num("05-05-1999")


  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CONST False = 0, True = NOT False

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  CLS
  LOCATE 3, 1, 0
  PRINT "Correcting Deposit Information. . ."

  PRINT
  PRINT

  UBCFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBCFile LEN = UBCustRecLen
  CNumOfRecs& = LOF(UBCFile) \ UBCustRecLen
  
  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBTFile) / UBTranRecLen
  FOR Cnt& = 1 TO TNumOfRecs&
    GET UBTFile, Cnt&, UBTranRec(1)
    LOCATE 5, 1: PRINT "Processing:"; Cnt&; " of"; TNumOfRecs&;
    SELECT CASE UBTranRec(1).TransType
    CASE TranRefundDeposit, TranRefundDeposit + 100
      FoundIt = FoundIt + 1
    END SELECT
  NEXT

SkiptoNext:
  CLOSE

LOCATE 9, 1
PRINT "Found:"; FoundIt

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

