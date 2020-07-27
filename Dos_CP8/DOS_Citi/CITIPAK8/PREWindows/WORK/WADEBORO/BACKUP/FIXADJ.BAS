DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION Date2Num% (Today$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION uRound# (DblNum#)
  
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  TDate = Date2Num%("10-14-1997")

  CLS
  LOCATE 3, 1, 0
  PRINT "Processing Transactions History. . ."

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  UBFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile LEN = UBTranRecLen

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  GET UBCust, 1416, UBCustRec(1)
  GET UBFile, UBCustRec(1).LastTrans, UBTranRec(1)
  IF UBTranRec(1).TransType = TranDownwardAdjustment THEN
    IF UBTranRec(1).TransDate = TDate THEN
      UBTranRec(1).TransAmt = 1189.89
      PUT UBFile, UBCustRec(1).LastTrans, UBTranRec(1)
    END IF
  END IF
  CLOSE

LOCATE 9, 1
PRINT "  Fixed: 1"

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

