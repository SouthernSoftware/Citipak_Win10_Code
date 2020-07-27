DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
  
  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubtrans.bi'

  CLS
  PRINT
  PRINT
  PRINT "Correcting Balances...."
  REDIM UBTrans(1)        AS UBTransRecType
  UBTranRecLen = LEN(UBTrans(1))
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  TRFile = FREEFILE
  OPEN "UBTRANS.dat" FOR RANDOM SHARED AS TRFile LEN = UBTranRecLen
  
  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfCust& = LOF(UBFile) / UBCustRecLen
  FOR Cnt& = 1 TO NumOfCust&
    GET UBFile, Cnt&, UBCustRec(1)
    LastTran& = UBCustRec(1).LastTrans

  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done"

