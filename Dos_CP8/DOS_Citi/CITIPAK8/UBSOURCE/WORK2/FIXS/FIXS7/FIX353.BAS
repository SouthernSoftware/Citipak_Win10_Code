DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Correcting Deposit Amount. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  GET UBFile, 353, UBCustRec(1)
  UBCustRec(1).DepositAmt = 100
  PUT UBFile, 353, UBCustRec(1)

  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done"

