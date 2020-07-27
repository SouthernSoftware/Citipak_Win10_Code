DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Removing Account 2279"
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  GET UBFile, 2279, UBCustRec(1)
  UBCustRec(1).DelFlag = -1
  UBCustRec(1).DepositAmt = 0
  PUT UBFile, 2279, UBCustRec(1)
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done"

