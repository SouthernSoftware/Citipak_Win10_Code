DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Clearing 2567"
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  GET UBFile, 2567, UBCustRec(1)
  LSET UBCustRec(1) = UBCustRec(2)
  UBCustRec(1).DelFlag = -1
  PUT UBFile, 2567, UBCustRec(1)

  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done"

