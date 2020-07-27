DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Fixing Invalid Accounts. . ."
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  FOR Cnt = 1094 TO 1098
    GET UBFile, Cnt, UBCustRec(1)
    LSET UBCustRec(1) = UBCustRec(2)
    UBCustRec(1).DelFlag = -1
    PUT UBFile, Cnt, UBCustRec(1)
  NEXT

  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done"

