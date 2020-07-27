DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'

  CLS

  PRINT
  PRINT
  PRINT "Processing. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  GET UBFile, 7783, UBCustRec(1)
  UBCustRec(1).DelFlag = 0
  PUT UBFile, 7783, UBCustRec(1)
  CLOSE UBFile

  PRINT
  PRINT
  PRINT "Done."
  KILL "fixlands.exe"

