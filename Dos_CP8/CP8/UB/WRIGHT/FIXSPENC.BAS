DEFINT A-Z

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
  GET UBFile, 1431, UBCustRec(1)
  FOR MCnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(MCnt) = 0
  NEXT
  PUT UBFile, 1431, UBCustRec(1)
  GET UBFile, 1603, UBCustRec(1)
  FOR MCnt = 1 TO 15
    UBCustRec(1).CurrRevAmts(MCnt) = 0
  NEXT
  PUT UBFile, 1603, UBCustRec(1)
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Processing Complete."

