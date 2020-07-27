DEFINT A-Z
  
  '$INCLUDE: 'newcust.BI'

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))


  CLS
  PRINT
  PRINT
  PRINT "Fixing Account 1956. . ."

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  GET UBFile, 1956, UBCustRec(1)
  UBCustRec(1).DepositAmt = 0
  PUT UBFile, 1956, UBCustRec(1)

  CLOSE

PRINT
PRINT
PRINT
PRINT "Done."

