DECLARE FUNCTION uRound# (N#)
DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
  
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
  GET UBFile, 1321, UBCustRec(1)
  UBCustRec(1).DepositAmt = 60
  PUT UBFile, 1321, UBCustRec(1)
  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Finished."

