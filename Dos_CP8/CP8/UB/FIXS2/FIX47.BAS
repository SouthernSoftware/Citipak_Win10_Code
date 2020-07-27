DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Changing Acct 47 to 'B' Status."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  GET UBFile, 1676, UBCustRec(1)
  'PUT UBFile, 47, UBCustRec(1)
  PRINT UBCustRec(1).PrevBalance, UBCustRec(1).CurrBalance

  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done"

