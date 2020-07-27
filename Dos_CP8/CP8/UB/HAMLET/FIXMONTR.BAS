DEFINT A-Z

  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Correcting Accounts. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  GET UBFile, 71, UBCustRec(1)
  UBCustRec(1).CurrRevAmts(1) = 35.11
  UBCustRec(1).CurrRevAmts(2) = 12.75
  PUT UBFile, 71, UBCustRec(1)

  GET UBFile, 341, UBCustRec(1)
  FOR zz = 1 TO 15
    UBCustRec(1).CurrRevAmts(zz) = 0
  NEXT
  UBCustRec(1).PrevBalance = 0
  UBCustRec(1).CurrBalance = 0
  PUT UBFile, 341, UBCustRec(1)

  GET UBFile, 426, UBCustRec(1)
  UBCustRec(1).CurrRevAmts(1) = .91
  UBCustRec(1).CurrRevAmts(2) = 9.56
  PUT UBFile, 426, UBCustRec(1)
  
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Processing Complete."

