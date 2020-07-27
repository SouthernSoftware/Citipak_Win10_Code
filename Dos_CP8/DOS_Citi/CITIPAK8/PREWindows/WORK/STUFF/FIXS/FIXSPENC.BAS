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
  GET UBFile, 16, UBCustRec(1)
  UBCustRec(1).CurrRevAmts(1) = 14.79
  UBCustRec(1).CurrRevAmts(2) = 10.44
  UBCustRec(1).CurrRevAmts(3) = 8
  PUT UBFile, 16, UBCustRec(1)

  GET UBFile, 193, UBCustRec(1)
  FOR zz = 1 TO 15
    UBCustRec(1).CurrRevAmts(zz) = 0
  NEXT
  UBCustRec(1).CurrBalance = 0
  UBCustRec(1).PrevBalance = 0
  PUT UBFile, 193, UBCustRec(1)

  GET UBFile, 363, UBCustRec(1)
  UBCustRec(1).CurrRevAmts(1) = 353.84
  UBCustRec(1).CurrRevAmts(2) = 198.8
  PUT UBFile, 363, UBCustRec(1)

  GET UBFile, 364, UBCustRec(1)
  UBCustRec(1).CurrRevAmts(1) = 365.72
  PUT UBFile, 364, UBCustRec(1)

  GET UBFile, 365, UBCustRec(1)
  UBCustRec(1).CurrRevAmts(1) = 5
  PUT UBFile, 365, UBCustRec(1)

  GET UBFile, 596, UBCustRec(1)
  UBCustRec(1).CurrRevAmts(1) = 6
  UBCustRec(1).CurrRevAmts(2) = 5.56
  UBCustRec(1).CurrRevAmts(3) = 8
  PUT UBFile, 596, UBCustRec(1)

  GET UBFile, 798, UBCustRec(1)
  UBCustRec(1).CurrRevAmts(1) = 7.16
  UBCustRec(1).CurrRevAmts(2) = 6.2
  UBCustRec(1).CurrRevAmts(3) = 8
  PUT UBFile, 798, UBCustRec(1)

  GET UBFile, 1166, UBCustRec(1)
  UBCustRec(1).CurrRevAmts(1) = 23.43
  UBCustRec(1).CurrRevAmts(2) = 15.24
  UBCustRec(1).CurrRevAmts(3) = 8
  PUT UBFile, 1166, UBCustRec(1)

  GET UBFile, 1407, UBCustRec(1)
  UBCustRec(1).CurrRevAmts(1) = 42.4
  UBCustRec(1).CurrRevAmts(2) = 25.78
  UBCustRec(1).CurrRevAmts(3) = 8
  PUT UBFile, 1407, UBCustRec(1)

  GET UBFile, 1408, UBCustRec(1)
  UBCustRec(1).CurrRevAmts(1) = 20.3
  PUT UBFile, 1408, UBCustRec(1)

  GET UBFile, 1418, UBCustRec(1)
  UBCustRec(1).CurrRevAmts(1) = 10.58
  UBCustRec(1).CurrRevAmts(2) = 8.1
  UBCustRec(1).CurrRevAmts(3) = 8
  PUT UBFile, 1418, UBCustRec(1)

  GET UBFile, 1431, UBCustRec(1)
  FOR zz = 1 TO 15
    UBCustRec(1).CurrRevAmts(zz) = 0
  NEXT
  UBCustRec(1).CurrBalance = 0
  UBCustRec(1).PrevBalance = 0
  PUT UBFile, 1431, UBCustRec(1)

  GET UBFile, 1603, UBCustRec(1)
  FOR zz = 1 TO 15
    UBCustRec(1).CurrRevAmts(zz) = 0
  NEXT
  UBCustRec(1).CurrBalance = 0
  UBCustRec(1).PrevBalance = 0
  PUT UBFile, 1603, UBCustRec(1)


  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Processing Complete."

