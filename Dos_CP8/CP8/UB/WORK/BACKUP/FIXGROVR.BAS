DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT "Processing accounts. . ."
  PRINT
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  GET UBFile, 808, UBCustRec(1)
  UBCustRec(1).DepositAmt = 75
  PUT UBFile, 808, UBCustRec(1)
  GET UBFile, 833, UBCustRec(1)
  UBCustRec(1).DepositAmt = 75
  PUT UBFile, 833, UBCustRec(1)

  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Processing complete."
  CALL KillFile("fixgrovr.exe")

