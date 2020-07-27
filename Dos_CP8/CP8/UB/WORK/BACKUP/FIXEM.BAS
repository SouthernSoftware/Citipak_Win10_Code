DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT "Processing. . ."
  PRINT
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  GET UBFile, 2997, UBCustRec(1)
  UBCustRec(1).DelFlag = 0
  PUT UBFile, 2997, UBCustRec(1)

  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Processing complete."
  CALL KillFile("fixem.exe")

