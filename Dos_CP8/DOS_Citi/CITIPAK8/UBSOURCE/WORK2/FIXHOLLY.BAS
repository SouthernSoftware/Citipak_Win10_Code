DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
  
  '$INCLUDE: 'newcust.BI'

  CLS
  PRINT
  PRINT
  PRINT "Clearing Accounts 2354 & 7323"
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  UBCustRec(1).DelFlag = -1
  PUT UBFile, 2354, UBCustRec(1)
  PUT UBFile, 7323, UBCustRec(1)

  CLOSE UBFile

PRINT
PRINT
PRINT
PRINT "Done"
KILL "FIXHOLLY.EXE"

