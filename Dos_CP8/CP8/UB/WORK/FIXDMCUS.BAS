DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)
DECLARE SUB KillFile (FileName$)

  '$INCLUDE: 'newcust.BI'

  CLS

  PRINT
  PRINT
  PRINT "Processing Customers. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  GET UBFile, 301, UBCustRec(1)
    UBCustRec(1).CurrBalance = 0
    UBCustRec(1).PrevBalance = 0
    PUT UBFile, 301, UBCustRec(1)
SkipThisOne:
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Done."
KillFile "FIXDMCUS.EXE"

