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
  GET UBFile, 3, UBCustRec(1)
    FOR revcnt = 1 TO 15
      UBCustRec(1).CurrRevAmts(revcnt) = 0
      'UBCustRec(1).CurrAmts(revcnt) = 0
    NEXT
    UBCustRec(1).CurrBalance = 0
      UBCustRec(1).PrevBalance = 0
    PUT UBFile, 3, UBCustRec(1)
SkipThisOne:
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Done."
KillFile "FIXDMCUS.EXE"

