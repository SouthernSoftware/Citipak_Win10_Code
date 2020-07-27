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
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    UBCustRec(1).USERCODE2 = "4"
    PUT UBFile, Cnt&, UBCustRec(1)
SkipThisOne:
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Done."
KillFile "FIXCALD.EXE"

