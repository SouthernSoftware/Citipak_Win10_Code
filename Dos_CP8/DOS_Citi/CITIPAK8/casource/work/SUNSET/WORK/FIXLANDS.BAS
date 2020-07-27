DEFINT A-Z
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)

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
    IF LEN(QPTrim$(UBCustRec(1).AcctType)) > 0 THEN
      UBCustRec(1).USEDRAFT = "Y"
      PUT UBFile, Cnt&, UBCustRec(1)
      Fixed& = Fixed& + 1
    END IF
SkipThisOne:
  NEXT
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Changed:"; Fixed&
  KillFile "FIXLANDS.EXE"

