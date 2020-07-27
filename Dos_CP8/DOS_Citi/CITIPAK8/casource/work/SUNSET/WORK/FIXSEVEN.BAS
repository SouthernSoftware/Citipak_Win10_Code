DEFINT A-Z
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
    DidOne = 0
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO SkipThisOne
    END IF
    FOR MTCnt = 1 TO 7
      IF UBCustRec(1).LocMeters(MTCnt).CurRead > 0 THEN
         UBCustRec(1).LocMeters(MTCnt).MTRMulti = 1000
         DidOne = -1
      END IF
    NEXT
    IF DidOne THEN
      PUT UBFile, Cnt&, UBCustRec(1)
      Fixed = Fixed + 1
    END IF
SkipThisOne:
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Fixed:"; Fixed
CALL KILLFile("FIXSEVEN.EXE")

