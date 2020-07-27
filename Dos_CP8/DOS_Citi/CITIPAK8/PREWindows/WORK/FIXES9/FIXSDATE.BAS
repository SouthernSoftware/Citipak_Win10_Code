DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)
DECLARE FUNCTION QPTrim$ (TheText$)

  '$INCLUDE: 'newcust.BI'

  CLS

  BDate1 = Date2Num("06-29-2000")


  PRINT
  PRINT
  PRINT "Processing Read Dates. . ."
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
    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO SkipThisOne
    END IF
    DidEm = 0
    FOR MCnt = 1 TO 7
      Mtr$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)
      IF LEN(Mtr$) > 0 THEN
        UBCustRec(1).LocMeters(MCnt).CurDate = BDate1
        DidEm = -1
      END IF
    NEXT
    IF DidEm THEN
      FixedCnt = FixedCnt + 1
    END IF
    PUT UBFile, Cnt&, UBCustRec(1)
SkipThisOne:
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Fixed:"; FixedCnt
KILL "FIXSDATE.EXE"

