DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.BI'

  CLS
  BDate0 = Date2Num("06-15-2000")
  BDate1 = Date2Num("06-16-2000")


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
    SELECT CASE UBCustRec(1).Book
    CASE "01"
      RDate = BDate0
    CASE "02"
      RDate = BDate1
    CASE ELSE
      GOTO SkipThisOne
    END SELECT

    Didem = 0
    FOR MCnt = 1 TO 7
      MtType$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MtrType)
      IF LEN(MtType$) > 0 THEN
        UBCustRec(1).LocMeters(MCnt).CurDate = RDate
        Didem = -1
      END IF
    NEXT
    IF Didem THEN
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
CALL KILLFile("FIXPENDL.EXE")

