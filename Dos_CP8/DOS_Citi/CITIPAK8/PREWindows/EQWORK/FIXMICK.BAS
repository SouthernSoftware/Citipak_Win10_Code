DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)

  '$INCLUDE: 'newcust.BI'

  CLS

  BDate = Date2Num("12-31-1979")
  GDate = Date2Num("08-14-2000")

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
    PRINT "Processing Customer:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    DidEm = 0
    FOR MCnt = 1 TO 7
      IF UBCustRec(1).LocMeters(MCnt).CurDate = BDate THEN
        UBCustRec(1).LocMeters(MCnt).CurDate = GDate
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
  KILL "FIXMICK.EXE"


