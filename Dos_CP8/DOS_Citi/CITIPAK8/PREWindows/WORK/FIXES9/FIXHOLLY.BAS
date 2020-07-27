DEFINT A-Z
DECLARE SUB SmallPause ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)

  '$INCLUDE: 'newcust.BI'

  CLS

  BDate = Date2Num("06-16-2000")
  GDate = Date2Num("08-16-2000")

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
    SmallPause
    LOCATE , 1
    PRINT "Processing Customer:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    DidEm = 0

    IF UBCustRec(1).DelFlag <> 0 THEN
      GOTO SkipThisOne
    END IF
    SELECT CASE UBCustRec(1).BOOK
    CASE "06"
      FOR MCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(MCnt).CurDate = BDate THEN
          UBCustRec(1).LocMeters(MCnt).CurDate = GDate
          DidEm = -1
        END IF
      NEXT
    END SELECT

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
  KILL "FIXHOLLY.EXE"

SUB SmallPause
st# = TIMER
st# = st# + .2
DO
LOOP UNTIL TIMER > st#
END SUB

