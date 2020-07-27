DECLARE SUB SmallPause ()
DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION Num2Date$ (DateNumb%)
DECLARE FUNCTION QPTrim$ (Text$)
  '$INCLUDE: 'newcust.BI'


  CLS
  PRINT
  PRINT
  PRINT "Processing Blank Readings. . ."
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
    DidEm = 0
    IF QPTrim$(UBCustRec(1).LocMeters(2).MTRUnit) = "" THEN
      IF UBCustRec(1).LocMeters(2).MtrType = "C" THEN
        UBCustRec(1).LocMeters(2).MtrType = ""
        DidEm = -1
      END IF
    END IF
 '   IF UBCustRec(1).LocMeters(1).PrevRead < 0 THEN
  '    UBCustRec(1).LocMeters(1).PrevRead = 0
  '    DidEm = -1
  '  END IF

      IF DidEm = -1 THEN
        FixedCnt = FixedCnt + 1
        PUT UBFile, Cnt&, UBCustRec(1)
      END IF
  '  SmallPause
  NEXT
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Fixed:"; FixedCnt
  CALL KillFile("fixrdate.exe")

SUB SmallPause
ST! = TIMER
st2! = ST! + .008
DO
LOOP UNTIL TIMER > st2!
END SUB

