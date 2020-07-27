DECLARE SUB SmallPause ()
DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION Num2Date$ (DateNumb%)

  '$INCLUDE: 'newcust.BI'

  RDate1 = Date2Num("12-31-2002")
  RDate2 = Date2Num("12-31-2001")

  CLS
  PRINT
  PRINT
  PRINT "Scanning for BAD read date FROM Sensus. . ."
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
    FOR MCnt = 1 TO 7
      IF UBCustRec(1).LocMeters(MCnt).CurDate = RDate1 THEN
        UBCustRec(1).LocMeters(MCnt).CurDate = RDate2
        DidEm = -1
      END IF
    NEXT
    IF DidEm THEN
      FixedCnt = FixedCnt + 1
      PUT UBFile, Cnt&, UBCustRec(1)
    END IF
    SmallPause
  NEXT
  CLOSE UBFile
  PRINT
  PRINT
  PRINT
  PRINT "Fixed:"; FixedCnt
  CALL KillFile("fixdjost.exe")

SUB SmallPause
  st1# = TIMER
  st2# = st1# + .1
  DO
  LOOP WHILE TIMER < st2#

END SUB

