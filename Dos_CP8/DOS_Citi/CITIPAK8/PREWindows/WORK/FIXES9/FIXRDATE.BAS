DEFINT A-Z
DECLARE SUB SmallPause ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)

  '$INCLUDE: 'newcust.BI'

  BDate1 = Date2Num("11-18-2001")
  BDate2 = Date2Num("01-18-2001")

  CLS

  PRINT
  PRINT "Warning Warning Warning"
  PRINT
  PRINT "ALL UTILITY BILLING OPERATIONS MUST!! STOP UNTIL THIS"
  PRINT "PROGRAM HAS COMPLETED!!!!!!!!"
  PRINT
  PRINT
  PRINT "Press any key when ready...";
  DO WHILE LEN(INKEY$) = 0: LOOP
  CLS
  PRINT
  PRINT

  PRINT "Processing Read Dates. . ."
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  PRINT
  PRINT

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  FOR Cnt& = 1 TO NumOfRecs&
    LOCATE , 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRecs&;
    GET UBFile, Cnt&, UBCustRec(1)
    'IF UBCustRec(1).DelFlag <> 0 THEN
    '  GOTO SkipThisOne
    'END IF

    SELECT CASE UBCustRec(1).Book
    CASE "05"
      DidEm = 0
      FOR MCnt = 1 TO 7
        'IF UBCustRec(1).LocMeters(MCnt).CurRead > 0 OR UBCustRec(1).LocMeters(MCnt).PrevRead > 0 THEN
          ReadDate = UBCustRec(1).LocMeters(MCnt).CurDate
          IF ReadDate = BDate1 THEN
            DidEm = -1
            UBCustRec(1).LocMeters(MCnt).CurDate = BDate2
          END IF
        'END IF
      NEXT
      IF DidEm THEN
        FixedCnt = FixedCnt + 1
        PUT UBFile, Cnt&, UBCustRec(1)
      END IF

    END SELECT
'    SmallPause
SkipThisOne:
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Fixed:"; FixedCnt
CALL KILLFile("FIXRDATE.EXE")

SUB SmallPause
st# = TIMER
st2# = st# + .13
DO
LOOP UNTIL TIMER > st2#
END SUB

