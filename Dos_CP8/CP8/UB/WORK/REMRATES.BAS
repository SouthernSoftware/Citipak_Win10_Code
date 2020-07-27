DEFINT A-Z
DECLARE SUB SmallPause ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION QPTrim$ (Text$)
  '$INCLUDE: 'newcust.BI'

  Code1$ = "FR"
  Code2$ = "FC"

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

  PRINT "Removing Ratecodes. . ."
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

      DidEm = 0
      FOR MCnt = 1 TO 15
        ThisRate$ = QPTrim$(UBCustRec(1).Serv(MCnt).RateCode)
        IF ThisRate$ = Code1$ OR ThisRate$ = Code2$ THEN
          UBCustRec(1).Serv(MCnt).RateCode = ""
          DidEm = -1
        END IF
      NEXT
      IF DidEm THEN
        FixedCnt = FixedCnt + 1
        PUT UBFile, Cnt&, UBCustRec(1)
      END IF

'    SmallPause
SkipThisOne:
  NEXT
  CLOSE UBFile
PRINT
PRINT
PRINT
PRINT "Fixed:"; FixedCnt
CALL KILLFile("remrates.exe")

SUB SmallPause
st# = TIMER
st2# = st# + .13
DO
LOOP UNTIL TIMER > st2#
END SUB

