DECLARE SUB SmallPause ()
DEFINT A-Z

  '$INCLUDE: 'CMFILES.BI'

  REDIM CMTRRec(1) AS CMTransRecType            ' open transaction file
  CMTRRecLen = LEN(CMTRRec(1))
  CLS
  PRINT
  PRINT "Correcting Operator Error. . ."
  PRINT
  PRINT

  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS 1 LEN = CMTRRecLen
  NumRT& = LOF(1) / CMTRRecLen
  FOR ZZ& = 1 TO NumRT&
    GET 1, ZZ&, CMTRRec(1)
    LOCATE 5, 1
    PRINT "Scanning:"; ZZ&; "of"; NumRT&;
    IF ZZ& = 74736 THEN
      CMTRRec(1).TransCash = 0
      CMTRRec(1).TransAmount = 64.2
      PUT 1, ZZ&, CMTRRec(1)
      BadCnt = BadCnt + 1
    END IF
    IF ZZ& MOD 5 = 0 THEN
      SmallPause
    END IF
  NEXT

  CLOSE

  PRINT
  PRINT
  PRINT "Corrected:"; BadCnt
  CALL KillFile("Fixgoof.exe")
  PRINT
  PRINT
  PRINT "Correction Complete."
  END

SUB SmallPause
st# = TIMER
DO
  st2# = st# + .000001
LOOP UNTIL TIMER > st2#

END SUB

