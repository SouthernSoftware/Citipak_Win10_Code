DEFINT A-Z

  '$INCLUDE: 'CMFILES.BI'

  REDIM CMTRRec(1) AS CMTransRecType            ' open transaction file
  CMTRRecLen = LEN(CMTRRec(1))
  CLS
  PRINT "Correcting Transaction. . ."
  PRINT
  PRINT

  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS 1 LEN = CMTRRecLen
  NumRT& = LOF(1) / CMTRRecLen
  FOR zz& = 1 TO NumRT&
    GET 1, zz&, CMTRRec(1)
    LOCATE 5, 1
    PRINT "Processing:"; zz&; "of"; NumRT&;
    IF CMTRRec(1).TransCheck < 0 THEN
      CMTRRec(1).TransCheck = 0
      PUT 1, zz&, CMTRRec(1)
      BadCnt = BadCnt + 1
    END IF
  NEXT

  CLOSE

  PRINT
  PRINT
  PRINT "Corrected:"; BadCnt
  CALL KillFile("Fixcmtr.exe")

