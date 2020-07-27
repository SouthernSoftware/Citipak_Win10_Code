DEFINT A-Z

  '$INCLUDE: 'CMFILES.BI'

  REDIM CMTRRec(1) AS CMTransRecType            ' open transaction file
  CMTRRecLen = LEN(CMTRRec(1))
  CLS
  PRINT
  PRINT "Correcting Transaction. . ."
  PRINT
  
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS 1 LEN = CMTRRecLen

  zz& = 112640
  GET #1, zz&, CMTRRec(1)
  CMTRRec(1).TransAmount = 55.6
  CMTRRec(1).TransCheck = 0
  PUT 1, zz&, CMTRRec(1)

  zz& = 112667
  GET 1, zz&, CMTRRec(1)
  CMTRRec(1).TransAmount = 90.88
  CMTRRec(1).TransCheck = 0
  PUT 1, zz&, CMTRRec(1)

  CLOSE

  CALL KillFile("Fixcmtr.exe")
  PRINT "Done."
  END

