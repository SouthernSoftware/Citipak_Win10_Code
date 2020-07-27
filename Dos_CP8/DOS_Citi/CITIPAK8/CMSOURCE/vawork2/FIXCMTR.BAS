DEFINT A-Z

  '$INCLUDE: 'CMFILES.BI'

  REDIM CMTRRec(1) AS CMTransRecType            ' open transaction file
  CMTRRecLen = LEN(CMTRRec(1))
  CLS
  PRINT "Correcting Transaction. . ."
  PRINT
  PRINT

  
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS 1 LEN = CMTRRecLen

  zz& = 129183
  GET 1, zz&, CMTRRec(1)
  CMTRRec(1).TransCheck = 80.57
  PUT #1, zz&, CMTRRec(1)
  CLOSE

  CALL KillFile("Fixcmtr.exe")
  PRINT "Done."

