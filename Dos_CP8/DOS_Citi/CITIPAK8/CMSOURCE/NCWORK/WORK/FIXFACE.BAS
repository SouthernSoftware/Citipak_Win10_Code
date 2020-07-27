DEFINT A-Z
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'CMFILES.BI'

  REDIM CMTRRec(1) AS CMTransRecType            ' open transaction file
  CMTRRecLen = LEN(CMTRRec(1))
  ThisDate = Date2Num("11-14-2003")
  CLS
  PRINT "Correcting Transaction. . ."
  PRINT
  PRINT


  CMFile = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CMFile LEN = CMTRRecLen

  GET CMFile, 28537, CMTRRec(1)
  CMTRRec(1).TransDate = ThisDate
  PUT CMFile, 28537, CMTRRec(1)

  CLOSE

  CALL KillFile("FixFACE.exe")
  PRINT "Done."

