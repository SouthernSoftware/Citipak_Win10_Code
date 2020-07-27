DEFINT A-Z

  '$INCLUDE: 'CMFILES.BI'

  DIM CMTRRec AS CMTransRecType            ' open transaction file
  CMTRRecLen = LEN(CMTRRec)
  CLS
  PRINT
  PRINT
  PRINT "Correcting Transaction. . ."
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS #1 LEN = CMTRRecLen

  FOR tt& = 154202 TO 154202
    GET #1, tt&, CMTRRec

    CMTRRec.TransCheck = 124.04
    CMTRRec.TransAmount = 124.04
    CMTRRec.TransCash = 0
    PUT #1, tt&, CMTRRec
  NEXT

  CLOSE
  
  PRINT
  PRINT
  PRINT
  PRINT "Done."

  CALL KillFile("fixcmtrn.exe")

