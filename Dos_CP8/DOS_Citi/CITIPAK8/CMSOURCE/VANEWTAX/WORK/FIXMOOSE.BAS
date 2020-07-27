DEFINT A-Z

  '$INCLUDE: 'CMFILES.BI'

  DIM CMTRRec AS CMTransRecType            ' open transaction file
  CMTRRecLen = LEN(CMTRRec)
  CLS
  PRINT
  PRINT
  PRINT "Correcting Transaction. . ."
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS #1 LEN = CMTRRecLen

  FOR tt& = 152159 TO 152159
    GET #1, tt&, CMTRRec
    
    CMTRRec.TransCheck = 0
    CMTRRec.TransAmount = 151
    'CMTRRec.TransCash =
    'CMTRRec.TransRevAmt(2) = 0
    PUT #1, tt&, CMTRRec
  NEXT

  CLOSE
  
  PRINT
  PRINT
  PRINT
  PRINT "Done."

  CALL KillFile("fixmoose.exe")

