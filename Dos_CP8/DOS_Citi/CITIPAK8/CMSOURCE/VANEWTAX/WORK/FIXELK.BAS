DEFINT A-Z

  '$INCLUDE: 'CMFILES.BI'

  DIM CMTRRec AS CMTransRecType            ' open transaction file
  CMTRRecLen = LEN(CMTRRec)

  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS #1 LEN = CMTRRecLen
  GET #1, 28828, CMTRRec
  CMTRRec.TransCheck = 0
  CMTRRec.TransAmount = CMTRRec.TransCash
  PUT #1, 28828, CMTRRec
  CLOSE
  CALL KillFile("fixelk.exe")

