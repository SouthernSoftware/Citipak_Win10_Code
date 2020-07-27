DEFINT A-Z


 '$INCLUDE: 'CMFILES.BI'

  CLS
  PRINT "Processing. . ."

  REDIM CMTrans(1) AS CMTransRecType

  CMTRLen = LEN(CMTrans(1))
  OPEN "cmtrans.dat" FOR RANDOM AS #1 LEN = CMTRLen
  NumCMTR& = LOF(1) / CMTRLen
  Cnt& = 112272
  GET #1, Cnt&, CMTrans(1)
  CMTrans(1).TransAmount = 200.5
  CMTrans(1).TransCheck = 0

  PUT #1, Cnt&, CMTrans(1)

  CLOSE
  PRINT "Done."

  CALL KillFile("fixelktn.exe")

