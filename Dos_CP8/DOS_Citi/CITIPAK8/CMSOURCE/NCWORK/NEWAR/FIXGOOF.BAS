DEFINT A-Z
DECLARE FUNCTION Date2Num (TheDate$)

  BDate = Date2Num("03-21-2000")

 '$INCLUDE: 'CMFILES.BI'
  CLS
  PRINT "Correcting Entry Error. . ."

  REDIM CMTrans(1) AS CMTransRecType

  CMTRLen = LEN(CMTrans(1))

  OPEN "cmtrans.dat" FOR RANDOM AS #1 LEN = CMTRLen
  GET #1, 39526, CMTrans(1)
  CMTrans(1).TransCheck = 212.63
  'CMTrans(1).TransCash = 212.63
  CMTrans(1).TransAmount = 212.63
  PUT #1, 39526, CMTrans(1)
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT "Correction Complete. . ."
  CALL KillFile("fixgoof.exe")

