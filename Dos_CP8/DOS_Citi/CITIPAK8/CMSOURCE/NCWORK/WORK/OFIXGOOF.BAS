DEFINT A-Z

 '$INCLUDE: 'CMFILES.BI'
  CLS
  PRINT "Voiding Entry Error. . ."

  REDIM CMTrans(1) AS CMTransRecType

  CMTRLen = LEN(CMTrans(1))

  OPEN "cmtrans.dat" FOR RANDOM AS #1 LEN = CMTRLen
  GET #1, 2483, CMTrans(1)
  CMTrans(1).Transamount = 0
  CMTrans(1).TransCash = 0
  CMTrans(1).TransCheck = 0
  CMTrans(1).TransAmtOwed = 0
  CMTrans(1).TransDesc = "VOID VOID VOID"
  CMTrans(1).TransRevAmt(1) = 0

  PUT #1, 2483, CMTrans(1)

  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT "Complete. . ."
  CALL KillFile("fixgoof.exe")

