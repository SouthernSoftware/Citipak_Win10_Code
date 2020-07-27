DEFINT A-Z

'$INCLUDE: 'CMFILES.BI'

  CLS
  LOCATE 3, 1
  PRINT "Correcting Transaction"
  RecNo& = 40454

  REDIM CMTrans(1) AS CMTransRecType
  CMTRLen = LEN(CMTrans(1))

  CMFile = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CMFile LEN = CMTRLen

  GET CMFile, RecNo&, CMTrans(1)

  CMTrans(1).TransCheck = 0
  CMTrans(1).TransAmount = CMTrans(1).TransCheck + CMTrans(1).TransCash

  PUT CMFile, RecNo&, CMTrans(1)
  CLOSE

  LOCATE 10, 1
  PRINT "Done."

  CALL Killfile("fixrussl.exe")

