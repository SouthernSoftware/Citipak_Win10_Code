DEFINT A-Z
DECLARE SUB KillFile (FileName$)

  '$INCLUDE: 'cmfiles.BI'
  CLS


  LOCATE 3, 1, 0
  PRINT "Correcting Transaction."

  REDIM CMTranRec(1) AS CMTransRecType
  CMTranRecLen = LEN(CMTranRec(1))


  CMFile = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CMFile LEN = CMTranRecLen
  GET CMFile, 8438, CMTranRec(1)
  CMTranRec(1).TransCheck = 48.73
  PUT CMFile, 8438, CMTranRec(1)
  CLOSE

  PRINT
  PRINT
  PRINT "Processing Complete."
  KillFile "fixcmtrn.exe"

