DEFINT A-Z
DECLARE FUNCTION Date2Num (TheDate$)

  BDate = Date2Num("04-20-2005")
  GDate = Date2Num("04-21-2005")

 '$INCLUDE: 'CMFILES.BI'
  CLS
  PRINT "Correcting operator error. . ."

  REDIM CMTrans(1) AS CMTransRecType

  CMTRLen = LEN(CMTrans(1))
  OPEN "cmtrans.dat" FOR RANDOM AS #1 LEN = CMTRLen
  NumCMTR& = LOF(1) / CMTRLen
  FOR Cnt& = 18588 TO 18591
    LOCATE 5, 1
'    PRINT "Processing:"; Cnt&; "of"; NumCMTR&;

    GET #1, Cnt&, CMTrans(1)
    CMTrans(1).TransDate = GDate

    PUT #1, Cnt&, CMTrans(1)
  NEXT
  CLOSE
  LOCATE 5, 1
  PRINT "Correction Complete."
  PRINT
  PRINT "Corrected:"; 4

  CALL KillFile("fixgoof.exe")

