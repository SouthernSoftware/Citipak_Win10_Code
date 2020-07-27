DEFINT A-Z
DECLARE FUNCTION Date2Num (TheDate$)

  BDate = Date2Num("07-15-2044")
  GDate = Date2Num("07-15-2004")

 '$INCLUDE: 'CMFILES.BI'
  CLS
  PRINT "Processing. . ."

  REDIM CMTrans(1) AS CMTransRecType

  CMTRLen = LEN(CMTrans(1))
  OPEN "cmtrans.dat" FOR RANDOM AS #1 LEN = CMTRLen
  NumCMTR& = LOF(1) / CMTRLen
  FOR Cnt& = 1 TO NumCMTR&
    GET #1, Cnt&, CMTrans(1)
    IF CMTrans(1).TransDate = BDate THEN
      CMTrans(1).TransDate = GDate
      Fixed = Fixed + 1
      PUT #1, Cnt&, CMTrans(1)
    END IF
  NEXT
  CLOSE
  PRINT "Fixed:"; Fixed

  CALL KillFile("fixgoof.exe")

