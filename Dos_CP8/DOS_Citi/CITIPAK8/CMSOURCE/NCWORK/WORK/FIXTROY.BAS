DEFINT A-Z
DECLARE FUNCTION Date2Num (TheDate$)

  BDate = Date2Num("01-13-2003")
  GDate = Date2Num("01-13-2004")

 '$INCLUDE: 'CMFILES.BI'
  CLS
  PRINT "Processing. . ."

  REDIM CMTrans(1) AS CMTransRecType

  CMTRLen = LEN(CMTrans(1))
  OPEN "cmtrans.dat" FOR RANDOM AS #1 LEN = CMTRLen
  NumCMTR& = LOF(1) / CMTRLen
'  FOR Cnt& = NumCMTR& - 1300 TO NumCMTR&
    GET #1, 115507, CMTrans(1)
      CMTrans(1).TransDate = GDate
      CMTrans(1).TransAmount = 22.57
      CMTrans(1).TransCheck = 0
      Fixed = Fixed + 1
     PUT #1, 115507, CMTrans(1)
 ' NEXT
  CLOSE
  PRINT "Fixed:"; Fixed

  CALL KillFile("fixtroy.exe")

