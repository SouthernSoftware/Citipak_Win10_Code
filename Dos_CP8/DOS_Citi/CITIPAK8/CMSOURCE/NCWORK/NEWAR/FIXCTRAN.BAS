DEFINT A-Z
DECLARE FUNCTION Date2Num (TheDate$)

  BDate = Date2Num("03-21-2000")

 '$INCLUDE: 'CMFILES.BI'
  CLS
  PRINT "Processing. . ."

  REDIM CMTrans(1) AS CMTransRecType

  CMTRLen = LEN(CMTrans(1))

  OPEN "cmtrans.dat" FOR RANDOM AS #1 LEN = CMTRLen
  FOR Cnt& = 1 TO LOF(1) / CMTRLen
    GET #1, Cnt&, CMTrans(1)
    IF CMTrans(1).TransDate = BDate THEN
      IF CMTrans(1).TransCheck < -100000 THEN
        CMTrans(1).TransAmount = CMTrans(1).TransCash
        CMTrans(1).TransCheck = 0
        Fixed = Fixed + 1
        PUT #1, Cnt&, CMTrans(1)
      END IF
    END IF
  NEXT
  CLOSE
  PRINT "Done"

