DEFINT A-Z
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'CMFILES.BI'

  BDate = Date2Num("03-01-2001")

  REDIM CMTrRec(1) AS CMTransRecType
  CMTrRecLen = LEN(CMTrRec(1))
  CHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS CHandle LEN = CMTrRecLen
  RecCnt& = (LOF(CHandle) \ CMTrRecLen)
  FOR Cnt& = 1 TO RecCnt&
    GET CHandle, Cnt&, CMTrRec(1)
    IF CMTrRec(1).TransDate >= BDate THEN
      IF CMTrRec(1).Trans2GL = "Y" THEN
        CMTrRec(1).Trans2GL = ""
        Fixed& = Fixed& + 1
        PUT CHandle, Cnt&, CMTrRec(1)
      END IF
    END IF
  NEXT
  CLOSE
  CLS
  PRINT "Fixed:"; Fixed&
  call KillFile("CMUNGRAB.EXE")

