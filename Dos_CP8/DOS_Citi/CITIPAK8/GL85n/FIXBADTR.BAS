DEFINT A-Z

DECLARE FUNCTION Date2Num% (d$)
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'GL.bi'                        'GL File & Record Types

  CONST False = 0, True = NOT False

  COLOR 15, 1
  CLS

  PRINT "Correcting Transactions"
  PRINT
  
  REDIM GLTrans(1) AS GLTransRecType
  GLRecLen = LEN(GLTrans(1))

  OPEN "GLTRANS.DAT" FOR RANDOM AS #1 LEN = GLRecLen
  NumOfRec& = LOF(1) \ GLRecLen
  FOR Cnt& = 1 TO NumOfRec&
    GET #1, Cnt&, GLTrans(1)
    IF (GLTrans(1).DrAmt < 2000000#) OR (GLTrans(1).CrAmt < 2000000#) THEN
      GLTrans(1).DrAmt = 0
      GLTrans(1).CrAmt = 0
      PUT #1, Cnt&, GLTrans(1)
      Fixed = Fixed + 1
    END IF
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT "Corrected: "; Fixed
  END

