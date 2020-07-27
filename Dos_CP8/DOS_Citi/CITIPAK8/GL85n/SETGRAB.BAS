DEFINT A-Z
DECLARE FUNCTION Date2Num% (Dat$)

'$INCLUDE: 'cmfiles.bi'
'$INCLUDE: 'arv85.bi'

  CLS

  ADate = Date2Num("11-01-2000")
  REDIM CMTransRec(1)  AS CMTransRecType
  CMTransRecLen = LEN(CMTransRec(1))
  REDIM ARTransRec(1)  AS ARTransRecType
  ARTransRecLen = LEN(ARTransRec(1))

'
  OPEN "CMTRANS.DAT" FOR RANDOM AS #1 LEN = CMTransRecLen
  NumTrans& = LOF(1) \ CMTransRecLen
  FOR Cnt& = 1 TO NumTrans&
    GET #1, Cnt&, CMTransRec(1)
    IF CMTransRec(1).TransDate <= ADate THEN
      CMTransRec(1).Trans2GL = "Y"
      PUT #1, Cnt&, CMTransRec(1)
      Found = Found + 1
    END IF
  NEXT
  CLOSE


  PRINT "Set-Grabbed:"; Found

  KILL "SETGRAB.EXE"

  END

