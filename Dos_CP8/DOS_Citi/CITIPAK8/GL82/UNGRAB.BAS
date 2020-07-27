DEFINT A-Z
DECLARE FUNCTION Date2Num% (Dat$)

'$INCLUDE: 'cmfiles.bi'

  CLS

  ADate = Date2Num("10-30-2000")
  REDIM CMTransRec(1)  AS CMTransRecType
  CMTransRecLen = LEN(CMTransRec(1))
'
  OPEN "CMTRANS.DAT" FOR RANDOM AS #1 LEN = CMTransRecLen
  NumTrans& = LOF(1) \ CMTransRecLen
  FOR Cnt& = 1 TO NumTrans&
    GET #1, Cnt&, CMTransRec(1)
    IF CMTransRec(1).TransDate >= ADate THEN
      IF CMTransRec(1).Trans2GL = "Y" THEN
        CMTransRec(1).Trans2GL = ""
        PUT #1, Cnt&, CMTransRec(1)
        Found = Found + 1
      END IF
    END IF
  NEXT
  CLOSE
  PRINT "UN-Grabbed:"; Found

  'KILL "UNGRAB.EXE"

  END

