DEFINT A-Z
DECLARE FUNCTION Date2Num% (Dat$)

'$INCLUDE: 'cmfiles.bi'
'$INCLUDE: 'arv85.bi'

  CLS

  ADate = Date2Num("10-30-2000")
  REDIM CMTransRec(1)  AS CMTransRecType
  CMTransRecLen = LEN(CMTransRec(1))
  REDIM ARTransRec(1)  AS ARTransRecType
  ARTransRecLen = LEN(ARTransRec(1))

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

  OPEN "ARTRANS.DAT" FOR RANDOM AS #1 LEN = ARTransRecLen
  NumTrans& = LOF(1) \ ARTransRecLen
  FOR Cnt& = 1 TO NumTrans&
    GET #1, Cnt&, ARTransRec(1)
    IF ARTransRec(1).TransDate >= ADate THEN
      IF ARTransRec(1).Posted2GL = "Y" THEN
        ARTransRec(1).Posted2GL = ""
        PUT #1, Cnt&, ARTransRec(1)
        Found = Found + 1
      END IF
    END IF
  NEXT
  CLOSE

  PRINT "UN-Grabbed:"; Found

  KILL "UNGRAB.EXE"

  END

