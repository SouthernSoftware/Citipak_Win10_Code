  DEFINT A-Z

  '$INCLUDE: 'ubsetup.bi'

  DIM LogRec AS UBLogiconGetReadRecType
  DIM OldLog AS UBLogiconReadRecType
  DIM LogLen AS INTEGER
  DIM OldLen AS INTEGER

  DIM OldRCnt AS LONG
  DIM LogRCnt AS LONG

  DIM Cnt AS INTEGER

  LogLen = LEN(LogRec)
  OldLen = LEN(OldLog)

  OPEN "c:\tst\wblogno3.dat" FOR RANDOM AS #1 LEN = LogLen
  OPEN "c:\tst\new3.txt" FOR RANDOM AS #2 LEN = OldLen
  LogRCnt = LOF(1) / LogLen
  OldRCnt = LOF(2) / OldLen
  FOR Cnt = 1 TO LogRCnt
    GET #1, Cnt, LogRec
    GET #2, Cnt, OldLog
    IF Cnt > 1 THEN STOP
    IF LogRec.RecType <> "A" THEN
      GOTO NotThisOne
    END IF
    IF VAL(LogRec.PrevRead) = 0 THEN
      STOP
    END IF
    GOTO NotThisOne

    IF Cnt > OldRCnt THEN  'IF VAL(LogRec.PrevRead) = 0 THEN
      GOTO NotThisOne
    ELSE
      GET #2, Cnt, OldLog
      IF VAL(OldLog.AcctNo) <> VAL(LogRec.AcctNo) THEN
        STOP
      ELSE
        LogRec.PrevRead = OldLog.PrevRead
        LOCATE 1, 1: PRINT Cnt
        PUT #1, Cnt, LogRec
      END IF
    END IF

NotThisOne:
  NEXT

  CLOSE

  END

