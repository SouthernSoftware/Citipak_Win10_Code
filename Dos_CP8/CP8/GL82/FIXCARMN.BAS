DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

'$INCLUDE: 'GL.BI'

  DIM GLRec AS GLAcctRecType                  'Account Record Type: ? bytes
  GLLen = LEN(GLRec)
  OPEN "glacct.dat" FOR RANDOM AS #1 LEN = GLLen
  NumOfRec = LOF(1) \ GLLen
  FOR Cnt = 1 TO NumOfRec
    GET #1, Cnt, GLRec
    Acct$ = GLRec.Num
    DO
      DashPos = INSTR(Acct$, "-")
      IF DashPos > 0 THEN
        Acct$ = LEFT$(Acct$, DashPos - 1) + MID$(Acct$, DashPos + 1)
      ELSE
        EXIT DO
      END IF
    LOOP
    Acct$ = QPTrim$(Acct$)
    Acct$ = LEFT$(Acct$, 2) + "-" + MID$(Acct$, 3, 2) + "-" + RIGHT$(Acct$, 4)
    PRINT Acct$
    GLRec.Num = Acct$
    PUT #1, Cnt, GLRec
  NEXT
  CLOSE

