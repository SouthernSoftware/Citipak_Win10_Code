DEFINT A-Z

'$INCLUDE: 'premprec.bi'
'$INCLUDE: 'pro7.bi'

DECLARE FUNCTION StripAcct$ (AcctNum$)

CLS
LOCATE 1, 1, 0
Colr = 15

DIM Emp2Rec AS EmpData2Type
Emp2RecLen = LEN(Emp2Rec)
OPEN "premp2.dat" FOR RANDOM AS #1 LEN = Emp2RecLen
NumRecs = LOF(1) \ Emp2RecLen

OPEN "PREMPCHG.LOG" FOR OUTPUT AS #2

FOR Record = 1 TO NumRecs

  GET 1, Record, Emp2Rec
  EmpName$ = QPTrim$(Emp2Rec.EmpLname) + ", " + Emp2Rec.EmpFName
  QPrintRC "Processing: " + EmpName$, 1, 1, Colr
  PRINT #2, EmpName$

  FOR WageCnt = 1 TO 8
    SAcct$ = StripAcct$(Emp2Rec.EDist(WageCnt).DAcct)
    NewAcct$ = QPTrim$(SAcct$)
    IF LEN(NewAcct$) = 7 THEN
      PRINT #2, Emp2Rec.EDist(WageCnt).DAcct;
      NewAcct$ = NewAcct$ + "00"
      PRINT #2, NewAcct$
      Emp2Rec.EDist(WageCnt).DAcct = NewAcct$
      PUT 1, Record, Emp2Rec
    END IF

  NEXT

NEXT
CLOSE

CLS
QPrintRC "Procedure complete. Press any key to continue.", 1, 1, Colr
K$ = INPUT$(1)
CLS

END

FUNCTION StripAcct$ (AcctNum$)

   X$ = QPTrim$(AcctNum$)  '(Form$(AcctNum, 0))
   DO
      DashPos = INSTR(X$, "-")
      IF DashPos > 0 THEN
         X$ = LEFT$(X$, DashPos - 1) + MID$(X$, DashPos + 1)
      END IF
    LOOP WHILE DashPos

    StripAcct$ = X$


END FUNCTION

