DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

'$INCLUDE: 'GL.bi'                        'GL File Types

 CLS
 PRINT "Correcting GL Transactions. . ."

 DIM Trans AS GLTransRecType
 RecLen = LEN(Trans)
 
 OPEN "GLTRANS.DAT" FOR RANDOM AS #1 LEN = RecLen
 NumOfTran& = LOF(1) / RecLen
 FOR Cnt& = 1 TO NumOfTran&
   LOCATE 5, 1
   PRINT "Processing:"; Cnt&; " of"; NumOfTran&;
   GET #1, Cnt&, Trans
   Acct$ = QPTrim$(Trans.AcctNum)
   DO
     DashPos = INSTR(Acct$, "-")
     IF DashPos > 0 THEN
       Acct$ = LEFT$(Acct$, DashPos - 1) + MID$(Acct$, DashPos + 1)
     END IF
   LOOP WHILE DashPos > 0
   Acct$ = QPTrim$(Acct$)
   Acct$ = LEFT$(Acct$, 2) + "-" + MID$(Acct$, 3, 2) + "-" + RIGHT$(Acct$, 4)
   LSET Trans.AcctNum = Acct$
   PUT #1, Cnt&, Trans
 NEXT
 CLOSE

 LOCATE 10, 1
 PRINT "Processing Complete."

