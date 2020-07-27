DEFINT A-Z

'$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'GL.bi'                        'GL File Types

   CLS
   PRINT "Clearing invalid entries. . ."
   ON ERROR RESUME NEXT
   DIM SHARED Acct AS GLAcctRecType

   AcctRecLen = LEN(Acct)
   AcctFileNum = FREEFILE

   OPEN "GLACCT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS AcctFileNum LEN = AcctRecLen
   NumAccts! = LOF(AcctFileNum) / AcctRecLen


   FOR CNT = 1 TO NumAccts!
   LOCATE 5, 1: PRINT "ACCT:"; CNT; " OF "; NumAccts!;

   GET #AcctFileNum, CNT, Acct
   IF Acct.PYAct < 0 THEN
     'LPRINT "NEG", Cnt, Acct.PYAct
     Acct.PYAct = 0
     PUT #AcctFileNum, CNT, Acct
   END IF
   IF Acct.PYAct > 2000000000 THEN
     'LPRINT "BIG", Cnt, Acct.PYAct
     Acct.PYAct = 0
     PUT #AcctFileNum, CNT, Acct
   END IF

   NEXT

