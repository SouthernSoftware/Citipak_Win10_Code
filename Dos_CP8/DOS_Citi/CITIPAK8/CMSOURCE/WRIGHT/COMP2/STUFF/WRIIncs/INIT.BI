DIM GLSetup AS GLSetupRecType
SetupRecLen = LEN(GLSetup)

OPEN "GLSETUP.DAT" FOR RANDOM AS #1 LEN = SetupRecLen

   IF LOF(1) = 0 THEN
      CLS
      CLOSE
      PRINT "You must run GLSYS.EXE from prompt to define an account structure first."
      AnyKey$ = INPUT$(1)
      END
   END IF

   GET #1, 1, GLSetup

   User$ = QPTrim$(GLSetup.UserName)
   TotLen = GLSetup.TotAcctLen
   FundLen = GLSetup.FundLen
   AcctLen = GLSetup.AcctLen
   DetLen = GLSetup.DetLen
   CashAcct$ = QPTrim$(GLSetup.CashAcct)
   APAcct$ = QPTrim$(GLSetup.APAcct)
   EncAcct$ = QPTrim$(GLSetup.EncAcct)
   FBAcct$ = QPTrim$(GLSetup.FBAcct)
   CDCash$ = QPTrim$(GLSetup.CDCash)
   CDDue$ = QPTrim$(GLSetup.CDDue)
   CDActive$ = QPTrim$(GLSetup.CDActive)
   CRCashAcct$ = QPTrim$(GLSetup.CRCashAcct)

CLOSE #1

