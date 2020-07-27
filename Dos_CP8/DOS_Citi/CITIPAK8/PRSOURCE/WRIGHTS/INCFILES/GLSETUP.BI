TYPE GLSetupRecType                 'still under const.
   UserName    AS STRING * 30
   TotAcctLen  AS INTEGER
   FundLen     AS INTEGER
   AcctLen     AS INTEGER
   DetLen      AS INTEGER
   CashAcct    AS STRING * 14
   APAcct      AS STRING * 14
   EncAcct     AS STRING * 14
   FBAcct      AS STRING * 14
   FYBeg       AS INTEGER
   FYEnd       AS INTEGER
   NYBeg       AS INTEGER
   NYEnd       AS INTEGER
   CDCash      AS STRING * 14  'new
   CDDue       AS STRING * 14
   CDActive    AS STRING * 1
END TYPE
