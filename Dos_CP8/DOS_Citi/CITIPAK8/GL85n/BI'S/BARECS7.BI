'9 AS mastacct$, 25 AS ACCTNAME$, 1 AS ACCTTYPE$, 8 AS acctbal$, 8 AS mtdact$, 8 AS ytdact$
' 8 AS BUDAMT$, 4 AS FIRST$, 4 AS LAST$, 8 AS begbal$

TYPE BAAcctRecType
   Num       AS STRING * 9
   Title      AS STRING * 25
   Typ       AS STRING * 1
   Bal       AS DOUBLE
   MTD       AS DOUBLE
   YTD       AS DOUBLE
   Bgt       AS DOUBLE
   FrstTran  AS SINGLE
   LastTran  AS SINGLE
   BegBal    AS DOUBLE
   'ReCalcBal AS DOUBLE
   'Encumb    AS DOUBLE
   Fill      AS STRING * 17
END TYPE


TYPE BAAcctNdxType
   AcctNum   AS DOUBLE
   RecNum    AS SINGLE
END TYPE


TYPE BATransRecType ' 7-9 digit
   AcctNum    AS STRING * 9
   TrDate   AS STRING * 8
   Desc       AS STRING * 20
   CrAmt      AS DOUBLE
   DrAmt      AS DOUBLE
   Ref        AS STRING * 8
   Source     AS STRING * 8
   NextTran   AS SINGLE
   Fill       AS STRING * 10
   Posted    AS INTEGER
END TYPE

TYPE BABgtTransType
   Acct    AS STRING * 9  '9 AS budacct$
   TrType  AS STRING * 1  '1 AS budtype$
   BgtAmt  AS DOUBLE      '8 AS budgamt$
   TrDate  AS STRING * 8  '8 AS buddate$
   Filler AS STRING * 8
END TYPE

TYPE BAFundType  'OPEN "bafund.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 64
   FundName AS STRING * 30    '30 AS fundname$
   FundCode AS INTEGER 'STRING * 2     '2 AS fundacct$
   Fill AS STRING * 32
END TYPE


TYPE BABgtPrepRecType
   EstAmt AS DOUBLE '8 AS ESTAMT$,
   ReqAmt AS DOUBLE '8 AS REQAMT$,
   RecAmt AS DOUBLE  '8 AS RECAMT$,
   AppAmt AS DOUBLE  '8 AS APPAMT$
   PYAct  AS DOUBLE
   Fill AS STRING * 24
END TYPE



