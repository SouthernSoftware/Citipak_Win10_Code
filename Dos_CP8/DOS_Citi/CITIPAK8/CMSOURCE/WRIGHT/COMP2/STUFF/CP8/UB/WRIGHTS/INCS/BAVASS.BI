'9 AS mastacct$, 25 AS ACCTNAME$, 1 AS ACCTTYPE$, 8 AS acctbal$, 8 AS mtdact$, 8 AS ytdact$
' 8 AS BUDAMT$, 4 AS FIRST$, 4 AS LAST$, 8 AS begbal$


TYPE BAAcctNdxType
   AcctNum   AS DOUBLE
   RecNum    AS SINGLE
END TYPE

'--this is the good one
TYPE BAAcctRecType
   Num       AS STRING * 11
   Title     AS STRING * 25
   Typ       AS STRING * 1
   Bal       AS DOUBLE
   MTD       AS DOUBLE
   YTD       AS DOUBLE
   Bgt       AS DOUBLE
   FrstTran  AS SINGLE
   LastTran  AS SINGLE
   BegBal    AS DOUBLE
   Fill      AS STRING * 15
END TYPE

'--this is for appalachia, va
'TYPE BAAcctRecType
'   Num       AS STRING * 12
'   Title     AS STRING * 25
'   Typ       AS STRING * 1
'   Bal       AS DOUBLE
'   MTD       AS DOUBLE
'   YTD       AS DOUBLE
'   Bgt       AS DOUBLE
'   FrstTran  AS SINGLE
'   LastTran  AS SINGLE
'   BegBal    AS DOUBLE
'   Fill      AS STRING * 14
'END TYPE

'TYPE BAAcctRecType
'   Num       AS STRING * 9
'   Title     AS STRING * 24
'   Typ       AS STRING * 1
'   Bal       AS DOUBLE
'   MTD       AS DOUBLE
'   YTD       AS DOUBLE
'   Bgt       AS DOUBLE
'   FrstTran  AS SINGLE
'   LastTran  AS SINGLE
'   BegBal    AS DOUBLE
'   'ReCalcBal AS DOUBLE
'   'Encumb    AS DOUBLE
'   Fill      AS STRING * 26
'END TYPE

'this is the good one
TYPE BATransRecType ' 7-9 digit
   AcctNum    AS STRING * 11
   TrDate   AS STRING * 8
   Desc       AS STRING * 20
   CrAmt      AS DOUBLE
   DrAmt      AS DOUBLE
   Ref        AS STRING * 8
   Source     AS STRING * 8
   NextTran   AS SINGLE
   Fill       AS STRING * 8
   Marked    AS INTEGER
END TYPE

'--this is for appalachia
'TYPE BATransRecType ' 7-9 digit
'   AcctNum    AS STRING * 12
'   TrDate   AS STRING * 8
'   Desc       AS STRING * 20
'   CrAmt      AS DOUBLE
'   DrAmt      AS DOUBLE
'   Ref        AS STRING * 8
'   Source     AS STRING * 8
'   NextTran   AS SINGLE
'   Fill       AS STRING * 7
'   Deleted    AS INTEGER
'END TYPE


'-THis is the good one
TYPE BABgtTransType
   Acct    AS STRING * 11  '9 AS budacct$
   TrType  AS STRING * 1  '1 AS budtype$
   BgtAmt  AS DOUBLE      '8 AS budgamt$
   TrDate  AS STRING * 8  '8 AS buddate$
   Filler AS STRING * 6
END TYPE

'--this is for appalchia
'TYPE BABgtTransType
'   Acct    AS STRING * 12  '9 AS budacct$
'   TrType  AS STRING * 1  '1 AS budtype$
'   BgtAmt  AS DOUBLE      '8 AS budgamt$
'   TrDate  AS STRING * 8  '8 AS buddate$
'   Filler AS STRING * 5
'END TYPE

'OPEN "bafund.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 64
TYPE BAFundType
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



