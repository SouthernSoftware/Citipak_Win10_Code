'OPEN "baacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 100
'11 AS mastacct$
' 25 AS ACCTNAME$,
'  1 AS ACCTTYPE$,
'   8 AS ACCTBAL$,
'    8 AS mtdact$,
'    8 AS ytdact$,
'    8 AS BUDAMT$,
'     4 AS first$,
'      4 AS last$,
'       8 AS begbal$

TYPE BAAcctRecType
   Num       AS STRING * 11
   Title      AS STRING * 25
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

'   OPEN "babudprp.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #7 LEN = 64
TYPE BABgtPrepRecType
   EstAmt AS DOUBLE '8 AS ESTAMT$,
   ReqAmt AS DOUBLE '8 AS REQAMT$,
   RecAmt AS DOUBLE  '8 AS RECAMT$,
   AppAmt AS DOUBLE  '8 AS APPAMT$
   Fill AS STRING * 32
END TYPE

TYPE BAAcctNdxType
   AcctNum   AS DOUBLE
   RecNum    AS SINGLE
END TYPE




TYPE BATransRecType
   AcctNum    AS STRING * 11
   TrDate     AS STRING * 8
   Desc       AS STRING * 20
   CrAmt      AS DOUBLE
   DrAmt      AS DOUBLE
   Ref        AS STRING * 8
   Source     AS STRING * 8
   NextTran   AS SINGLE
   Fill       AS STRING * 8
   Deleted   AS INTEGER
END TYPE

TYPE TmpCRType   'OPEN "R", 4, "tempcr.dat", 64
   TranDate AS STRING * 8     '8 AS tidate$
   Detail AS STRING * 8       '8 AS tidet$
   Amt AS DOUBLE              '8 AS tiAMOUNT$
   Desc AS STRING * 20        '20 AS tidesc$
   FrstDist AS SINGLE         '4 AS TFDIST$
   LastDist AS SINGLE         '4 AS tldist$
END TYPE

TYPE BAFundType  'OPEN "bafund.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 64
   FundName AS STRING * 30    '30 AS fundname$
   FundCode AS INTEGER        '2 AS fundacct$
   Filler AS STRING * 32
END TYPE

TYPE CRDistType 'OPEN "r", 7, "crdist.dat", 33
   TrRec    AS SINGLE         '4 AS trrec$
   Acct     AS STRING * 9     '9 AS daacct$
   CrAmt    AS DOUBLE         '8 AS dcramt$
   DrAmt    AS DOUBLE         '8 AS ddramt$
   NextDist AS SINGLE         '4 AS DNEX$
END TYPE

TYPE BABgtTransType
   Acct    AS STRING * 11  '9 AS budacct$
   TrType  AS STRING * 1  '1 AS budtype$
   BgtAmt  AS DOUBLE      '8 AS budgamt$
   TrDate  AS STRING * 8  '8 AS buddate$
   Filler AS STRING * 6
END TYPE


