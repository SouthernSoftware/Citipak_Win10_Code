TYPE BAAcctRecType
   Num       AS STRING * 12   '12 AS MASTACCT$
   Title     AS STRING * 35   '35 AS ACCTNAME$
   Typ       AS STRING * 1    ' 1 AS ACCTTYPE$
   Bal       AS DOUBLE        ' 8 AS acctbal$
   MTD       AS DOUBLE        ' 8 AS mtdact$
   YTD       AS DOUBLE        ' 8 AS ytdact$
   Bgt       AS DOUBLE         ' 8 AS BUDAMT$
   FrstTran  AS SINGLE        ' 4 AS FIRST$
   LastTran  AS SINGLE        ' 4 AS LAST$
   BegBal    AS DOUBLE        ' 8 AS begbal$
   ReCalcBal AS DOUBLE        ' 8 AS recalcbal$
   Encumb    AS DOUBLE        ' 8 AS encumb$
   Fill      AS STRING * 16   '16 AS blank$
END TYPE

TYPE BAAcctNdxType
   AcctNum   AS DOUBLE
   RecNum    AS SINGLE
END TYPE

TYPE BATransRecType ' McCormick, SC
   AcctNum    AS STRING * 11
   TrDate     AS STRING * 8
   Desc       AS STRING * 20
   CrAmt      AS DOUBLE
   DrAmt      AS DOUBLE
   Ref        AS STRING * 8
   Source     AS STRING * 8
   NextTran   AS SINGLE
   Fill       AS STRING * 9
   'FromFlag   AS STRING * 1
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
   FundCode AS STRING * 2     '2 AS fundacct$
END TYPE

TYPE CRDistType 'OPEN "r", 7, "crdist.dat", 33
   TrRec    AS SINGLE         '4 AS trrec$
   Acct     AS STRING * 9     '9 AS daacct$
   CrAmt    AS DOUBLE         '8 AS dcramt$
   DrAmt    AS DOUBLE         '8 AS ddramt$
   NextDist AS SINGLE         '4 AS DNEX$
END TYPE

