TYPE WBCastCustRecType    'len=256

     LastName       AS STRING * 25        '25 AS lname$
     firstname      AS STRING * 25        '25 AS FIRSTNAME$
     ADD1           AS STRING * 30        '30 AS ADD1$
     ADD2           AS STRING * 30        '30 AS add2$
     CITY           AS STRING * 20        '20 AS CITY$

     St             AS STRING * 2         '2 AS ST$
     Zp            AS STRING * 9         '5 AS ZP$

     PHONE          AS STRING * 12        '12 AS PHONE$
     zBook          AS STRING * 2
     meter          AS STRING * 12        '12 AS meter$
     RateC          AS INTEGER            '2 AS WRATE$

     Watertest      AS STRING * 8
     Sanitation     AS STRING * 8
     AvgUse         AS DOUBLE
     AvgCnt         AS INTEGER
     CurBal         AS DOUBLE
     PastDue        AS DOUBLE
     Cycle          AS STRING * 1
     Fill1          AS STRING * 5
     CurrRead       AS DOUBLE
     PastRead       AS DOUBLE
     FirstTR        AS DOUBLE
     LastTR         AS DOUBLE
     waterdep       AS SINGLE
     Fill2          AS STRING * 3
END TYPE

TYPE WBTransType  'len=50
     TransDate   AS STRING * 8   'Trans date
     TransAmt    AS DOUBLE       'total amt Past + Curr
     TransType   AS INTEGER      '1=Charge,2=Payment,3=Pen'ty/Recon,4=Adjust
     TransDesc   AS STRING * 15  'Transaction description
     CustRecNum  AS SINGLE       'cust acct  rec
     NextTrans   AS SINGLE       'Pointer to next trans
     TCurrAmt    AS SINGLE       'current amt
     Pad         AS STRING * 5
END TYPE

TYPE WBCastCust1RecType  'len=256

    RAMT1     AS SINGLE
    RAMT2     AS SINGLE
    RAMT3     AS SINGLE
    RDESC1    AS STRING * 25
    RDESC2    AS STRING * 25
    RDESC3    AS STRING * 25

    RACCT1    AS STRING * 9
    RACCT2    AS STRING * 9
    RACCT3    AS STRING * 9

    OAMT1     AS SINGLE
    OAMT2     AS SINGLE
    ODESC1    AS STRING * 25
    ODESC2    AS STRING * 25
    OACCT1    AS STRING * 9
    OACCT2    AS STRING * 9
    sewerrate AS INTEGER
    CUSTTYPE  AS INTEGER
    nu        AS STRING * 62

END TYPE



'FIELD 2,
'8 AS TDATE$,
'8 AS TAMT$,
'2 AS TTYPE$,
'15 AS TDESC$,
'4 AS TMREC$,
'4 AS TNEXTTR$,
'4 AS tcuramt$

