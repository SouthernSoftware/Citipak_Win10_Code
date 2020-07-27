'TYPE Acct9Type
'   Acct AS STRING * 9
'END TYPE

TYPE GLAcctIdxType
    AcctNum AS DOUBLE
    RecNo   AS SINGLE
END TYPE

TYPE JGLAcctIdxType
    AcctNum AS STRING * 14
    RecNo   AS INTEGER
END TYPE

TYPE GLIFDataType79
   TranAcct   AS STRING * 9
   TranDate   AS STRING * 8
   TranDesc   AS STRING * 20
   CrAmt      AS DOUBLE
   DrAmt      AS DOUBLE
   Detail     AS DOUBLE
   Source     AS STRING * 8
   rRNEXTTR   AS SINGLE
   Fill       AS STRING * 11
   FromFlag   AS STRING * 1
END TYPE

TYPE GLIFDataType10
   TranAcct   AS STRING * 10     ' 10 AS ARTRANACCT$
   TranDate   AS STRING * 8      ' 8 AS artrandate$
   TranDesc   AS STRING * 20     ' 20 AS artrandesc$
   CrAmt      AS DOUBLE          ' 8 AS ARCRAMT$
   DrAmt      AS DOUBLE          ' 8 AS ARDRAMT$
   Detail     AS DOUBLE          ' 8 AS ardetail$
   Source     AS STRING * 8      ' 8 AS arsource$
   rRNEXTTR   AS SINGLE          ' 4 AS ARNEXTTR$
   Fill       AS STRING * 10
   FromFlag   AS STRING * 1
END TYPE

TYPE GLIFDataType11
   TranAcct   AS STRING * 11     ' 11 AS ARTRANACCT$
   TranDate   AS STRING * 8      ' 8 AS artrandate$
   TranDesc   AS STRING * 20     ' 20 AS artrandesc$
   CrAmt      AS DOUBLE          ' 8 AS ARCRAMT$
   DrAmt      AS DOUBLE          ' 8 AS ARDRAMT$
   Detail     AS DOUBLE          ' 8 AS ardetail$
   Source     AS STRING * 8      ' 8 AS arsource$
   rRNEXTTR   AS SINGLE          ' 4 AS ARNEXTTR$
   Fill       AS STRING * 9
   FromFlag   AS STRING * 1
END TYPE

TYPE GLIFDataType12              'Hummm.
   TranAcct   AS STRING * 12     ' 12 AS ARTRANACCT$
   TranDate   AS STRING * 8      ' 8 AS artrandate$
   TranDesc   AS STRING * 20     ' 19 AS artrandesc$
   CrAmt      AS DOUBLE          ' 8 AS ARCRAMT$
   DrAmt      AS DOUBLE          ' 8 AS ARDRAMT$
   Detail     AS DOUBLE          ' 8 AS ardetail$
   Source     AS STRING * 8      ' 8 AS arsource$
   rRNEXTTR   AS SINGLE          ' 4 AS ARNEXTTR$
   Fill       AS STRING * 8
   FromFlag   AS STRING * 1
END TYPE

TYPE GLIFDataType14
   TranAcct   AS STRING * 14   'For New G/L
   TranDate   AS STRING * 8
   TranDesc   AS STRING * 20
   CrAmt      AS DOUBLE
   DrAmt      AS DOUBLE
   Detail     AS DOUBLE
   Source     AS STRING * 8
   rRNEXTTR   AS SINGLE
   Fill       AS STRING * 6    'Adj. from 8 to 6
   FromFlag   AS STRING * 1
END TYPE

TYPE FundType
   FundCode   AS STRING * 14
   Credit     AS DOUBLE
   Debit      AS DOUBLE
   Net        AS DOUBLE
END TYPE

TYPE AcctSumType
   FundCode   AS STRING * 14
   Credit     AS DOUBLE
   Debit      AS DOUBLE
   'Net        AS DOUBLE
END TYPE

'TYPE oFundType
'   FundCode   AS STRING * 11
'   Credit     AS DOUBLE
'   Debit      AS DOUBLE
'   Net        AS DOUBLE
'END TYPE

'TYPE oSplitDedsType
'   Acct       AS STRING * 11
'   FedTax     AS DOUBLE
'   StaTax     AS DOUBLE
'   MedTax     AS DOUBLE
'   SocTax     AS DOUBLE
'   RetTax     AS DOUBLE
'   EICPay     AS DOUBLE
'   DedAmts(1 TO 12) AS DOUBLE
'END TYPE

TYPE SplitDedsType
   Acct       AS STRING * 14
   FedTax     AS DOUBLE
   StaTax     AS DOUBLE
   MedTax     AS DOUBLE
   SocTax     AS DOUBLE
   RetTax     AS DOUBLE
   EICPay     AS DOUBLE
   DedAmts(1 TO 12) AS DOUBLE
END TYPE

