TYPE WBLongViewCustRecType    'len=512
     LastName       AS STRING * 25        '25 AS lname$
     FirstName      AS STRING * 25        '25 AS FIRSTNAME$
     Add1           AS STRING * 30        '30 AS ADD1$
     Add2           AS STRING * 30        '30 AS add2$
     City           AS STRING * 20        '20 AS CITY$
     State          AS STRING * 2         '2 AS ST$
     Zip            AS STRING * 5         '5 AS ZP$
     Phone          AS STRING * 12        '12 AS PHONE$
     Meter          AS STRING * 12        '12 AS meter$
     WRate          AS INTEGER            '2 AS WRATE$
     ERate          AS INTEGER            '2 AS ERATE$
     CityLimit      AS STRING * 1         '1 AS CTYLIMIT$
     CustType       AS STRING * 1         '1 AS CUSTYPE$
     WAvgUse        AS LONG               '4 AS WAVGUSE$
     WAvgCnt        AS INTEGER            '2 AS wavgct$
     CurBal         AS SINGLE             '4 AS curbal$
     PastDue        AS SINGLE             '4 AS pastdue$
'     FIELD 1     181 AS dum1$
     Active         AS STRING * 1         '1 AS active$
     WCurRead       AS LONG               '4 AS wcurread$
     WPastRead      AS LONG               '4 AS wpastread$
     ECurRead       AS LONG               '4 AS ecurread$
     EPastRead      AS LONG               '4 AS epastread$

     FirstTrans     AS LONG               '4 AS FIRSTTR$
     LastTrans      AS LONG               '4 AS LASTTR$
     
     WaterDep       AS SINGLE             '4 AS WATERDEP$
     ElecDep        AS SINGLE             '4 AS ELECDEP$
     IRevenue1      AS SINGLE             '4 AS IREVENUE1$
     IRevenue2      AS SINGLE             '4 AS IREVENUE2$
     IRevenue3      AS SINGLE             '4 AS IREVENUE3$
     IRevenue4      AS SINGLE             '4 AS IREVENUE4$
     IRevenue5      AS SINGLE             '4 AS irevenue5$
'     FIELD 1     234 AS dum2$
     IRevenue6      AS SINGLE             ',4 AS irevenue6$
     RAmt1          AS SINGLE             ' 4 AS ramt1$
     RAmt2          AS SINGLE             ' 4 AS ramt2$
     RAmt3          AS SINGLE             ' 4 AS ramt3$

     RDesc1         AS STRING * 25        ' 25 AS RDESC1$
     RDesc2         AS STRING * 25        ' 25 AS RDESC2$
     RDesc3         AS STRING * 25        ' 25 AS RDESC3$

     RAcct1         AS STRING * 2          '2 AS racct1$
     RAcct2         AS STRING * 2          '2 AS racct2$
     RAcct3         AS STRING * 2          '2 AS racct3$

     OAmt1          AS SINGLE              '4 AS oamt1$
     OAmt2          AS SINGLE              '4 AS oamt2$
     ODesc1         AS STRING * 25         '25 AS ODESC1$
     ODesc2         AS STRING * 25         '25 AS ODESC2$
     OAcct1         AS STRING * 2          '2 AS oacct1$
     OAcct2         AS STRING * 2          '2 AS oacct2$
'     FIELD 1     393 AS dum3$
     WMin           AS INTEGER             '2 AS WMIN$
     EMulti         AS INTEGER             '2 AS emult$
     WMeterNum1     AS LONG                '4 AS WMETERNUM$
     EMeterNum      AS LONG                '4 AS EMETERNUM$
     Tax            AS STRING * 1          '1 AS tax$
     Security1      AS INTEGER             '2 AS Security1$
     Security2      AS INTEGER             '2 AS Security2$

     FAdj           AS SINGLE              '4 AS fadj$
     EAvgUse        AS LONG                '4 AS eavguse$
     EAvgCnt        AS INTEGER             '2 AS eavgct$
     Demand         AS SINGLE              '4 AS demand$

     Address3       AS STRING * 25         '25 AS ADD3$

     NumMeters      AS INTEGER             '2 AS numbermeters$

     WCurrRead1     AS LONG                '4 AS wcurread1$
     WPastRead1     AS LONG                '4 AS wpastread1$
     WCurrRead2     AS LONG                '4 AS wcurread2$
     WPastRead2     AS LONG                '4 AS wpastread2$
     WCurrRead3     AS LONG                '4 AS wcurread3$
     WPastRead3     AS LONG                '4 AS wpastread3$

     LastReadDate   AS STRING * 8          '8 AS lmeterdate$
     CurrReadDate   AS STRING * 8          '8 AS cmeterdate$

     SensusType     AS STRING * 1          '1 AS sensustype$
     MeterSeq       AS LONG                '4 AS meterseq$

     WMeterNum2     AS LONG                '4 AS meter2id$
     WMeterNum3     AS LONG                '4 AS meter3id$
     EPACode        AS STRING * 1          '1 AS EPACODE$

     Fill           AS STRING * 7          '7 AS notused$
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

'FIELD 2,
'8 AS TDATE$,
'8 AS TAMT$,
'2 AS TTYPE$,
'15 AS TDESC$,
'4 AS TMREC$,
'4 AS TNEXTTR$,
'4 AS tcuramt$

