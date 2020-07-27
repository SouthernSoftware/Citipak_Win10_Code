
TYPE CopanWBCustType
     cLName AS STRING * 25 '25 AS LNAME$,
     cFName AS STRING * 25 '25 AS FIRSTNAME$,
     Addr1 AS STRING * 30 '30 AS add1$,
     Addr2 AS STRING * 30 '30 AS add2$,
     City AS STRING * 20 '20 AS city$,
     State AS STRING * 2 '2 AS st$,
     Zip AS STRING * 5 ' 5 AS zp$,
     Phone AS STRING * 12  '12 AS PHONE$,
     Meter AS STRING * 12  '12 AS Meter$ 'this is the actual account i think
     WRate AS INTEGER      '2 AS WRATE$
     ERate AS INTEGER      '2 AS ERATE$
     InCtyLmt AS STRING * 1 '1 AS ctylimit$,
     CustType AS STRING * 1  '1 AS custype$,
     WAvgUse AS LONG       '4 AS WAVGUSE$
     WAvgCnt AS INTEGER   '2 AS wavgct$,
     CurrBal AS SINGLE     '4 AS CURBAL$,
     PastBal AS SINGLE    '4 AS Pastdue$
     '''''GOOD to here FIELD 1, 181 AS dum1$
     Active AS STRING * 1  ' 1 AS active$
     WCurrRead AS LONG     ' 4 AS wcurread$,
     WPastRead AS LONG     ' 4 AS wpastread$,
     ECurrRead AS LONG     '4 AS ecurread$,
     EPastRead AS LONG     '4 AS epastread$,
     FirstTR AS LONG       '4 AS FIRSTTR$,
     LastTR AS LONG        '4 AS LASTTR$,
     WaterDep AS SINGLE    '4 AS WATERDEP$, 'Check
     ElecDep AS SINGLE     '4 AS ELECDEP$,
     IRev1 AS SINGLE      '4 AS IREVENUE1$,
     IRev2 AS SINGLE      '4 AS IREVENUE2$,
     IRev3 AS SINGLE      '4 AS IREVENUE3$,
     IRev4 AS SINGLE      '4 AS IREVENUE4$,
     IRev5 AS SINGLE      '4 AS IREVENUE5$
     '''''GOOD to here FIELD 1, 234 AS dum2$,
     IRev6 AS SINGLE      '4 AS IREVENUE6$,
     RAmt1 AS SINGLE      '4 AS ramt1$,
     RAmt2 AS SINGLE      '4 AS ramt2$,
     RAmt3 AS SINGLE      '4 AS ramt3$,
     RDesc1 AS STRING * 25  '25 AS RDESC1$,
     RDesc2 AS STRING * 25  '25 AS RDESC2$,
     RDesc3 AS STRING * 25  '25 AS RDESC3$,
     RAcct1 AS STRING * 2   '2 AS racct1$,''source% = VAL(racct1$)
     RAcct2 AS STRING * 2   '2 AS racct2$,
     RAcct3 AS STRING * 2   '2 AS racct3$,
     OAmt1 AS SINGLE       '4 AS oamt1$,
     OAmt2 AS SINGLE       '4 AS oamt2$,
     ODesc1 AS STRING * 25 '25 AS ODESC1$,
     ODesc2 AS STRING * 25 '25 AS ODESC2$,
     OAcct1 AS STRING * 2   '2 AS oacct1$, source% = VAL(oacct1$)
     OAcct2 AS STRING * 2   '2 AS oacct2$
     '''''GOOD to here. FIELD 1, 393 AS dum3$,
     WMin AS STRING * 2      '2 AS WMIN$, Unknown
     EMult AS INTEGER '2 AS emult$, Electric meter multiplier
     WMeterNum AS STRING * 4   '4 AS WMETERNUM$,
     EMeterNum AS STRING * 4   '4 AS EMETERNUM$,
     Tax AS STRING * 1         '1 AS tax$, if 'Y" exempt unless CityLimit="O"
     'IF tax$ = "Y" THEN TAXAMOUNT = 0
     'IF ctylimit$ = "O" THEN TAXAMOUNT = calcbill * .1
     Secur1 AS INTEGER         '2 AS Security1$,
     Secur2 AS INTEGER         '2 AS Security2$,
     FAdj AS STRING * 4        '4 AS fadj$,
     EAvgUse AS LONG           '4 AS eavguse$,
     EAvgCnt AS INTEGER        '2 AS eavgct$,
     Demand AS SINGLE          '4 AS demand$,
     IRev7 AS SINGLE          '4 AS irevenue7$, Tax revenue
     IRev8 AS SINGLE          '4 AS irevenue8$,
     GCurrRead AS LONG        '4 AS gcurread$,
     GPastRead AS LONG        '4 AS gpastread$,
     MSeq AS STRING * 4       '4 AS mseq$,
     GasDep AS SINGLE         '4 AS GASDEP$,
     Fill64  AS STRING * 64   '64 AS NOTUSED$
     'Good 512
END TYPE

