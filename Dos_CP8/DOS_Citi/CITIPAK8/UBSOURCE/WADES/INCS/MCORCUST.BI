TYPE WBMCCoreCustRecType
    LastName       AS STRING * 25        ',25 AS LNAME$        , 25 AS lname$
    FirstName      AS STRING * 25        ',25 AS FIRSTNaME$    , 25 AS firstnam
    Add1           AS STRING * 30        ', 30 AS add1$        , 30 AS ADD1$
    Add2           AS STRING * 30        ', 30 AS ADD2$        , 30 AS ADD2$
    City           AS STRING * 20        ', 20 AS city$        , 20 AS city$
    State          AS STRING * 2         ', 2 AS st$           , 2 AS ST$
    Zip            AS STRING * 5         ', 5 AS zp$           , 5 AS zp$
    Phone          AS STRING * 12        ', 12 AS PHONE$       , 12 AS phone$
    Meter          AS STRING * 12        ', 12 AS meter$       , 12 AS meter$
    WRate          AS INTEGER            ', 2 AS wrate$        , 2 AS wrate$
    ERate          AS INTEGER            ' , 2 AS erate$       , 2 AS erate$
    CityLimit      AS STRING * 1         ' , 1 AS ctylimit$    , 1 AS ctylimit$
    CustType       AS STRING * 1         ' , 1 AS custype$     , 1 AS CUSTYPE$
    WAvgUse        AS LONG               ' , 4 AS WAVGUSE$     , 4 AS WAVGUSE$
    WAvgCnt        AS INTEGER            ' , 2 AS wavgct$      , 2 AS wavgct$
    CurBal         AS SINGLE             ', 4 AS CURBAL$       , 4 AS CURBAL$
    PastDue        AS SINGLE             ' , 4 AS PASTDUE$     , 4 AS PASTDUE$
'Field 1, 181 AS dum1$
    Active         AS STRING * 1         ' , 1 AS active$      , 1 AS active$
    WCurRead       AS LONG               ', 4 AS wcurread$     , 4 AS wcurread$
    WPastRead      AS LONG               ' , 4 AS wpastread$   , 4 AS wpastread$
    ECurRead       AS LONG               ', 4 AS ecurread$     , 4 AS ecurread$
    EPastRead      AS LONG               ' , 4 AS epastread$   , 4 AS epastread$

    FirstTrans     AS LONG               ' , 4 AS FIRSTTR$     , 4 AS FIRSTTR$
    LastTrans      AS LONG               ' , 4 AS LASTTR$      , 4 AS LASTTR$

    WaterDep       AS SINGLE             ' , 4 AS WATERDEP$    , 4 AS WATERDEP$
    ElecDep        AS SINGLE             ' , 4 AS ELECDEP$     , 4 AS ELECDEP$
    IRevenue1      AS SINGLE             ', 4 AS irevenue1$    , 4 AS IREVENUE1$
    IRevenue2      AS SINGLE             ' , 4 AS IREVENUE2$   , 4 AS IREVENUE2$
    IRevenue3      AS SINGLE             ' , 4 AS irevenue3$   , 4 AS IREVENUE3$
    IRevenue4      AS SINGLE             ' , 4 AS irevenue4$   , 4 AS IREVENUE4$
    IRevenue5      AS SINGLE             ' , 4 AS irevenue5$   , 4 AS irevenue5$

'FIELD 1 , 234 AS dum2$
    IRevenue6      AS SINGLE             ', 4 AS irevenue6$    , 4 AS irevenue6$
    RAmt1          AS SINGLE             ' , 4 AS ramt1$       , 4 AS ramt1$
    RAmt2          AS SINGLE             ' , 4 AS ramt2$       , 4 AS ramt2$
    RAmt3          AS SINGLE             ' , 4 AS ramt3$       , 4 AS ramt3$
    RDesc1         AS STRING * 25        ' , 25 AS RDESC1$     , 25 AS RDESC1$
    RDesc2         AS STRING * 25        ' , 25 AS RDESC2$     , 25 AS RDESC2$
    RDesc3         AS STRING * 25        ' , 25 AS RDESC3$     , 25 AS RDESC3$
    RAcct1         AS INTEGER             ', 2 AS racct1$      , 2 AS racct1$
    RAcct2         AS INTEGER             ', 2 AS racct2$      , 2 AS racct2$
    RAcct3         AS INTEGER             ', 2 AS racct3$      , 2 AS racct3$
    OAmt1          AS SINGLE              ', 4 AS oamt1$       , 4 AS oamt1$
    OAmt2          AS SINGLE              ', 4 AS oamt2$       , 4 AS oamt2$
    ODesc1         AS STRING * 25         ', 25 AS ODESC1$     , 25 AS ODESC1$
    ODesc2         AS STRING * 25         ', 25 AS ODESC2$     , 25 AS ODESC2$
    OAcct1         AS STRING * 2          ' , 2 AS oacct1$     , 2 AS oacct1$
    OAcct2         AS STRING * 2          ' , 2 AS oacct2$     , 2 AS oacct2$
'FIELD 1, 393 AS dum3$
    WMin           AS INTEGER             ' , 2 AS wmin$       , 2 AS WMIN$
    EMulti         AS INTEGER             ' , 2 AS emult$      , 2 AS emult$
    WMeterNum      AS LONG                ' , 4 AS WMETERNUM$  , 4 AS WMETERNUM$
    EMeterNum      AS LONG                ' , 4 AS EMETERNUM$  , 4 AS EMETERNUM$
    Tax            AS STRING * 1          ' , 1 AS tax$        , 1 AS Tax$
    Security1      AS INTEGER             ', 2 AS security1$   , 2 AS Security1$
    Security2      AS INTEGER             ' , 2 AS security2$  , 2 AS Security2$
    SMin           AS SINGLE              ' , 4 AS sewermin$   , 4 AS sewermin$
    EAvgUse        AS LONG                ' , 4 AS eavguse$    , 4 AS eavguse$
    EAvgCnt        AS INTEGER             ' , 2 AS eavgct$     , 2 AS eavgct$
    SRate          AS SINGLE              ' , 4 AS sewerrate$  , 4 AS sewerrate$
    IRevenue7      AS SINGLE              ' , 4 AS irevenue7$  , 4 AS IREVENUE7$
    IRevenue8      AS SINGLE              ' , 4 AS irevenue8$  , 4 AS IREVENUE8$
    IRevenue9      AS SINGLE              ' , 4 AS IREVENUE9$  , 4 AS irevenue9$
    IRevenue10     AS SINGLE              ' , 4 AS irevenue10$ , 4 AS irevenue10$
    EMtrNum        AS DOUBLE              '                    , 8 AS nEMETERNUM$
    DraftNum       AS SINGLE              '                    , 4 AS DRAFTNUMBER$
    fill           AS STRING * 60         ' , 72 AS NOTUSED$   , 60 AS NOTUSED$
END TYPE

TYPE WBTransType
  TransDate   AS STRING * 8   'Trans date
  TransAmt    AS DOUBLE       'total amt Past + Curr
  TransType   AS INTEGER      '1=Charge,2=Payment,3=Pen'ty/Recon,4=Adjust
  TransDesc   AS STRING * 15  'Transaction description
  CustRecNum  AS SINGLE       'cust acct  rec
  NextTrans   AS SINGLE       'Pointer to next trans
  TCurrAmt    AS SINGLE       'current amt
  RevAmt1(1 TO 10) AS SINGLE  'Revenues 1 to 10
  PayCode     AS INTEGER      'Paycode Check, cash
  PayDate     AS STRING * 6   'Payment Date
  RevAmt2(1 TO 2) AS SINGLE   'Revenues 11 & 12
END TYPE

