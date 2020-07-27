TYPE Cust72RecType  'WBCUST.DAT  512 FILE LAYOUT
    LNAME        AS STRING * 25
    FIRSTNAME    AS STRING * 25
    add1         AS STRING * 30
    add2         AS STRING * 30
    city         AS STRING * 20
    st           AS STRING * 2
    zip          AS STRING * 5
    PHONE        AS STRING * 12
    meter        AS STRING * 12  '
    wrate        AS INTEGER
    erate        AS INTEGER
    citylimit    AS STRING * 1
    Custype      AS STRING * 1
    WAvgUse   AS SINGLE
    WAvgCt    AS INTEGER
    CurBal    AS SINGLE
    PastDue   AS SINGLE
    Active    AS STRING * 1
    WCurRead  AS LONG
    WPastRead AS LONG
    ECurRead  AS LONG
    EPastRead AS LONG
    FirstTr   AS SINGLE
    LastTr    AS SINGLE
    WaterDep  AS SINGLE
    ElecDep   AS SINGLE
    Irev1 AS SINGLE
    Irev2 AS SINGLE
    Irev3 AS SINGLE
    Irev4 AS SINGLE
    Irev5 AS SINGLE
    Irev6 AS SINGLE
    Ramt1 AS SINGLE
    Ramt2 AS SINGLE
    Ramt3 AS SINGLE
    RDesc1 AS STRING * 25
    RDesc2 AS STRING * 25
    RDesc3 AS STRING * 25
    Racct1 AS STRING * 2
    Racct2 AS STRING * 2
    Racct3 AS STRING * 2
    Oamt1 AS SINGLE
    Oamt2 AS SINGLE
    ODesc1 AS STRING * 25
    ODesc2 AS STRING * 25
    Oacct1 AS STRING * 2
    Oacct2 AS STRING * 2
    WMin AS INTEGER
    EMult AS INTEGER
    WMeterNum AS LONG
    EMeterNum AS LONG
    Tax AS STRING * 1
    Security1 AS INTEGER
    Security2 AS INTEGER
    SewerMin AS SINGLE
    EAvgUse AS SINGLE
    EavgCt AS INTEGER
    SewerRate AS SINGLE
    Irev7 AS SINGLE
    irev8 AS SINGLE
    irev9 AS SINGLE
    irev10 AS SINGLE
    notuse AS STRING * 72
END TYPE
TYPE WBTransRecType      ' 128 BYTE FILE FORMAT
   TDATE  AS STRING * 8  ' Transaction Date
   TAMT   AS DOUBLE      ' Total Balance
   TTYPE  AS INTEGER     ' Code trans type
   TDESC  AS STRING * 15 ' trans desc
   tmrec  AS SINGLE      ' Customer rec num
   tnexttr AS SINGLE     ' Pointer to next Trans Rec
   tcuramt  AS SINGLE     ' Current Trans Amt
   wbrev1 AS SINGLE
   wbrev2 AS SINGLE
   wbrev3 AS SINGLE
   wbrev4 AS SINGLE
   wbrev5 AS SINGLE
   wbrev6 AS SINGLE
   wbrev7 AS SINGLE
   wbrev8 AS SINGLE
   wbrev9 AS SINGLE
   wbrev10 AS SINGLE
   pdcode AS INTEGER
   pddate AS STRING * 6
   wbrev11 AS SINGLE
   wbrev12 AS SINGLE
END TYPE


