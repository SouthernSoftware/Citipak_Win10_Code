TYPE CustRecType  'WBCUST.DAT  512 FILE LAYOUT
    LNAME        AS STRING * 25     ' 25 AS lname$
    firstname    AS STRING * 25     ' 25 AS firstname$
    ADD1         AS STRING * 30     ' 30 AS add1$
    ADD2         AS STRING * 30     ' 30 AS add2$
    CITY         AS STRING * 20     ' 20 AS CITY$
    ST           AS STRING * 2      ' 2 AS ST$
    zip          AS STRING * 5      ' 5 AS ZP$
    PHONE        AS STRING * 12     ' 12 AS phone$
    meter        AS STRING * 12    ' 12 AS meter$
    wrate        AS INTEGER
    ERATE        AS INTEGER
    CityLimit    AS STRING * 1
    Custtype     AS STRING * 1
    WAVGUSE      AS LONG
    wavgct       AS INTEGER
    curbal       AS SINGLE
    PASTDUE      AS SINGLE
    active       AS STRING * 1
    CurRead      AS LONG
    PastRead     AS LONG
    ecurread     AS LONG
    epastread    AS LONG
    FIRSTTR      AS LONG
    LASTTR       AS LONG
    WATERDEP     AS SINGLE    ' 4 AS waterdep$    'Water deposit
    ELECDEP      AS SINGLE
    REV1    AS SINGLE
    REV2    AS SINGLE
    REV3    AS SINGLE
    REV4    AS SINGLE
    REV5    AS SINGLE
    REV6    AS SINGLE
    ramt1  AS SINGLE
    ramt2   AS SINGLE
    ramt3   AS SINGLE
    RDESC1  AS STRING * 25
    RDESC2  AS STRING * 25
    RDESC3  AS STRING * 25
    racct1  AS STRING * 2
    racct2  AS STRING * 2
    racct3  AS STRING * 2
    OAMT1   AS SINGLE
    OAMT2   AS SINGLE
    ODESC1  AS STRING * 25
    ODESC2  AS STRING * 25
    oacct1  AS STRING * 2
    oacct2  AS STRING * 2
    wmin   AS INTEGER
    emult  AS INTEGER
    WMETERNUM AS LONG
    EMETERNUM AS LONG
    sewerrate AS STRING * 1
    Security1 AS INTEGER
    Security2 AS INTEGER
    fadj      AS SINGLE
    eavguse   AS LONG
    eavgct    AS INTEGER
    demand    AS LONG
    ADD3      AS STRING * 25
    NUMBERMETERS AS INTEGER
    WCURREAD1 AS LONG
    WPASTREAD1 AS LONG
    WCURREAD2 AS LONG
    WPASTREAD2 AS LONG
    wcurread3 AS LONG
    WPastread3 AS LONG
    lmeterdate AS STRING * 8
    cmeterdate AS STRING * 8
    acctopen AS STRING * 6
    runits AS INTEGER
 END TYPE

TYPE WBMeter
  meter AS STRING * 12
  Record AS SINGLE
END TYPE



TYPE CustRecType1    'WBCUST.DAT  512 FILE LAYOUT
    LNAME        AS STRING * 25
    firstname    AS STRING * 25
    ADD1         AS STRING * 30
    ADD2         AS STRING * 30
    CITY         AS STRING * 20
    ST           AS STRING * 2
    zip          AS STRING * 9
    PHONE        AS STRING * 12
    BOOK         AS INTEGER
    meter        AS STRING * 12  '
    WON          AS STRING * 8
    MINSTALL     AS STRING * 8
    AVGUSE       AS DOUBLE       'avg use
    AVGCT        AS INTEGER      'counter for number of times read
    curbal       AS DOUBLE
    PASTDUE      AS DOUBLE
    cycle        AS STRING * 1   'most will be the act status
    CurRead      AS DOUBLE
    PastRead     AS DOUBLE
    FIRSTTR      AS DOUBLE       'cust trans pointers
    LASTTR       AS DOUBLE       'cust trans pointers
    WATERDEP     AS SINGLE       'Water deposit
    REV1    AS DOUBLE
    REV2    AS DOUBLE
    REV3    AS DOUBLE
    REV4    AS DOUBLE
    REV5    AS DOUBLE
    REV6    AS DOUBLE
    rev7    AS DOUBLE
    rev8    AS DOUBLE
    Custtype AS STRING * 2
    METERNUM AS STRING * 14
    ADD3  AS STRING * 30
    mINIMUMS AS INTEGER
    RATE1 AS STRING * 2
    RATE2 AS STRING * 2
    RATE3 AS STRING * 2
    RATE4 AS STRING * 2
    RATE5 AS STRING * 2
    RATE6 AS STRING * 2
    RATE7 AS STRING * 2
    RATE8 AS STRING * 2
    MeterReadDate AS STRING * 6
    METEREST AS STRING * 1
    WorkOrder AS STRING * 1
    WorkOrderNum AS SINGLE
    WorkOrdBeg AS STRING * 6
    WorkOrdEnd AS STRING * 6
    messfirst AS SINGLE
    messlast AS SINGLE
    comment AS STRING * 30
    SeqNum AS SINGLE
    CustID AS STRING * 10
    oLDProRate AS STRING * 1
    DepDate AS STRING * 6
    FirstWo AS SINGLE
    LastWo AS SINGLE
    MeterType AS STRING * 1
    NotUse AS STRING * 46
END TYPE

TYPE CustRecType2
    ramt1    AS SINGLE
    ramt2    AS SINGLE
    ramt3    AS SINGLE
    RDes1    AS STRING * 25
    RDes2    AS STRING * 25
    RDes3    AS STRING * 25
    racct1   AS STRING * 9
    racct2   AS STRING * 9
    racct3   AS STRING * 9
    OAMT1 AS SINGLE
    OAMT2 AS SINGLE
    ODESC1 AS STRING * 25
    ODESC2 AS STRING * 25
    oacct1 AS STRING * 9
    oacct2 AS STRING * 9
    NOTUSED AS STRING * 66
END TYPE


TYPE WBCUSTHHRECTYPE
   LINE1 AS STRING * 20
   LINE2 AS STRING * 20
   LINE3 AS STRING * 20
END TYPE

