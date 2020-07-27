
     

TYPE CustRecType  'WBCUST.DAT  512 FILE LAYOUT
    lname        AS STRING * 25     ' 25 AS lname$
    firstname    AS STRING * 25     ' 25 AS firstname$
    add1         AS STRING * 30     ' 30 AS add1$
    add2         AS STRING * 30     ' 30 AS add2$
    CITY         AS STRING * 20     ' 20 AS CITY$
    ST           AS STRING * 2      ' 2 AS ST$
    zip          AS STRING * 5      ' 5 AS ZP$
    phone        AS STRING * 12     ' 12 AS phone$
    meter        AS STRING * 12    ' 12 AS meter$
    WRATE        AS INTEGER
    ERATE        AS INTEGER
    GRate        AS STRING * 1
    CUSTTYPE     AS STRING * 1
    WAVGUSE      AS LONG
    wavgct       AS INTEGER
    CURBAL       AS SINGLE
    pastdue      AS SINGLE
    active       AS STRING * 1
    curread      AS LONG
    pastread     AS LONG
    ecurread     AS LONG
    epastread    AS LONG
    firsttr      AS LONG
    lasttr       AS LONG
    waterdep     AS SINGLE    ' 4 AS waterdep$    'Water deposit
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
    oamt1   AS SINGLE
    oamt2   AS SINGLE
    ODESC1  AS STRING * 25
    ODESC2  AS STRING * 25
    oacct1  AS STRING * 2
    oacct2  AS STRING * 2
    WMIN   AS INTEGER
    emult  AS INTEGER
    WMETERNUM AS LONG
    EMETERNUM AS LONG
    Taxable   AS STRING * 1
    Security1 AS INTEGER
    Security2 AS INTEGER
    fadj      AS SINGLE
    eavguse   AS LONG
    eavgct    AS INTEGER
    demand    AS LONG
    Rev7      AS SINGLE
    Rev8      AS SINGLE
 END TYPE

TYPE WBMeter
  meter AS STRING * 12
  Record AS SINGLE
END TYPE


      




TYPE CustRecType1    'WBCUST.DAT  512 FILE LAYOUT
    lname        AS STRING * 25
    firstname    AS STRING * 25
    add1         AS STRING * 30
    add2         AS STRING * 30
    CITY         AS STRING * 20
    ST           AS STRING * 2
    zip          AS STRING * 9
    phone        AS STRING * 12
    BOOK         AS INTEGER
    meter        AS STRING * 12  '
    WON          AS STRING * 8
    MINSTALL     AS STRING * 8
    AVGUSE       AS DOUBLE       'avg use
    AVGCT        AS INTEGER      'counter for number of times read
    CURBAL       AS DOUBLE
    pastdue      AS DOUBLE
    cycle        AS STRING * 1   'most will be the act status
    curread      AS DOUBLE
    pastread     AS DOUBLE
    firsttr      AS DOUBLE       'cust trans pointers
    lasttr       AS DOUBLE       'cust trans pointers
    waterdep     AS SINGLE       'Water deposit
    REV1    AS DOUBLE
    REV2    AS DOUBLE
    REV3    AS DOUBLE
    REV4    AS DOUBLE
    REV5    AS DOUBLE
    REV6    AS DOUBLE
    Rev7    AS DOUBLE
    Rev8    AS DOUBLE
    CUSTTYPE AS STRING * 2
    METERNUM AS STRING * 14
    add3  AS STRING * 30
    minimums AS INTEGER
    rate1 AS STRING * 2
    rate2 AS STRING * 2
    rate3 AS STRING * 2
    rate4 AS STRING * 2
    rate5 AS STRING * 2
    rate6 AS STRING * 2
    rate7 AS STRING * 2
    rate8 AS STRING * 2
    meterreaddate AS STRING * 6
    meterest AS STRING * 1
    workorder AS STRING * 1
    WorkOrderNum AS SINGLE
    WorkOrdBeg AS STRING * 6
    WorkOrdEnd AS STRING * 6
    messfirst AS SINGLE
    messlast AS SINGLE
    comment AS STRING * 30
    seqnum AS SINGLE
    custid AS STRING * 10
    oLDProRate AS STRING * 1
    BillCycle AS STRING * 1
    NotUsed AS STRING * 60
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
    oamt1 AS SINGLE
    oamt2 AS SINGLE
    ODESC1 AS STRING * 25
    ODESC2 AS STRING * 25
    oacct1 AS STRING * 9
    oacct2 AS STRING * 9
    NotUsed AS STRING * 66
END TYPE


TYPE WBCUSTHHRECTYPE
   LINE1 AS STRING * 20
   LINE2 AS STRING * 20
   LINE3 AS STRING * 20
END TYPE

