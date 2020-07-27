TYPE Cust71RecType  'WBCUST.DAT  256 FILE LAYOUT
    lname        AS STRING * 25
    firstname    AS STRING * 25
    add1         AS STRING * 30
    add2         AS STRING * 30
    city         AS STRING * 20
    st           AS STRING * 2
    zip          AS STRING * 9
    PHONE        AS STRING * 12

'    zip          AS STRING * 5
'    PHONE        AS STRING * 12
'    Zip1         AS STRING * 4

    book         AS INTEGER
    meter        AS STRING * 12  '
    rate        AS INTEGER
    won         AS STRING * 8
    MINSTALL    AS STRING * 8
    AVGUSE       AS DOUBLE       'avg use
    AVGCT        AS INTEGER      'counter for number of times read
    CURBAL       AS DOUBLE
    pastdue      AS DOUBLE
    cycle        AS STRING * 1   'most will be the act status

'    Zip2         AS STRING * 4
'    nu1          AS STRING * 1

    NU           AS STRING * 5   'garb for lawrence
    curread      AS DOUBLE
    pastread     AS DOUBLE
    firsttr      AS DOUBLE       'cust trans pointers
    lasttr       AS DOUBLE       'cust trans pointers
    waterdep     AS SINGLE       'Water deposit
    ' BCycle       AS STRING * 1
END TYPE

TYPE WBCust1RecType
    RAMT1     AS SINGLE
    RAMT2     AS SINGLE
    RAMT3     AS SINGLE
    RDESC1     AS STRING * 25
    RDESC2     AS STRING * 25
    RDESC3     AS STRING * 25
    RACCT1    AS STRING * 9
    RACCT2    AS STRING * 9
    RACCT3    AS STRING * 9
    OAMT1     AS SINGLE
    OAMT2     AS SINGLE
    ODESC1    AS STRING * 25
    ODESC2    AS STRING * 25
    OACCT1    AS STRING * 9
    OACCT2    AS STRING * 9
    SewerRate AS INTEGER
    custtype AS INTEGER
    METERNUM AS STRING * 14
    MeterSeq AS SINGLE
    HHMess AS STRING * 30

    'Minimum   AS INTEGER
    'GARBAGE AS INTEGER
    'Add3 AS STRING * 30
    'Add3 AS STRING * 28
    'Units  AS INTEGER
END TYPE

TYPE WBCust2RecType
    REV1    AS DOUBLE
    REV2    AS DOUBLE
    REV3    AS DOUBLE
    REV4    AS DOUBLE
    REV5    AS DOUBLE
    REV6    AS DOUBLE
    REV7    AS DOUBLE
    REV8    AS DOUBLE
END TYPE

TYPE OwnerRecType
 Owner AS STRING * 25
 Add4 AS STRING * 25
 Info AS STRING * 25
 Rental AS STRING * 30
END TYPE


TYPE nsumDataType
  BD AS STRING * 8        'Billing Date
  CR AS DOUBLE            'Current Reading
  PR AS DOUBLE            'Past read
  BT AS INTEGER           'Bill total
END TYPE
    
TYPE WBTransRecType      ' 50 BYTE FILE FORMAT
   TDATE  AS STRING * 8  ' Transaction a Double??
   TAMT   AS DOUBLE      ' Total Balance
   TTYPE  AS INTEGER     ' Code trans type
   TDESC  AS STRING * 15 ' trans desc
   TMREC  AS SINGLE      ' Customer rec num
   TNEXTTR AS SINGLE     ' Pointer to next Trans Rec
   tcuramt  AS SINGLE     ' Current Trans Amt
END TYPE

'TYPE Cust71RecType  'WBCUST.DAT  512 FILE LAYOUT
'    LNAME        AS STRING * 25
'    FIRSTNAME    AS STRING * 25
'    add1         AS STRING * 30
'    add2         AS STRING * 30
'    city         AS STRING * 20
'    st           AS STRING * 2
'    zip          AS STRING * 5
'    PHONE        AS STRING * 12
'    meter        AS STRING * 12  '
'    WRATE        AS INTEGER
'    ERATE        AS INTEGER
'    CTYLIMIT     AS STRING * 1
'    CUSTYPE      AS STRING * 1
'    AVGUSE       AS SINGLE       'avg use
'    AVGCT        AS INTEGER      'counter for number of times read
'    CURBAL       AS SINGLE
'    Pastdue      AS SINGLE
'    cycle        AS STRING * 1   'most will be the act status
'    WCurRead     AS LONG
'    WPastRead    AS LONG
'    ecurread     AS LONG
'    epastread    AS LONG
'    FIRSTTR      AS SINGLE       'cust trans pointers
'    LASTTR       AS SINGLE       'cust trans pointers
'    WATERDEP     AS SINGLE       'Water deposit
'    ELECDEP      AS SINGLE
'    REV1    AS SINGLE
'    REV2    AS SINGLE
'    REV3    AS SINGLE
'    REV4    AS SINGLE
'    REV5    AS SINGLE
'    REV6    AS SINGLE
'    ramt1    AS SINGLE
'    ramt2    AS SINGLE
'    ramt3    AS SINGLE
'    RDes1    AS STRING * 25
'    RDes2    AS STRING * 25
'    RDes3    AS STRING * 25
'    racct1   AS STRING * 2
'    racct2   AS STRING * 2
'    racct3   AS STRING * 2
'    oamt1 AS SINGLE
'    oamt2 AS SINGLE
'    ODESC1 AS STRING * 25
'    ODESC2 AS STRING * 25
'    oacct1 AS STRING * 2
'    oacct2 AS STRING * 2
'    WMIN AS INTEGER
'    emult AS INTEGER
'    WMETERNUM AS SINGLE
'    EMETERNUM AS SINGLE
'    sewerrate AS STRING * 1
'    Security1 AS INTEGER
'    Security2 AS INTEGER
'    fadj    AS SINGLE
'    eavguse AS SINGLE
'    eavgct AS INTEGER
'    demand AS SINGLE
'    ADD3 AS STRING * 25
'    numbermeters AS INTEGER
'    wcurread1 AS LONG
'    wpastread1 AS LONG
'    wcurread2 AS LONG
'    wpastread2 AS LONG
'    wcurread3 AS LONG
'    wpastread3 AS LONG
'    lmeterdate AS STRING * 8
'    cmeterdate AS STRING * 8
'    acctopen AS STRING * 6
'    runits AS INTEGER
'    NotUsed AS STRING * 13
'END TYPE

CLOSE 1
: OPEN "wbcust.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 256

TYPE Cust71RecType  'WBCUST.DAT  312 FILE LAYOUT
    lname        AS STRING * 25
    firstname    AS STRING * 25
    add1         AS STRING * 30
    add2         AS STRING * 30
    city         AS STRING * 20
    st           AS STRING * 2
    zip          AS STRING * 9
    PHONE        AS STRING * 12
    book         AS INTEGER
    meter        AS STRING * 12  '
    rate        AS INTEGER
    won         AS STRING * 8
    MINSTALL    AS STRING * 8
    AVGUSE       AS DOUBLE       'avg use
    AVGCT        AS INTEGER      'counter for number of times read
    CURBAL       AS DOUBLE
    pastdue      AS DOUBLE
    cycle        AS STRING * 1   'most will be the act status
    mtrreadflag  AS STRING * 1
    pad1         as string * 4
    curread      AS DOUBLE
    pastread     AS DOUBLE
    firsttr      AS DOUBLE       'cust trans pointers
    lasttr       AS DOUBLE       'cust trans pointers
    waterdep     AS SINGLE       'Water deposit
    ' BCycle       AS STRING * 1
END TYPE

25 AS lname$
25 AS firstname$
30 AS add1$
30 AS add2$
20 AS city$
2 AS st$
9 AS zp$
12 AS PHONE$
2 AS book$
12 AS meter$
2 AS rate$
8 AS won$
8 AS MINSTALL$
8 AS AVGUSE$
2 AS AVGCT$
8 AS CURBAL$
8 AS pastdue$
1 AS cycle$
1  AS mtrreadflag$
4 AS NOTUSED$

8 AS curread$
8 AS pastread$
8 AS firsttr$
8 AS lasttr$
4 AS waterdep$




















CLOSE 9
OPEN "wbcust2.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #9 LEN = 64

8 AS irevenue1$
8 AS irevenue2$
8 AS irevenue3$
8 AS irevenue4$
8 AS irevenue5$
8 AS irevenue6$
8 AS irevenue7$
8 AS irevenue8$

CLOSE 6
OPEN "wbcust1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #6 LEN = 312

4 AS RAMT1$
4 AS RAMT2$
4 AS RAMT3$
25 AS RDESC1$
25 AS RDESC2$
25 AS RDESC3$

'87 AS dum1$
9 AS RACCT1$
9 AS RACCT2$
9 AS RACCT3$
4 AS OAMT1$
4 AS OAMT2$

'122 AS dum2$
25 AS ODESC1$
25 AS ODESC2$
9 AS OACCT1$
9 AS OACCT2$

'190 AS dum3$
2 AS custtype$
14 AS METERNUM$
30 AS ADD3$
2 AS minimums$

'238 AS dum4$
8 AS gcurread$
8 AS gpastread$
1 AS garbage$
1 AS sewer$

'256 AS dum5$
2 AS gminimums$
1 AS grate$
1 AS utilityfee$
4 AS penaltyamt$
1 AS exempt$
1 AS gmtrreadflag$
46 AS blk1$


