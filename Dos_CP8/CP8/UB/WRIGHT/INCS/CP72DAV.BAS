TYPE Cust72RecType  'WBCUST.DAT  512 FILE LAYOUT
    LNAME        AS STRING * 25
    FIRSTNAME    AS STRING * 25
    add1         AS STRING * 30
    add2         AS STRING * 30
    city         AS STRING * 20
    st           AS STRING * 2
    zip          AS STRING * 9
    PHONE        AS STRING * 12
    book         AS INTEGER
    meter        AS STRING * 12  '
    RATE         AS STRING * 2
    won          AS STRING * 8
    minstall     AS STRING * 8
    AVGUSE       AS DOUBLE       'avg use
    AVGCT        AS INTEGER      'counter for number of times read
    CURBAL       AS DOUBLE
    Pastdue      AS DOUBLE
    cycle        AS STRING * 1   'most will be the act status
    NUSE         AS STRING * 5
    CURREAD      AS DOUBLE
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
    sEWERrATE AS STRING * 2
    CustType AS INTEGER
    MeterNum AS STRING * 14
    Add3  AS STRING * 30
    dHEC AS INTEGER
    mINIMUMS AS INTEGER
    rate1 AS STRING * 2
    gRATE AS STRING * 2
END TYPE
TYPE Cust72RecType1
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
   testimate AS STRING * 1
   tcurread AS DOUBLE
   tpastread AS DOUBLE
   tconsump AS DOUBLE
   tbook AS INTEGER
   tcusttype AS STRING * 2
   tcustloc AS STRING * 1
   tstatus AS STRING * 1
   tblank AS STRING * 20
END TYPE

TYPE WBCUSTHHRECTYPE
   LINE1 AS STRING * 20
   LINE2 AS STRING * 20
   LINE3 AS STRING * 20
END TYPE

