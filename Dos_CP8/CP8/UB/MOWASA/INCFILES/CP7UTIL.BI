TYPE Cust7RecType  'WBCUST.DAT
    LNAME        AS STRING * 25
    FIRSTNAME    AS STRING * 25
    add1         AS STRING * 30
    add2         AS STRING * 30
    city         AS STRING * 20
    st           AS STRING * 2
    zip          AS STRING * 9
    phone        AS STRING * 12
    BOOK         AS INTEGER      'this is part of meter
    meter        AS STRING * 12  '
    WON          AS STRING * 8   'i or O     meter
    MINSTALL     AS STRING * 8   'Meter installed date
    AVGUSE       AS DOUBLE       'avg use
    AVGCT        AS INTEGER      'counter for number of times read
    curbal       AS DOUBLE       '
    pastdue      AS DOUBLE
    cycle        AS STRING * 1   'most will be the act status
    curread      AS DOUBLE       'current
    pastread     AS DOUBLE
    firsttr      AS DOUBLE       'cust trans pointers
    lasttr       AS DOUBLE       'cust trans pointers
    waterdep     AS SINGLE       'Water deposit
    Irev1        AS DOUBLE
    Irev2        AS DOUBLE
    Irev3        AS DOUBLE
    Irev4        AS DOUBLE
    Irev5        AS DOUBLE
    Irev6        AS DOUBLE
    Irev7        AS DOUBLE
    Irev8        AS DOUBLE
    CustType     AS STRING * 2
    MeterNum    AS STRING * 14
    Add3        AS STRING * 30
    Minimum     AS INTEGER
    Rate1       AS STRING * 2
    Rate2       AS STRING * 2
    Rate3       AS STRING * 2
    Rate4       AS STRING * 2
    Rate5       AS STRING * 2
    Rate6       AS STRING * 2
    Rate7       AS STRING * 2
    Rate8       AS STRING * 2
    meterreaddate AS STRING * 6
    meterest AS STRING * 1
    workorder AS STRING * 1
    workordernumber AS STRING * 4
    workorderbeg AS STRING * 6
    workorderend AS STRING * 6
    messfirst AS SINGLE
    messlast AS SINGLE
    comment AS STRING * 30
    seqnum AS SINGLE
    CUSTID AS STRING * 10
    PRORATE AS STRING * 1
    depdate AS STRING * 6
    firstwo AS SINGLE
    lastwo AS SINGLE
    notused AS STRING * 47
 END TYPE

TYPE nsumDataType
  BD AS STRING * 8        'Billing Date
  CR AS DOUBLE            'Current Reading
  PR AS DOUBLE            'Past read
  BT AS INTEGER           'Bill total
END TYPE


TYPE WBTransRecType
   TDATE  AS STRING * 8  ' Transaction a Double??
   tamt   AS DOUBLE      ' Total Balance
   TTYPE  AS INTEGER     ' Code trans type
   TDESC  AS STRING * 15 ' trans desc
   tmrec  AS SINGLE      ' Customer rec num
   tnexttr AS SINGLE     ' Pointer to next Trans Rec
   curamt  AS SINGLE     ' Current Trans Amt
   wbrev1 AS SINGLE
   wbrev2 AS SINGLE
   wbrev3 AS SINGLE
   wbrev4 AS SINGLE
   wbrev5 AS SINGLE
   wbrev6 AS SINGLE
   wbrev7 AS SINGLE
   wbrev8 AS SINGLE
   TESTIMATE AS STRING * 1
   TCURREAD AS DOUBLE
   TPASTREAD AS DOUBLE
   TCONSUMP AS DOUBLE
   TBOOK AS INTEGER
   TCUSTTYPE AS STRING * 2
   TCUSTLOC AS STRING * 1
   
END TYPE

