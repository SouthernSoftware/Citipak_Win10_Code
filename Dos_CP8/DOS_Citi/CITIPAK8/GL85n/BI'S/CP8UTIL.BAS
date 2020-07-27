TYPE TroyCust0RecType  'WBCUST.DAT
    lname        AS STRING * 25
    firstname    AS STRING * 25
    add1         AS STRING * 30
    add2         AS STRING * 30
    city         AS STRING * 20
    st           AS STRING * 2
    zip          AS STRING * 9
    PHONE        AS STRING * 12
    BOOK         AS INTEGER      'this is part of meter
    meter        AS STRING * 12  '
    RATE         AS INTEGER      'Rate 0-3
    WON          AS STRING * 8   'i or O     meter
    MINSTALL     AS STRING * 8   'Meter installed date
    AVGUSE       AS DOUBLE       'avg use
    AVGCT        AS INTEGER      'counter for number of times read
    curbal       AS DOUBLE       '
    pastdue      AS DOUBLE
    cycle        AS STRING * 1   'most will be the act status
    NOTUSED      AS STRING * 5
    curread      AS DOUBLE       'current
    pastread     AS DOUBLE
    firsttr      AS DOUBLE       'cust trans pointers
    lasttr       AS DOUBLE       'cust trans pointers
    waterdep     AS SINGLE       'Water deposit
    FILL1        AS STRING * 3
END TYPE

TYPE TroyCust1RecType       'WBCUST1.DAT
  RAMT1  AS SINGLE          'recuring amts
  RAMT2  AS SINGLE
  RAMT3  AS SINGLE
  RDESC1 AS STRING * 25     'Desc
  RDESC2 AS STRING * 25
  RDESC3 AS STRING * 25
  RACCT1 AS STRING * 9      'Rev Source
  RACCT2 AS STRING * 9      '
  RACCT3 AS STRING * 9
  OAMT1  AS SINGLE          'One time amts
  OAMT2  AS SINGLE
  ODESC1 AS STRING * 25
  ODESC2 AS STRING * 25
  OACCT1 AS STRING * 9      'Rev Source
  OACCT2 AS STRING * 9
  sewerrate      AS INTEGER       'Sewerrate Codes
  CUSTTYPE       AS STRING * 2    'Cust Type user defined
  METERNUM       AS STRING * 14   'Meter #
  add3           AS STRING * 30   'Address line 2
  sprinklerrate  AS INTEGER       'possibly blank
  minimums       AS INTEGER   'num of users
  FILL           AS STRING * 10 '
END TYPE

TYPE TroyCust2RecType  'WBCUST2.DAT
  IRev1   AS DOUBLE
  IRev2   AS DOUBLE
  IRev3   AS DOUBLE
  IRev4   AS DOUBLE   'rev amts
  IRev5   AS DOUBLE
  IRev6   AS DOUBLE
  irev7   AS DOUBLE
  irev8   AS DOUBLE
END TYPE

TYPE TroyCust3RecType  'WBCUST3.DAT
 wateron    AS STRING * 6  'date?
 wateroff   AS STRING * 6  'date?
 HHMESS1    AS STRING * 20  'hand held message
 HHMESS2    AS STRING * 20  'hand held message
 HHMESS3    AS STRING * 20  'hand held message
 FILL1      AS STRING * 56
END TYPE


TYPE TroyConsumDataType
  BD AS STRING * 8        'Billing Date
  CR AS DOUBLE            'Current Reading
  PR AS DOUBLE            'Past read
  BT AS INTEGER           'Bill total
END TYPE

TYPE TroyConsumRecType
  Consum(1 TO 12) AS TroyConsumDataType
END TYPE

TYPE TroyWBARTransRecType
   TDATE  AS STRING * 8  ' Transaction a Double??
   TAMT   AS DOUBLE      ' Total Balance
   TTYPE  AS INTEGER     ' Code trans type
   TDESC  AS STRING * 15 ' trans desc
   tmrec  AS SINGLE      ' Customer rec num
   tnexttr AS SINGLE     ' Pointer to next Trans Rec
   curamt  AS SINGLE     ' Current Trans Amt
END TYPE

