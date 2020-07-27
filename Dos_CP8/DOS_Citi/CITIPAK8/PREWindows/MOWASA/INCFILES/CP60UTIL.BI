
TYPE WBCustRecType  'WBCUST.DAT  256 FILE LAYOUT
    LNAME        AS STRING * 25
    FIRSTNAME    AS STRING * 25
    ADD1         AS STRING * 30
    add2         AS STRING * 30
    CITY         AS STRING * 20
    ST           AS STRING * 2
    zip          AS STRING * 5
    Phone        AS STRING * 12
    acct         AS SINGLE
    BOOK         AS INTEGER
    meter        AS STRING * 12  '
    RATE        AS INTEGER
    WON         AS STRING * 8
    MINSTALL    AS STRING * 8
    AVGUSE       AS DOUBLE       'avg use
    AVGCT        AS INTEGER      'counter for number of times read
    CURBAL       AS DOUBLE
    PASTDUE      AS DOUBLE
    cycle        AS STRING * 1
    NOTUSED      AS STRING * 5
    CURREAD      AS DOUBLE
    PASTREAD     AS DOUBLE
    firsttr      AS DOUBLE       'cust trans pointers
    lasttr       AS DOUBLE       'cust trans po
    waterdep     AS SINGLE       'Water deposit
END TYPE


 
 
TYPE WBCustRecType1
RAMT1        AS SINGLE
RAMT2        AS SINGLE
RAMT3        AS SINGLE
RDESC1       AS STRING * 25
RDESC2       AS STRING * 25
RDESC3       AS STRING * 25
RACCT1       AS STRING * 9
RACCT2       AS STRING * 9
RACCT3       AS STRING * 9
OAMT1        AS SINGLE
OAMT2        AS SINGLE
ODESC1       AS STRING * 25
ODESC2       AS STRING * 25
OACCT1       AS STRING * 9
OACCT2       AS STRING * 9
Minimums     AS INTEGER
CUSTTYPE     AS INTEGER
METERNUM     AS STRING * 14
add3         AS STRING * 30
END TYPE


TYPE WBCustRecType2
    REV1    AS DOUBLE
    REV2    AS DOUBLE
    REV3    AS DOUBLE
    REV4    AS DOUBLE
    REV5    AS DOUBLE
    REV6    AS DOUBLE
    rev7    AS DOUBLE
    rev8    AS DOUBLE
END TYPE


TYPE WBCustRecType3
      wateron AS STRING * 6
      wateroff AS STRING * 6
      mess1 AS STRING * 20
      mess2 AS STRING * 20
      mess3 AS STRING * 20
      MEMFEE AS SINGLE
      owner AS STRING * 30
      oadd AS STRING * 30
      oadd2 AS STRING * 30
      odeposit AS SINGLE
      MESSFIRST AS SINGLE
      MESSLAST AS SINGLE
      lastpaid AS STRING * 6
      lastamt AS DOUBLE
END TYPE

TYPE OwnerRecType
 owner AS STRING * 25
 add4 AS STRING * 25
 info AS STRING * 25
 rental AS STRING * 30
END TYPE





