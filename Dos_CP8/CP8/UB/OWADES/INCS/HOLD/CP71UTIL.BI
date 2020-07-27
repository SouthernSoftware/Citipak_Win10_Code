TYPE WBNoteType
  note1 AS STRING * 20
  note2 AS STRING * 20
  note3 AS STRING * 20
END TYPE


 TYPE WBCustRecType  'WBCUST.DAT  256 FILE LAYOUT
    LNAME        AS STRING * 25
    FIRSTNAME    AS STRING * 25
    add1         AS STRING * 30
    add2         AS STRING * 30
    CITY         AS STRING * 20
    ST           AS STRING * 2
    zip          AS STRING * 9
    phone        AS STRING * 12
    BOOK         AS INTEGER
    meter        AS STRING * 12  '
    rate        AS INTEGER
    won         AS STRING * 8
    MINSTALL    AS STRING * 8
    AVGUSE       AS DOUBLE       'avg use
    AVGCT        AS INTEGER      'counter for number of times read
    CURBAL       AS DOUBLE
    pastdue      AS DOUBLE
    STATUS        AS STRING * 1   'most will be the act status
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
 SewerRate    AS INTEGER
 CUSTTYPE     AS INTEGER
 METERNUM     AS STRING * 14
 add3         AS STRING * 30
 SprRate      AS INTEGER
 Minimums     AS INTEGER
 END TYPE
 
  
TYPE WBCustRecType3
       wateron AS STRING * 6
       wateroff AS STRING * 6
       mess1 AS STRING * 20
       mess2 AS STRING * 20
       mess3 AS STRING * 20
       'memfee AS SINGLE
       'owner AS STRING * 30
       'oadd AS STRING * 30
       'oadd2 AS STRING * 30
       'odeposit AS SINGLE
       'MESSFIRST AS SINGLE
       'MESSLAST AS SINGLE
      'spacer AS STRING * 20
      'MeterReadDate AS STRING * 6
END TYPE


TYPE OwnerRecType
 owner AS STRING * 25
 Add4 AS STRING * 25
 Info AS STRING * 25
 Rental AS STRING * 30
END TYPE



TYPE Cust71RecType1
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
    
END TYPE
TYPE Cust71RecType  'WBCUST.DAT  512 FILE LAYOUT
  LNAME     AS STRING * 25
  FIRSTNAME AS STRING * 25
  add1      AS STRING * 30
  add2      AS STRING * 30
  CITY      AS STRING * 20
  ST        AS STRING * 2
  ZP        AS STRING * 9
  phone     AS STRING * 12
  BOOK      AS STRING * 2
  meter     AS STRING * 12
  won       AS STRING * 8
  MINSTALL  AS STRING * 8
  AVGUSE    AS DOUBLE
  AVGCT     AS INTEGER
  CURBAL    AS DOUBLE
  pastdue   AS DOUBLE
  cycle     AS STRING * 1
  CURREAD   AS DOUBLE
  PASTREAD  AS DOUBLE
  firsttr   AS DOUBLE
  lasttr    AS DOUBLE
  waterdep  AS SINGLE
  irev1     AS DOUBLE
  irev2     AS DOUBLE
  irev3     AS DOUBLE
  irev4     AS DOUBLE
  irev5     AS DOUBLE
  irev6     AS DOUBLE
  irev7     AS DOUBLE
  IREV8     AS DOUBLE
  CUSTTYPE  AS STRING * 2
  METERNUM  AS STRING * 14
  add3      AS STRING * 30
  Minimums  AS INTEGER
  Rate1     AS STRING * 2
  Rate2     AS STRING * 2
  Rate3     AS STRING * 2
  Rate4     AS STRING * 2
  Rate5     AS STRING * 2
  Rate6     AS STRING * 2
  Rate7     AS STRING * 2
  Rate8     AS STRING * 2
  MeterReadDate AS STRING * 6
  MeterEst  AS STRING * 1
  WorkOrder AS STRING * 1
  WorkOrderNum AS SINGLE
  WorOrdBeg AS STRING * 6
  WorOrdEnd AS STRING * 6
  MESSFIRST AS SINGLE
  MESSLAST  AS SINGLE
  Comment   AS STRING * 30
  SeqNum    AS SINGLE
  CustId    AS STRING * 10
  ProRate   AS STRING * 1
  MtrTime   AS STRING * 6

END TYPE



