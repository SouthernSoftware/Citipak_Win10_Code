 




TYPE WBNoteType
  note1 AS STRING * 25
  note2 AS STRING * 25
  note3 AS STRING * 25
  Note4 AS STRING * 30
END TYPE

 

TYPE WBCustRecType  'WBCUST.DAT  256 FILE LAYOUT
    LName        AS STRING * 25
    FirstName    AS STRING * 25
    add1         AS STRING * 30
    add2         AS STRING * 30
    CITY         AS STRING * 20
    ST           AS STRING * 2
    zip          AS STRING * 9
    phone        AS STRING * 12
    BOOK         AS INTEGER
    meter        AS STRING * 12  '
    rate         AS INTEGER
     won          AS STRING * 8
    MINSTALL     AS STRING * 8
    AVGUSE       AS DOUBLE       'avg use
    AVGCT        AS INTEGER      'counter for number of times read
    CURBAL       AS DOUBLE
    pastdue      AS DOUBLE
    Status       AS STRING * 1   'most will be the act status
    NOTUSED      AS STRING * 5
    curread      AS DOUBLE
    pastread     AS DOUBLE
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
 SRate        AS STRING * 2
 CUSTTYPE     AS INTEGER
 METERNUM     AS STRING * 14
 add3         AS STRING * 30
 SprRate      AS INTEGER
 minimums     AS INTEGER
END TYPE
      
TYPE WBCustRecType2
 DEPDATE AS STRING * 6
 PAYHIST AS STRING * 12
 meterseq AS SINGLE
 WRate AS STRING * 3
 IRate AS STRING * 3
 SRate AS STRING * 3
 billcycle AS INTEGER
 curread1 AS DOUBLE
 pastread1 AS DOUBLE
 Blank AS STRING * 9
 irrmeter AS STRING * 9
 curreaddate AS STRING * 6
 PastReadDate AS STRING * 6
 draftnumber AS SINGLE
 mtrtime AS STRING * 6
 State AS STRING * 25
 END TYPE



TYPE WBCustRecType3
       wateron AS STRING * 6
       wateroff AS STRING * 6
       mess1 AS STRING * 20
       mess2 AS STRING * 20
       mess3 AS STRING * 20
       memfee AS SINGLE
       owner AS STRING * 30
       oadd AS STRING * 30
       oadd2 AS STRING * 30
       'odeposit AS SINGLE
       MESSFIRST AS SINGLE
       MESSLAST AS SINGLE
      'spacer AS STRING * 20
      'MeterReadDate AS STRING * 6
END TYPE


TYPE OwnerRecType
 owner AS STRING * 25
 add4 AS STRING * 25
 info AS STRING * 25
 rental AS STRING * 30
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
  LName     AS STRING * 25
  FirstName AS STRING * 25
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
  curread   AS DOUBLE
  pastread  AS DOUBLE
  firsttr   AS DOUBLE
  lasttr    AS DOUBLE
  waterdep  AS SINGLE
  IRev1     AS DOUBLE
  IRev2     AS DOUBLE
  IRev3     AS DOUBLE
  IRev4     AS DOUBLE
  IRev5     AS DOUBLE
  IRev6     AS DOUBLE
  irev7     AS DOUBLE
  IREV8     AS DOUBLE
  CUSTTYPE  AS STRING * 2
  METERNUM  AS STRING * 14
  add3      AS STRING * 30
  minimums  AS INTEGER
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
  mtrtime   AS STRING * 6
  Etime     AS STRING * 8
  NProRate AS STRING * 3
  SSN      AS STRING * 9
  MOrder   AS STRING * 1
  OwnerCode AS STRING * 1
  RentalDep AS SINGLE
  AmtOwed AS SINGLE
  AmtPd AS SINGLE
  Pymt AS SINGLE
  RevS AS INTEGER
END TYPE

     
     

TYPE CustRecType7  'WBCUST.DAT  512 FILE LAYOUT LongView
  LName     AS STRING * 25
  FirstName AS STRING * 25
  add1      AS STRING * 30
  add2      AS STRING * 30
  CITY      AS STRING * 20
  ST        AS STRING * 2
  ZP        AS STRING * 5
  phone     AS STRING * 12
  meter     AS STRING * 12
  WRate     AS INTEGER
  ERate     AS INTEGER
  CtyLimit  AS STRING * 1
  Custype   AS STRING * 1
  AVGUSE    AS LONG
  AVGCT     AS INTEGER
  CURBAL    AS SINGLE
  pastdue   AS SINGLE
 
  Status    AS STRING * 1
  curread   AS LONG
  pastread  AS LONG
  ECurRead  AS LONG
  EPastRead AS LONG
  firsttr   AS SINGLE
  lasttr    AS SINGLE
  waterdep  AS SINGLE
  ElecDep   AS SINGLE

  IRev1     AS SINGLE
  IRev2     AS SINGLE
  IRev3     AS SINGLE
  IRev4     AS SINGLE
  IRev5     AS SINGLE
  IRev6     AS SINGLE
  RAMT1        AS SINGLE
 RAMT2        AS SINGLE
 RAMT3        AS SINGLE
 RDESC1       AS STRING * 25
 RDESC2       AS STRING * 25
 RDESC3       AS STRING * 25
 RACCT1       AS STRING * 2
 RACCT2       AS STRING * 2
 RACCT3       AS STRING * 2
 OAMT1        AS SINGLE
 OAMT2        AS SINGLE
 ODESC1       AS STRING * 25
 ODESC2       AS STRING * 25
 OACCT1       AS STRING * 2
 OACCT2       AS STRING * 2
 
 WMin         AS INTEGER
 EMult        AS INTEGER
 WMeterNum    AS SINGLE
 EMeterNum    AS SINGLE
 Recycle      AS STRING * 1
 Security1    AS INTEGER
 Security2    AS INTEGER
 Fadj         AS SINGLE
 EAvgUse      AS SINGLE
 Eavgct       AS INTEGER
 
 Demand       AS LONG
 add3         AS STRING * 25
 NumberMtr    AS INTEGER
 WCurRead     AS LONG
 WPastRead    AS LONG
 WCurRead1    AS LONG
 WPastRead1   AS LONG
 WCurRead2    AS LONG
 WPastRead2   AS LONG


END TYPE


TYPE WBDraftType
 bankname AS STRING * 30
 BankLoc AS STRING * 30
 BankTransit AS STRING * 9
 BankAcct AS STRING * 17
 Blank AS STRING * 4
 PreNote AS STRING * 1
END TYPE



