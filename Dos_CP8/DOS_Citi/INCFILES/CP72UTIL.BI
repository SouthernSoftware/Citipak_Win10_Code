

TYPE CustRecType  'WBCUST.DAT  512 FILE LAYOUT
    lname        AS STRING * 25     ' 25 AS lname$
    firstname    AS STRING * 25     ' 25 AS firstname$
    ADD1         AS STRING * 30     ' 30 AS add1$
    ADD2         AS STRING * 30     ' 30 AS add2$
    city         AS STRING * 20     ' 20 AS CITY$
    st           AS STRING * 2      ' 2 AS ST$
    zip          AS STRING * 5      ' 5 AS ZP$
    phone        AS STRING * 12     ' 12 AS phone$
    meter        AS STRING * 12    ' 12 AS meter$
    WRATE        AS INTEGER
    ERATE        AS INTEGER
    GRate        AS STRING * 1
    CUSTTYPE     AS STRING * 1
    WAvgUse      AS LONG
    WAvgCt       AS INTEGER
    CurBal       AS SINGLE
    PastDue      AS SINGLE
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
    WMeterNum AS LONG
    EMETERNUM AS LONG
    Taxable   AS STRING * 1
    Security1 AS INTEGER
    Security2 AS INTEGER
    fadj      AS SINGLE
    eavguse   AS LONG
    eavgct    AS INTEGER
    demand    AS LONG
    add3      AS STRING * 25
  END TYPE

TYPE WBMeter
  meter AS STRING * 12
  Record AS SINGLE
END TYPE


    

TYPE CustRecType1    'WBCUST.DAT
    lname        AS STRING * 25
    firstname    AS STRING * 25
    ADD1         AS STRING * 30
    ADD2         AS STRING * 30
    SADD         AS STRING * 30
    city         AS STRING * 20
    st           AS STRING * 2
    zip          AS STRING * 9
    phone        AS STRING * 12
    book         AS INTEGER
    meter        AS STRING * 12  '
    Cycle        AS STRING * 1
    firsttr      AS DOUBLE       'cust trans pointers
    lasttr       AS DOUBLE       'cust trans pointers
    CUSTTYPE     AS STRING * 2
    CurBal       AS DOUBLE
    PastDue      AS DOUBLE
    AcctStatus   AS STRING * 1
    SSN          AS STRING * 12
    DrvLic       AS STRING * 12
    WMInstall    AS STRING * 6
    GMInstall    AS STRING * 6
    WAvgUse      AS SINGLE
    WAvgCt       AS INTEGER
    GAvgUse      AS SINGLE
    GAvgCnt      AS INTEGER
    WCurRead     AS DOUBLE
    WPastRead    AS DOUBLE
    WCurReadDate AS STRING * 6
    WPastReadDate AS STRING * 6
    WMeterNum    AS STRING * 10
    WMinimum     AS INTEGER
    WMeterEst   AS STRING * 1
    GMeterEst   AS STRING * 1
    GCurRead   AS DOUBLE
    GPastRead  AS DOUBLE
    GCurReadDate AS STRING * 6
    GPastReadDate AS STRING * 6
    GMeterNum    AS STRING * 10
    GMinimum     AS INTEGER
    waterdep     AS SINGLE       'Water deposit
    GasDep  AS SINGLE
    REV1    AS DOUBLE
    REV2    AS DOUBLE
    REV3    AS DOUBLE
    REV4    AS DOUBLE
    REV5    AS DOUBLE
    REV6    AS DOUBLE
    Rev7    AS DOUBLE
    Rev8    AS DOUBLE
    rate1 AS STRING * 2
    rate2 AS STRING * 2
    rate3 AS STRING * 2
    rate4 AS STRING * 2
    rate5 AS STRING * 2
    rate6 AS STRING * 2
    rate7 AS STRING * 2
    rate8 AS STRING * 2
    workorder AS STRING * 1
    WorkOrderNum AS SINGLE
    WorkOrdBeg AS STRING * 6
    WorkOrdEnd AS STRING * 6
    messfirst AS SINGLE
    messlast AS SINGLE
    comment AS STRING * 30
    seqnum AS SINGLE
   FirstWo AS SINGLE
    LastWo AS SINGLE
    MtrTime AS STRING * 6
    ETime AS DOUBLE
    ProRate AS STRING * 3
    AmtOwed AS SINGLE
    AmtPd AS SINGLE
    Payment AS SINGLE
    RS AS INTEGER
    MeterSeq AS SINGLE
    BillCopy AS SINGLE
    HWCredit AS STRING * 1
    ACCredit AS STRING * 1
    BankDraft AS STRING * 1
  END TYPE

      
TYPE DraftType
     bankname AS STRING * 30
     BankLoc AS STRING * 30
     BankTransit AS STRING * 9
     BankAcct AS STRING * 17
     MasterAcct AS SINGLE
     PreNote AS STRING * 1
     DraftType AS STRING * 1
END TYPE

TYPE CustRecType2
    ramt1    AS SINGLE
    ramt2    AS SINGLE
    ramt3    AS SINGLE
    RDESC1    AS STRING * 25
    RDESC2    AS STRING * 25
    RDESC3    AS STRING * 25
    racct1   AS STRING * 9
    racct2   AS STRING * 9
    racct3   AS STRING * 9
    oamt1 AS SINGLE
    oamt2 AS SINGLE
    ODESC1 AS STRING * 25
    ODESC2 AS STRING * 25
    oacct1 AS STRING * 9
    oacct2 AS STRING * 9
    notused AS STRING * 66
END TYPE


TYPE WBCUSTHHRECTYPE
   line1 AS STRING * 20
   line2 AS STRING * 20
   line3 AS STRING * 20
END TYPE

