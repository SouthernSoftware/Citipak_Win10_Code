TYPE oUBCustRecType
    STATUS   AS STRING * 1
    OPENDATE AS INTEGER
    SEARCH   AS STRING * 10
    CUSTNAME AS STRING * 35
    ADDR1    AS STRING * 35
    ADDR2    AS STRING * 35
    CITY     AS STRING * 18
    STATE    AS STRING * 2
    ZIPCODE  AS STRING * 10
    CUSTTYPE AS STRING * 3
    PostRte  AS STRING * 4
    BILLCOPY AS INTEGER
    HPHONE   AS STRING * 14
    WPHONE   AS STRING * 14
    SOSEC    AS STRING * 11
    DRVLIC   AS STRING * 16
    TAXEXPT  AS STRING * 1
    SRCIT    AS STRING * 1
    EPP      AS STRING * 1
    EPPAMT   AS DOUBLE
'page 2
    CASHONLY     AS STRING * 1
    LATEFEE      AS STRING * 1
    CUTOFFYN     AS STRING * 1

    AMTOWED      AS DOUBLE
    AMTPD        AS DOUBLE
    PAYMENT      AS DOUBLE
    REVSRCE      AS INTEGER

    DRAFTYPE     AS STRING * 1
    BANKNAME     AS STRING * 34
    BANKLOC      AS STRING * 30
    TRANSIT      AS STRING * 9
    BANKACCT     AS STRING * 20
    BILLCMNT     AS STRING * 25
    PAYCMNT      AS STRING * 25
'END OF QS FORM

    CustPIN         AS LONG

    OwnFrstLocRec   AS LONG
    OwnLastLocRec   AS LONG

    CustFrstLocRec  AS LONG
    CustLastLocRec  AS LONG

    Pend2Rec        AS LONG

    LastTrans       AS LONG

    CurrBalance     AS DOUBLE
    PrevBalance     AS DOUBLE

    CurrRevAmts(1 TO 15) AS DOUBLE

    DepositAmt     AS DOUBLE
    DelFlag        AS INTEGER
    PreNoteFlag    AS INTEGER

    PrevLocRec     AS LONG
    FillPad        AS STRING * 24
END TYPE


TYPE oUBCustRecType3
    STATUS   AS STRING * 1
    OPENDATE AS INTEGER
    SEARCH   AS STRING * 10
    CUSTNAME AS STRING * 35
    ADDR1    AS STRING * 35
    ADDR2    AS STRING * 35
    CITY     AS STRING * 18
    STATE    AS STRING * 2
    ZIPCODE  AS STRING * 10
    CUSTTYPE AS STRING * 3
    PostRte  AS STRING * 4
    BILLCOPY AS INTEGER
    HPHONE   AS STRING * 14
    WPHONE   AS STRING * 14
    SOSEC    AS STRING * 11
    DRVLIC   AS STRING * 16
    TAXEXPT  AS STRING * 1
    SRCIT    AS STRING * 1
    EPP      AS STRING * 1
    EPPAMT   AS DOUBLE
'page 2
    CASHONLY     AS STRING * 1
    LATEFEE      AS STRING * 1
    CUTOFFYN     AS STRING * 1
    AMTOWED      AS DOUBLE
    AMTPD        AS DOUBLE
    PAYMENT      AS DOUBLE
    REVSRCE      AS INTEGER
    DRAFTYPE     AS STRING * 1
    BANKNAME     AS STRING * 34
    BANKLOC      AS STRING * 30
    TRANSIT      AS STRING * 9
    BANKACCT     AS STRING * 20
    BILLCMNT     AS STRING * 25
    PAYCMNT      AS STRING * 25
'END OF QS FORM

    CustPIN      AS LONG

    OwnFrstLocRec     AS LONG
    OwnLastLocRec     AS LONG

    CustFrstLocRec  AS LONG
    CustLastLocRec  AS LONG

    Pend2Rec       AS LONG

    LastTrans      AS LONG

    CurrBalance    AS DOUBLE
    PrevBalance    AS DOUBLE

    CurrRevAmts(1 TO 15) AS DOUBLE

    DepositAmt     AS DOUBLE
    DelFlag        AS INTEGER
    PreNoteFlag    AS INTEGER

    'FillPad        AS STRING * 60
END TYPE


TYPE oUBCustRecType1
    BOOK     AS STRING * 2
    SEQNUMB  AS STRING * 6
    ASTATUS  AS STRING * 10
    SEARCHN  AS STRING * 10
    FULLNAME AS STRING * 35
    CADDR1   AS STRING * 35
    CADDR2   AS STRING * 35
    CITY     AS STRING * 18
    STATE    AS STRING * 2
    ZIPCODE  AS STRING * 10
    SOSEC    AS STRING * 11
    DRVLIC   AS STRING * 16
    OPENDATE AS INTEGER
    BILLCOM  AS STRING * 30
    PAYCOM   AS STRING * 30
    CASHONLY AS STRING * 1
    HPHONE   AS STRING * 14
    WPHONE   AS STRING * 14
    LATEFEE  AS STRING * 1
    CUTOFFYN AS STRING * 1
'Page 2
    BillCycle AS INTEGER
    CUSTTYPE AS STRING * 3
    ZONE     AS STRING * 3
    PostRte  AS STRING * 4
    METERSEQ AS LONG
    BILLCOPY AS INTEGER
    PRORATE  AS INTEGER
    TAXEXPT  AS STRING * 1
    SRCIT    AS STRING * 1
    BARCODE  AS STRING * 1
    EPP      AS STRING * 1
    EPPAMT   AS SINGLE
    BILLWHO  AS STRING * 1
    OWNER    AS STRING * 23
    OWNADD1  AS STRING * 23
    OWNADD2  AS STRING * 23
    OWNADD3  AS STRING * 23
    OWNSOSEC AS STRING * 11
    DRAFTYPE AS STRING * 1
    BANKNAME AS STRING * 15
    BANKLOC  AS STRING * 15
    TRANSIT  AS STRING * 9
    BANKACCT AS STRING * 15
'Page 3
    FlatRates(1 TO 4) AS FlatRateDataType
    RevMeters(1 TO 10) AS RevDataType
'Page 4
    SERADDR  AS STRING * 35

    MeterData(1 TO 7) AS MeterDescType

    AMTOWED  AS SINGLE
    AMTPD    AS SINGLE
    PAYMENT  AS SINGLE
    REVSRCE  AS INTEGER
    REFMFEE  AS DOUBLE
    MFEENORF AS DOUBLE
'end of actual quick screen form
    LockedFlag AS INTEGER
    PendingRec AS INTEGER

    CurBalance AS DOUBLE
    PastDue    AS DOUBLE
    FirstTrans AS LONG
    LastTrans  AS LONG
END TYPE

TYPE oLocMeterType3
    MTRNum    AS STRING * 12
    MTRMulti  AS INTEGER
    MTRType   AS STRING * 1
    MTRUnit   AS STRING * 1
    NumUser   AS INTEGER
    InsDate   AS INTEGER
    CurRead   AS LONG
    PrevRead  AS LONG
    CurDate   AS INTEGER
END TYPE

TYPE oUBLocationRecType
    BOOK     AS STRING * 2
    SEQNUMB  AS STRING * 6
    SERVADDR AS STRING * 35
    BILLCYCL AS INTEGER
    ZONE     AS STRING * 3
    SEQ      AS LONG
    PUMPCODE AS STRING * 4
    USER1    AS STRING * 4
    USER2    AS STRING * 4
    REFMFEE  AS DOUBLE
    MFEENORF AS DOUBLE

'Page 2
    Serv(1 TO 15) AS ServicesType

'Page 3
'NOTE: Meters 1 thru 5 are on Page 3
'             6 and 7  are on Page 4
    LocMeters(1 TO 7) AS LocMeterType

'Page 4
    FlatRates(1 TO 4) AS FlatRateType
'
    LocatStatus       AS STRING * 1
    OwnerRec          AS LONG
    PrevLocRec        AS LONG
    NextLocRec        AS LONG
    CustRec           AS LONG
    PendCustRec       AS LONG
    OwnerPIN          AS LONG
    LocaPIN           AS LONG
    NotesRec          AS LONG
    Pad               AS STRING * 14
END TYPE

TYPE oLocMeterType
    MTRNum   AS STRING * 12
    CurRead  AS LONG
    CurDate  AS INTEGER
    MTRMulti AS INTEGER
    MTRType  AS STRING * 1
    MTRUnit  AS STRING * 1
    NoUser   AS INTEGER
    InsDate  AS INTEGER
END TYPE

TYPE oLocMeterType2
    MTRNum   AS STRING * 12
    CurRead  AS LONG
    CurDate  AS INTEGER
    MTRMulti AS INTEGER
    MTRType  AS STRING * 1
    MTRUnit  AS STRING * 1
    NumUser   AS INTEGER
    InsDate  AS INTEGER
    PrevRead AS LONG
END TYPE

TYPE oUBTransRecType
   TransDate              AS INTEGER      '
   TransType              AS INTEGER      '
   TransDesc              AS STRING * 21  'may change
   TransAmt               AS DOUBLE       'total revenue amount
   RevAmt(1 TO 15)        AS DOUBLE       'Revenue amounts
   TaxAmt(1 TO 15)        AS SINGLE       'Tax Amounts
   CurRead(1 TO 7)        AS LONG         'Last/Current meter readings
   PrevRead(1 TO 7)       AS LONG         'Previous readings
   EstRead(1 TO 7)        AS STRING * 1   'Y/N Flags for meter est's
   BillNumber             AS LONG         'Number on the bill that Printed
   ReadDate               AS INTEGER
   BillDate               AS INTEGER
   PastDueDate            AS INTEGER
   DraftDate              AS INTEGER
   CustLocation           AS LONG         'Pointer to Location RecNo
   CustStatus             AS STRING * 1   'Customer Status at Time of Transaction
   CustCode               AS STRING * 2   'Customer Code at Time of Transaction
   PostRoute              AS STRING * 4   'Postal Route at Time of Transaction
   BillCycle              AS STRING * 2   'Customer Billing Cycle at Time of Transaction
   PayTypeCode            AS INTEGER      'Payment Type:  1=Cash, 2=Check
   OperatorNumber         AS INTEGER      '
   CustAcctNo             AS LONG         'Pointer to RecNo in ubcust.dat
   PrevTrans              AS LONG
   NextTrans              AS LONG
   ActiveFlag             AS INTEGER      'Valid transaction flag
   RunBalance             AS DOUBLE
   CheckAmount            AS DOUBLE
   CashAmount             AS DOUBLE
   BillMsg                AS STRING * 20
   ApplyDepFlag           AS STRING * 1
   Posted2GL              AS STRING * 1
   PrevDate               AS INTEGER
END TYPE

TYPE UBCustRecTypeV1
    STATUS   AS STRING * 1
    OPENDATE AS INTEGER
    SEARCH   AS STRING * 10
    CUSTNAME AS STRING * 35
    ADDR1    AS STRING * 35
    ADDR2    AS STRING * 35
    CITY     AS STRING * 18
    STATE    AS STRING * 2
    ZIPCODE  AS STRING * 10
    CUSTTYPE AS STRING * 3
    PostRte  AS STRING * 4
    BILLCOPY AS INTEGER
    HPHONE   AS STRING * 14
    WPHONE   AS STRING * 14
    SOSEC    AS STRING * 11
    DRVLIC   AS STRING * 16
    TAXEXPT  AS STRING * 1
    SRCIT    AS STRING * 1
    EPP      AS STRING * 1
    EPPAMT   AS DOUBLE
'page 2
    CASHONLY     AS STRING * 1
    LATEFEE      AS STRING * 1
    CUTOFFYN     AS STRING * 1

    Monthly(1 TO 2) AS MonthlyPayType

    DRAFTYPE     AS STRING * 1
    BANKNAME     AS STRING * 34
    BANKLOC      AS STRING * 30
    TRANSIT      AS STRING * 9
    BANKACCT     AS STRING * 20
    BILLCMNT     AS STRING * 25
    PAYCMNT      AS STRING * 25
'END OF QS FORM
    CustPIN           AS LONG
    OwnFrstLocRec     AS LONG
    OwnLastLocRec     AS LONG
    CustFrstLocRec    AS LONG
    CustLastLocRec    AS LONG
    Pend2Rec          AS LONG
    LastTrans         AS LONG
    CurrBalance       AS DOUBLE
    PrevBalance       AS DOUBLE
    CurrRevAmts(1 TO 15) AS DOUBLE
    DepositAmt        AS DOUBLE
    DelFlag           AS INTEGER
    PreNoteFlag       AS INTEGER
    PrevLocRec        AS LONG
    WOLastTrans       AS LONG
    'work order last trans pointer
    EstFlag           AS STRING * 1
    MessageRec        AS LONG
    FillPad           AS STRING * 15
END TYPE

TYPE UBLocationRecTypeV1
    BOOK           AS STRING * 2
    SEQNUMB        AS STRING * 6
    LocatStatus    AS STRING * 1    ' 1
    SERVADDR       AS STRING * 35
    BILLCYCL       AS INTEGER
    ZONE           AS STRING * 3
    SEQ            AS LONG
    BILLWHO        AS STRING * 1
    PUMPCODE       AS STRING * 4
    USERCODE1      AS STRING * 4
    USERCODE2      AS STRING * 4
    MFEE1          AS DOUBLE
    MFEE2          AS DOUBLE

    Serv(1 TO 15)     AS ServicesType
    LocMeters(1 TO 7) AS LocMeterType
    FlatRates(1 TO 4) AS FlatRateType

    OwnerRec          AS LONG          ' 5
    PrevLocRec        AS LONG          ' 9
    NextLocRec        AS LONG          ' 13
    CustRec           AS LONG          ' 17
    PendCustRec       AS LONG          ' 21
    OwnerPIN          AS LONG          ' 25
    LocaPIN           AS LONG          ' 29
    NotesRec          AS LONG          ' 33
    DelFlag           AS INTEGER
    EstFlag           AS STRING * 1
    'add message
    MessageRec        AS INTEGER
    Pad               AS STRING * 1   ' 48

END TYPE

