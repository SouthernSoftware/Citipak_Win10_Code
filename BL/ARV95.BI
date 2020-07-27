TYPE ARCustRecType
    CUSTNUMB AS STRING * 10
    SORTNAME AS STRING * 10
    BILLNAME AS STRING * 35
    ADDRESS1 AS STRING * 35
    ADDRESS2 AS STRING * 35
    CITY     AS STRING * 20
    STATE    AS STRING * 2
    ZIPCODE  AS STRING * 10
    CustName AS STRING * 35
    Contact  AS STRING * 30

    BILLCAT1     AS STRING * 5
    DESC1        AS STRING * 15
    REV1         AS LONG
    Fee1         AS DOUBLE
    BILLCAT2     AS STRING * 5
    DESC2        AS STRING * 15
    REV2         AS LONG
    Fee2         AS DOUBLE
    BILLCAT3     AS STRING * 5
    DESC3        AS STRING * 15
    REV3         AS LONG
    Fee3         AS DOUBLE
    BILLCAT4     AS STRING * 5
    DESC4        AS STRING * 15
    REV4         AS LONG
    Fee4         AS DOUBLE
    BILLCAT5     AS STRING * 5
    DESC5        AS STRING * 15
    REV5         AS LONG
    Fee5         AS DOUBLE

'************
    IssuanceFee  AS DOUBLE
    CustLocation AS STRING * 1
    WPHONE       AS STRING * 14
    FeeAmt       AS DOUBLE
    LICENSE      AS STRING * 12
    VALID        AS INTEGER
    Inactive     AS STRING * 1    '"Y" if account is inactive
    ProRate      AS INTEGER       'prorate percentage
    AcctBal      AS DOUBLE
    IssueLicense AS STRING * 1    'y/n
    Deleted      AS STRING * 1    '(yY)=deleted, anything else isn't
    FirstTrans   AS LONG
    LastTrans    AS LONG

    LicBal       AS DOUBLE
    FeeBal       AS DOUBLE
    PenBal       AS DOUBLE
    RoomtoGrow   AS STRING * 136
    ChkByte      AS STRING * 1
'************
END TYPE

TYPE ARCustRecTypeV85
    CUSTNUMB     AS STRING * 10
    SORTNAME     AS STRING * 10
    BILLNAME     AS STRING * 35
    ADDRESS1     AS STRING * 35
    ADDRESS2     AS STRING * 35
    CITY         AS STRING * 20
    STATE        AS STRING * 2
    ZIPCODE      AS STRING * 10
    CustName     AS STRING * 35
    Contact      AS STRING * 30

    BILLCAT1     AS STRING * 5
    DESC1        AS STRING * 15
    REV1         AS LONG
    Fee1         AS DOUBLE

    BILLCAT2     AS STRING * 5
    DESC2        AS STRING * 15
    REV2         AS LONG
    Fee2         AS DOUBLE

    BILLCAT3     AS STRING * 5
    DESC3        AS STRING * 15
    REV3         AS LONG
    Fee3         AS DOUBLE

    BILLCAT4     AS STRING * 5
    DESC4        AS STRING * 15
    REV4         AS LONG
    Fee4         AS DOUBLE

    BILLCAT5     AS STRING * 5
    DESC5        AS STRING * 15
    REV5         AS LONG
    Fee5         AS DOUBLE

    IssuanceFee  AS CURRENCY
    CustLocation AS STRING * 1
    WPHONE       AS STRING * 14
    FeeAmt       AS DOUBLE
    LICENSE      AS STRING * 12
    VALID        AS INTEGER

'022101
    Inactive     AS STRING * 1    '"Y" if account is inactive

    AcctBal      AS DOUBLE
    IssueLicense AS STRING * 1    'y/n
    Deleted      AS STRING * 1    'y=deleted AnyThing Else is Not Deleted
    FirstTrans   AS LONG
    LastTrans    AS LONG

    LicBal       AS DOUBLE
    FeeBal       AS DOUBLE
    PenBal       AS DOUBLE
    ProRate      AS INTEGER
    RoomtoGrow   AS STRING * 137

END TYPE

'TYPE ARCustRecType
'    CUSTNUMB     AS STRING * 10
'    SORTNAME     AS STRING * 10
'    BILLNAME     AS STRING * 35
'    ADDRESS1     AS STRING * 35
'    ADDRESS2     AS STRING * 35
'    CITY         AS STRING * 20
'    STATE        AS STRING * 2
'    ZIPCODE      AS STRING * 10
'    CUSTNAME     AS STRING * 35
'    CONTACT      AS STRING * 30
'    BILLCAT1     AS STRING * 5
'    DESC1        AS STRING * 15
'    REV1         AS LONG
'    Fee1         AS DOUBLE
'    BILLCAT2     AS STRING * 5
'    DESC2        AS STRING * 15
'    REV2         AS LONG
'    Fee2         AS DOUBLE
'
'    BILLCAT3     AS STRING * 5
'    DESC3        AS STRING * 15
'    REV3         AS LONG
'    Fee3         AS DOUBLE
'
'    BILLCAT4     AS STRING * 5
'    DESC4        AS STRING * 15
'    REV4         AS LONG
'    Fee4         AS DOUBLE
'
'    BILLCAT5     AS STRING * 5
'    DESC5        AS STRING * 15
'    REV5         AS LONG
'    Fee5         AS DOUBLE
'
'    IssuanceFee  AS CURRENCY
'    CustLocation AS STRING * 1
'    WPHONE       AS STRING * 14
'    FeeAmt       AS DOUBLE
'    LICENSE      AS STRING * 12
'    VALID        AS INTEGER
'
'    AcctBal      AS DOUBLE
'    IssueLicense AS STRING * 1    'rem y/n
'    Deleted      AS STRING * 1    'rem y=deleted :AnyThing Else is Non-Deleted
'    FirstTrans   AS LONG
'    LastTrans    AS LONG
'    RoomtoGrow   AS STRING * 164
'END TYPE


TYPE ARCatCodeRecType
    CATCODE    AS STRING * 5    'Not Used in Version 8.5 work2 directory
    CODEDESC   AS STRING * 35
    Fee     AS SINGLE
    REVGLNUM   AS STRING * 14
    CASHACCT   AS STRING * 14
    ARGLACCT   AS STRING * 14
    CodeType   AS STRING * 1    ' F=Flat M=Multiplier S=Step
    Percent    AS SINGLE
    Maximum    AS DOUBLE
    Extra      AS STRING * 157
END TYPE

TYPE ARNewCatCodeRecType
    CATCODE    AS STRING * 5    'Not Used in Version 8.5 work2 directory
    CodeType   AS STRING * 1    ' F=Flat M=Multiplier S=Step
    CODEDESC   AS STRING * 35
    Fee        AS SINGLE
    BaseAmt1   AS SINGLE
    Recpt1     AS DOUBLE
    Percent1   AS SINGLE
    Maximum1   AS DOUBLE
    BaseAmt2   AS SINGLE
    Recpt2     AS DOUBLE
    Percent2   AS SINGLE
    Maximum2   AS DOUBLE
    BaseAmt3   AS SINGLE
    Recpt3     AS DOUBLE
    Percent3   AS SINGLE
    Maximum3   AS DOUBLE
    BaseAmt4   AS SINGLE
    Recpt4     AS DOUBLE
    Percent4   AS SINGLE
    Maximum4   AS DOUBLE
    BaseAmt5   AS SINGLE
    Recpt5     AS DOUBLE
    Percent5   AS SINGLE
    Maximum5   AS DOUBLE
    REVGLNUM   AS STRING * 14
    CASHACCT   AS STRING * 14
    ARGLACCT   AS STRING * 14

    BaseAmt6   AS SINGLE
    Recpt6     AS DOUBLE
    Percent6   AS SINGLE
    Maximum6   AS DOUBLE
    RateStep   AS LONG
    Extra      AS STRING * 36
END TYPE

TYPE ARCustIDXRecType
    IDXNAME     AS STRING * 10
    IDXRECORD   AS INTEGER
    ExtraRoom   AS STRING * 52
END TYPE

'*****************
'TYPE ARTransRecType
'    CustomerNumber      AS STRING * 10
'    TransDate           AS INTEGER
'    TransAmount         AS DOUBLE
'    TransType           AS INTEGER
'    TransDesc           AS STRING * 35
'    CashAmount          AS DOUBLE
'    ChkAmount           AS DOUBLE
'    BalanceAfterTrans   AS DOUBLE
'    Posted2GL           AS STRING * 1
'    CatCodeRec          AS INTEGER               'Place to Grab G/L Acct #'s
'    ExtraRoom           AS STRING * 40
'    NextTrans           AS LONG
'END TYPE


'*****************
TYPE ARTransRecType
    CustomerNumber      AS STRING * 10
    TransDate           AS INTEGER
    TransAmount         AS DOUBLE
    TransType           AS INTEGER
    TransDesc           AS STRING * 35
    CashAmount          AS DOUBLE
    ChkAmount           AS DOUBLE
    BalanceAfterTrans   AS DOUBLE
    Posted2GL           AS STRING * 1
    CatCodeRec1         AS INTEGER           'Place to Grab G/L Acct #'s
    CatCodeRec2         AS INTEGER           'Place to Grab G/L Acct #'s
    CatCodeRec3         AS INTEGER           'Place to Grab G/L Acct #'s
    CatCodeRec4         AS INTEGER           'Place to Grab G/L Acct #'s
    CatCodeRec5         AS INTEGER           'Place to Grab G/L Acct #'s
    LicAMT              AS DOUBLE
    FeeAmt              AS DOUBLE
    PenAMT              AS DOUBLE
    ExtraRoom           AS STRING * 8
    NextTrans           AS LONG
END TYPE

TYPE AREditBegBalRecType
    CustNumber                  AS STRING * 10
    CustName                    AS STRING * 35
    TranDate                    AS INTEGER
    Amount                      AS DOUBLE
    ISSUELIC                    AS STRING * 1       'Y/N Answer
    SetFee                      AS STRING * 1       'Y/N Answer
    LicAMT                      AS DOUBLE
    PenAMT                      AS DOUBLE
    TRType                      AS INTEGER
    TDesc                       AS STRING * 20
    Extra                       AS STRING * 4
END TYPE

TYPE oAREditBegBalRecType
    CustNumber                  AS STRING * 10
    CustName                    AS STRING * 35
    TranDate                    AS INTEGER
    Amount                      AS DOUBLE
    ISSUELIC                    AS STRING * 1       'Y/N Answer
    SetFee                      AS STRING * 1       'Y/N Answer
    Extra                       AS STRING * 42
END TYPE

TYPE AREditPaymentRecType
    TRANTYPE        AS STRING * 15
    TranDate        AS INTEGER
    CustNumber      AS STRING * 10
    CustName        AS STRING * 35
    ADD1            AS STRING * 35
    CITY            AS STRING * 25
    STATE           AS STRING * 2
    ZIPCODE         AS STRING * 10
    Amount          AS DOUBLE
    CASHCHK         AS STRING * 9
    CASHAMT         AS DOUBLE
    CHKAMT          AS DOUBLE
    AMTPAID         AS DOUBLE
    CHANGE          AS DOUBLE
    ISSUELIC        AS STRING * 1
    SetFee          AS STRING * 1
    ISSueFEE        AS DOUBLE
    DESC            AS STRING * 20
    LICDUE          AS DOUBLE
    FEEDUE          AS DOUBLE
    PENDUE          AS DOUBLE
    LICPAID         AS DOUBLE
    FEEPAID         AS DOUBLE
    PENPAID         AS DOUBLE
    TOTDUE          AS DOUBLE
    TOTPAID         AS DOUBLE
END TYPE



'TYPE AREditPaymentRecType
'    CustNumber          AS STRING * 10
'    CUSTNAME            AS STRING * 35
'    TranDate            AS INTEGER
'    Amount              AS DOUBLE
'    ISSUELIC            AS STRING * 1      'Y/N Answer
'    Desc                AS STRING * 35
'    SetFee              AS STRING * 1      'Y/N Answer  if Y the set Fee in Customer Record=Payment Amount Entered
'    ISSUEFEE            AS SINGLE          'Amount of Issuance Fee Pd
'    CatCodeRecord       AS INTEGER
'    Extra               AS STRING * 1
'END TYPE

CONST User$ = " v9.5"



TYPE ARCustRecTypeMaggieVally       'used in verion 8.5 work2dir
    CUSTNUMB     AS STRING * 10
    SORTNAME     AS STRING * 10
    BILLNAME     AS STRING * 35
    ADDRESS1     AS STRING * 35
    ADDRESS2     AS STRING * 35
    CITY         AS STRING * 20
    STATE        AS STRING * 2
    ZIPCODE      AS STRING * 10
    CustName     AS STRING * 35
    Contact      AS STRING * 30

    BILLCAT1     AS STRING * 5
    DESC1        AS STRING * 15
    REV1         AS LONG
    Fee1         AS DOUBLE
    BILLCAT2     AS STRING * 5
    DESC2        AS STRING * 15
    REV2         AS LONG
    Fee2         AS DOUBLE
    BILLCAT3     AS STRING * 5
    DESC3        AS STRING * 15
    REV3         AS LONG
    Fee3         AS DOUBLE
    BILLCAT4     AS STRING * 5
    DESC4        AS STRING * 15
    REV4         AS LONG
    Fee4         AS DOUBLE
    BILLCAT5     AS STRING * 5
    DESC5        AS STRING * 15
    REV5         AS LONG
    Fee5         AS DOUBLE

    IssuanceFee  AS CURRENCY
    CustLocation AS STRING * 1
    WPHONE       AS STRING * 14
    FeeAmt       AS DOUBLE
    LICENSE      AS STRING * 12
    VALID        AS INTEGER
    'dodo AS STRING * 7
    AcctBal      AS DOUBLE
    IssueLicense AS STRING * 1    'rem y/n
    Deleted      AS STRING * 1    'rem y=deleted :AnyThing Else is Non-Deleted
    FirstTrans   AS LONG
    LastTrans    AS LONG
    RoomtoGrow   AS STRING * 164
END TYPE

