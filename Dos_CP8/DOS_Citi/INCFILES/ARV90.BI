TYPE ARCustRecType85
    CUSTNUMB AS STRING * 10
    SORTNAME AS STRING * 10
    BILLNAME AS STRING * 35
    ADDRESS1 AS STRING * 35
    ADDRESS2 AS STRING * 35
    CITY     AS STRING * 20
    STATE    AS STRING * 2
    ZIPCODE  AS STRING * 10
    CUSTNAME AS STRING * 35
    CONTACT  AS STRING * 30

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

TYPE oARCustRecTypev90
    CUSTNUMB     AS STRING * 10
    SORTNAME     AS STRING * 10
    BILLNAME     AS STRING * 35
    ADDRESS1     AS STRING * 35
    ADDRESS2     AS STRING * 35
    CITY         AS STRING * 20
    STATE        AS STRING * 2
    ZIPCODE      AS STRING * 10
    CUSTNAME     AS STRING * 35
    CONTACT      AS STRING * 30

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
    RoomtoGrow   AS STRING * 163
END TYPE

TYPE ARCustRecType
    CUSTNUMB     AS STRING * 10
    SORTNAME     AS STRING * 10
    BILLNAME     AS STRING * 35
    ADDRESS1     AS STRING * 35
    ADDRESS2     AS STRING * 35
    CITY         AS STRING * 20
    STATE        AS STRING * 2
    ZIPCODE      AS STRING * 10
    CUSTNAME     AS STRING * 35
    CONTACT      AS STRING * 30
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

    AcctBal      AS DOUBLE
    IssueLicense AS STRING * 1    'rem y/n
    Deleted      AS STRING * 1    'rem y=deleted :AnyThing Else is Non-Deleted
    FirstTrans   AS LONG
    LastTrans    AS LONG
    RoomtoGrow   AS STRING * 164
END TYPE


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

TYPE ARCatCodeRecTypeWB
    OLDCATCODE    AS STRING * 3    'Not Used in Version 8.5 work2 directory
    CODEDESC      AS STRING * 35
    APPNUMB       AS INTEGER       'Not Used
    BILLCODE      AS INTEGER       'Not Used
    REVGLNUM      AS STRING * 14
    CASHACCT      AS STRING * 14
    CATCODE       AS STRING * 5    'Alabama Code Specific
    ARGLACCT      AS STRING * 14
    Fee           AS SINGLE
    Extra         AS STRING * 35
END TYPE

TYPE AROldCatCodeRecType
    CATCODE    AS STRING * 3    'Not Used in Version 8.5 work2 directory
    CODEDESC   AS STRING * 35
    APPNUMB    AS INTEGER       'Not Used
    BILLCODE   AS INTEGER       'Not Used
    REVGLNUM   AS STRING * 14
    CASHACCT   AS STRING * 14
    ALCATCODE  AS STRING * 5    'Alabama Code Specific
    ARGLACCT   AS STRING * 14
    Extra      AS STRING * 39
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
    CUSTNAME                    AS STRING * 35
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


TYPE AREditPaymentRecType
    CustNumber          AS STRING * 10
    CUSTNAME            AS STRING * 35
    TranDate            AS INTEGER
    Amount              AS DOUBLE
    ISSUELIC            AS STRING * 1      'Y/N Answer
    Desc                AS STRING * 35
    SetFee              AS STRING * 1      'Y/N Answer  if Y the set Fee in Customer Record=Payment Amount Entered
    ISSUEFEE            AS SINGLE          'Amount of Issuance Fee Pd
    CatCodeRecord       AS INTEGER
    Extra               AS STRING * 1
END TYPE

TYPE WBeechCode
   Code AS STRING * 5
   Desc AS STRING * 40
   Filler AS STRING * 10
   Amount AS STRING * 6
   ACTIVE AS STRING * 1
   Filler1 AS STRING * 2
END TYPE

TYPE WBeechRecType
    ActNum AS STRING * 5
    BName  AS STRING * 30
    Add1   AS STRING * 30
    Add2   AS STRING * 30
    Add3   AS STRING * 30
    OName  AS STRING * 25
    SName  AS STRING * 10
    Fil1   AS STRING * 12
    Phone  AS STRING * 10
    Fil2   AS STRING * 6
    EDate  AS STRING * 6
    Fil3   AS STRING * 4
    O1     AS STRING * 5
    O2     AS STRING * 5
    O3     AS STRING * 5
    O4     AS STRING * 5
    O5     AS STRING * 5
    q1     AS STRING * 3
    q2     AS STRING * 3
    q3     AS STRING * 3
    q4     AS STRING * 3
    q5     AS STRING * 3
    a1     AS STRING * 7
    a2     AS STRING * 7
    a3     AS STRING * 7
    a4     AS STRING * 7
    a5     AS STRING * 7
    Fil4   AS STRING * 91
    Status AS STRING * 1
    Pad AS STRING * 2
END TYPE

CONST User$ = " v9.0"



