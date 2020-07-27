
TYPE ARCustRecType
    CUSTNUMB     AS STRING * 10
    SORTNAME     AS STRING * 10
    BILLNAME     AS STRING * 35
    ADDRESS1     AS STRING * 35
    ADDRESS2     AS STRING * 35
    CITY         AS STRING * 20
    STATE        AS STRING * 2
    ZIPCODE      AS STRING * 10
    CustName     AS STRING * 35
    BILLCAT      AS STRING * 5
    SOSEC        AS STRING * 11
    DRVLIC       AS STRING * 12
    DATEOPED     AS INTEGER
    BILLCMT      AS STRING * 20
    PAYCMT       AS STRING * 20
    CASHONLY     AS STRING * 1
    APPNUMB      AS INTEGER
    BILLFORM     AS INTEGER
    HPHONE       AS STRING * 14
    WPHONE       AS STRING * 14
    FeeAmt       AS DOUBLE
    LICENSE      AS STRING * 12
    VALID        AS INTEGER
    AcctBal      AS DOUBLE
    OldFirstTrans   AS INTEGER
    OldLastTrans    AS INTEGER
    Deleted      AS STRING * 1      'rem y=deleted :AnyThing Else is Non-Deleted
    FirstTrans   AS LONG
    LastTrans    AS LONG
    IssueLicense AS STRING * 1      ' rem y/n
    IssuanceFee  AS CURRENCY        ' Alabama Specific
    CustLocation AS STRING * 1      ' Alabama Specific
    RoomtoGrow   AS STRING * 164
END TYPE

TYPE ARCatCodeRecType
    CATCODE    AS STRING * 3
    CODEDESC   AS STRING * 35
    APPNUMB    AS INTEGER       'Not Used
    BILLCODE   AS INTEGER       'Not Used
    REVGLNUM   AS STRING * 14
    CASHACCT   AS STRING * 14
    ALCATCODE  AS STRING * 5    'Alabama Code Specific
    ARGLACCT   AS STRING * 14
    Extra      AS STRING * 39
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
    CatCodeRec          AS INTEGER               'Place to Grab G/L Acct #'s
    ExtraRoom           AS STRING * 40
    NextTrans           AS LONG
END TYPE

TYPE AREditBegBalRecType
    CustNumber                  AS STRING * 10
    CustName                    AS STRING * 35
    TranDate                    AS INTEGER
    Amount                      AS DOUBLE
    IssueLic                    AS STRING * 1      'Y/N Answer
    SetFee                      AS STRING * 1        'Y/N Answer
    Extra                       AS STRING * 42
END TYPE

TYPE AREditPaymentRecType
    CustNumber          AS STRING * 10
    CustName            AS STRING * 35
    TranDate            AS INTEGER
    Amount              AS DOUBLE
    IssueLic            AS STRING * 1      ' Y/N Answer
    Desc                AS STRING * 35
    SetFee              AS STRING * 1        'Y/N Answer  if Y the set Fee in Customer Record=Payment Amount Entered
    IssueFee            AS SINGLE          'Amount of Issuance Fee Pd
    CatCodeRecord       AS INTEGER
    Extra               AS STRING * 1
END TYPE

