
TYPE DCCustRecType
    CUSTNUMB     AS STRING * 10
    SORTNAME     AS STRING * 10
    BILLNAME     AS STRING * 35
    ADDRESS1     AS STRING * 35
    ADDRESS2     AS STRING * 35
    city         AS STRING * 20
    STATE        AS STRING * 2
    ZIPCODE      AS STRING * 10
    SOSEC        AS STRING * 11
    DRVLIC       AS STRING * 12
    DATEOPED     AS INTEGER
    CASHONLY     AS STRING * 1
    resident     AS STRING * 1
    Owner        AS STRING * 1
    HPHONE       AS STRING * 14
    WPHONE       AS STRING * 14
    LICENSE      AS STRING * 12
    Valid        AS INTEGER
    AcctBal      AS DOUBLE
    Deleted      AS STRING * 1      'rem y=deleted :AnyThing Else is Non-Deleted
    FirstTrans   AS LONG
    LastTrans    AS LONG
    FirstCar     AS LONG
    LastCar      AS LONG
    SocSec1      AS STRING * 11
    OtherName    AS STRING * 25
    RoomtoGrow   AS STRING * 224
END TYPE

TYPE DCCatCodeRecType
    CATCODE    AS STRING * 3
    CODEDESC   AS STRING * 35
    APPNUMB    AS INTEGER
    BILLCODE   AS INTEGER
    REVGLNUM   AS STRING * 14
    CASHACCT   AS STRING * 14
    Fee        AS SINGLE
    Extra      AS STRING * 54
END TYPE

TYPE DCCustIDXRecType
    IDXNAME     AS STRING * 10
    IDXRECORD   AS INTEGER
    ExtraRoom   AS STRING * 52
END TYPE

TYPE DCTransRecType
    CustomerNumber AS STRING * 10
    TransDate AS INTEGER
    TransAmount AS DOUBLE
    TransType AS INTEGER
    TransDesc AS STRING * 35
    CashAmount AS DOUBLE
    ChkAmount AS DOUBLE
    BalanceAfterTrans AS DOUBLE
    makemodel AS STRING * 25
    StateTag AS STRING * 25
    ExpireDate AS INTEGER
    Sticker AS STRING * 12
    NextTrans AS LONG
    OperNum   AS LONG
    GLInterfaced  AS STRING * 1
    DecalCat AS STRING * 5
    ExtraRoom AS STRING * 97
END TYPE

TYPE DCEditPaymentRecType
    CustNumber AS STRING * 10
    CustName AS STRING * 35
    TranDate AS INTEGER
    Amount AS DOUBLE
    DecalCat AS STRING * 5
    Sticker AS STRING * 12
    Desc AS STRING * 35
    ExpDate AS INTEGER
    makemodel AS STRING * 25
    StateTag AS STRING * 25
    resident AS STRING * 1
    Owner AS STRING * 1
    OldMake AS STRING * 23
    OldDesc AS STRING * 23
    CashAmt   AS DOUBLE
    CheckAmt  AS DOUBLE
    Change    AS DOUBLE
    VehRecord AS LONG
    OperNum   AS LONG
    Pad AS STRING * 41
END TYPE

TYPE DCVehType
    DecalCat AS STRING * 5
    makemodel AS STRING * 25
    StateTag AS STRING * 35
    ExpireDate AS INTEGER
    Sticker AS STRING * 12
    Valid AS STRING * 1         'y/n means is it current
    Active AS STRING * 1        'y/n  n=deleted record
    Notes AS STRING * 40
    Desc AS STRING * 40
    Fee AS SINGLE
    MasterRecord AS LONG
    NextRec AS LONG
    MoreRoom AS STRING * 83
END TYPE

TYPE DCOldVehType
     Make AS STRING * 25
     year AS STRING * 25
     CustRec AS LONG
     MoreRoom AS STRING * 74
END TYPE

TYPE OldCustType
 lname AS STRING * 25
 firstname AS STRING * 25
 add1 AS STRING * 30
 city AS STRING * 20
 ST AS STRING * 2
 zp AS STRING * 5
 PHONE AS STRING * 12
 acct AS SINGLE
 BOOK AS INTEGER
 year AS STRING * 2
 makemodel AS STRING * 18
 idnumber AS STRING * 20
 statelic AS STRING * 11
 ssnum AS STRING * 9
 firsttr AS SINGLE
 lasttr AS SINGLE
 TResident AS STRING * 1
 TOwned AS STRING * 1
 ssnum2 AS STRING * 9
 OtherName AS STRING * 35
 SSNum3 AS STRING * 9
END TYPE

TYPE OldCustTransType
 tdate AS STRING * 6
 tamt AS SINGLE
 tsticker AS STRING * 12
 tyr AS STRING * 4
 tmake AS STRING * 20
 tid AS STRING * 20
 tstlic AS STRING * 15
 tmrec AS SINGLE
 ttype AS INTEGER
 tnexttr AS SINGLE
 tstatus AS STRING * 1
END TYPE
                    


TYPE DCOldCustRecType
    CUSTNUMB     AS STRING * 10
    SORTNAME     AS STRING * 10
    BILLNAME     AS STRING * 35
    ADDRESS1     AS STRING * 35
    ADDRESS2     AS STRING * 35
    city         AS STRING * 20
    STATE        AS STRING * 2
    ZIPCODE      AS STRING * 10
    RoomtoGrow   AS STRING * 353
END TYPE


