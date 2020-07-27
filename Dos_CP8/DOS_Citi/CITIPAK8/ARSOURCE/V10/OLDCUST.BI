TYPE oARCustRecType
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
    RoomtoGrow   AS STRING * 173
END TYPE
