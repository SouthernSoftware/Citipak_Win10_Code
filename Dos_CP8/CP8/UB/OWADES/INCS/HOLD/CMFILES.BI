TYPE CMOperRecType
    OperatorNumber AS INTEGER
    OperatorName AS STRING * 30
    OperatorPassword AS STRING * 8
    NotUsed AS STRING * 24
END TYPE

TYPE MiscCodeRecType
    MiscCode AS STRING * 7
    Description AS STRING * 25
    GlAcctNumb AS STRING * 14
    NotUsed AS STRING * 18
END TYPE

TYPE RMReceiptRecType
    RecName AS STRING * 25
    RecAddress AS STRING * 25
    RecDesc AS STRING * 25
    RecAmtOwed AS DOUBLE
    RecPayType AS INTEGER
    RecCashAmt AS DOUBLE
    RecCheckAmt AS DOUBLE
    RecChangeDue AS DOUBLE
    RecDate AS STRING * 10
    RecOperator AS STRING * 2
    RecptNumber AS SINGLE
    RecBlank AS STRING * 1
    RecBalance AS SINGLE

END TYPE

TYPE CMTransRecType
    TransDate    AS INTEGER
    TransAmount  AS DOUBLE
    TransCash    AS DOUBLE
    TransCheck   AS DOUBLE
    TransAmtOwed AS DOUBLE
    TransDesc    AS STRING * 25
    TransSource  AS INTEGER            '1-Misc 2-Utility 3-Tax 4-License
                                       '5-decal
    TransName    AS STRING * 25
    TransAcctNum AS LONG               'Holds Master Acct Record Number in Module
    TransDetNum  AS LONG               'Holds Record Number of Transaction Detail in Module
    TransRevAmt(1 TO 15) AS DOUBLE
    TransOperNum AS LONG
    Trans2GL      AS STRING * 1
    TransPad     AS STRING * 25
END TYPE

TYPE CMConfigType
    TOWNNAME AS STRING * 30
    CASHACCT AS STRING * 14
    LPTPORT  AS STRING * 1
    PRNDEFYN AS STRING * 1
    ENDMSG   AS STRING * 30
END TYPE

