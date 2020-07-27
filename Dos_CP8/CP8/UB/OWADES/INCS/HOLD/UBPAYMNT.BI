TYPE PaidOwedType
    AMTOWE1  AS DOUBLE
    AMTPD1   AS DOUBLE
END TYPE

TYPE UBPaymentRecType
    OPERNUM         AS INTEGER
    PAYDATE         AS INTEGER
    CUSTACCT        AS LONG
    CUSTNAME        AS STRING * 24
    CUSTADDR        AS STRING * 24
    CUSTCMNT        AS STRING * 32
'052598 Added tax exmpt flag, changed cust comment len to 32
    TaxExempt       AS STRING * 1
    AMTOWED         AS DOUBLE
    TENDERTY        AS STRING * 12
    CASHAMT         AS DOUBLE
    CHKAMT          AS DOUBLE
    AMTRECD         AS DOUBLE
    CHANGE          AS DOUBLE
    DESC            AS STRING * 20
    PaidOwed(1 TO 15)   AS PaidOwedType
    TOTOWED         AS DOUBLE
    AMTPAID         AS DOUBLE
    
END TYPE

TYPE PayListType
  CustRec   AS LONG
  ListRec   AS LONG
END TYPE

'TYPE CMOperRecType
'    OperatorNumber AS INTEGER
'    OperatorName AS STRING * 30
'    OperatorPassword AS STRING * 8
 '   NotUsed AS STRING * 24
'END TYPE

'TYPE MiscCodeRecType
'    MiscCode AS STRING * 7
'    Description AS STRING * 25
'    GlAcctNumb AS STRING * 14
'    NotUsed AS STRING * 18
'END TYPE
'
'TYPE RMReceiptRecType
'    RecName AS STRING * 25
'    RecAddress AS STRING * 25
'    RecDesc AS STRING * 25
'    RecAmtOwed AS DOUBLE
'    RecPayType AS INTEGER
'    RecCashAmt AS DOUBLE
'    RecCheckAmt AS DOUBLE
'    RecChangeDue AS DOUBLE
'    RecOperator AS STRING * 2
'    RecDate AS STRING * 12
'    RecBlank AS STRING * 5
'END TYPE



