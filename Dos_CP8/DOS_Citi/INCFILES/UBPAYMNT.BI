TYPE PaidOwedType
    AMTOWE1  AS DOUBLE
    AMTPD1   AS DOUBLE
END TYPE

TYPE oUBPaymentRecType
    OPERNUM         AS INTEGER
    PAYDATE         AS INTEGER
    CUSTACCT        AS LONG
    CUSTNAME        AS STRING * 24
    CUSTADDR        AS STRING * 24
    CUSTCMNT        AS STRING * 32
'052598 Added tax exempt flag changed cust comment len to 32
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
'112801 Added cust status changed desc to 19
END TYPE


TYPE UBPaymentRecType
    OPERNUM         AS INTEGER
    PAYDATE         AS INTEGER
    CUSTACCT        AS LONG
    CUSTNAME        AS STRING * 24
    CUSTADDR        AS STRING * 24
    CUSTCMNT        AS STRING * 32
'052598 Added tax exempt flag changed cust comment len to 32
    TaxExempt       AS STRING * 1
    AMTOWED         AS DOUBLE
    TENDERTY        AS STRING * 12
    CASHAMT         AS DOUBLE
    CHKAMT          AS DOUBLE
    AMTRECD         AS DOUBLE
    CHANGE          AS DOUBLE
    DESC            AS STRING * 19
    PaidOwed(1 TO 15)   AS PaidOwedType
    TOTOWED         AS DOUBLE
    AMTPAID         AS DOUBLE
'112801 Added cust status changed desc to 19
    Status          AS STRING * 1
END TYPE

TYPE PayListType
  CustRec   AS LONG
  ListRec   AS LONG
END TYPE

TYPE BookTotalType
  Count   AS LONG
  Cash    AS DOUBLE
  Check   AS DOUBLE
END TYPE

TYPE LockBoxRecType
  AcctNum     AS STRING * 6     '1-6
  Amount      AS STRING * 10    '7-16
  TenderType  AS STRING * 2     '17/18 CH for Check, CA for cash
  ChkNum      AS STRING * 8     '19/26 chk num
  PaymentDate AS STRING * 10    '27-36 01/01/1999
  Fill1       AS STRING * 44    '37/80
  CrLf        AS STRING * 2     '81/82 CrLf
END TYPE

'First Citizens

TYPE LockBoxRecTypeFC
  AcctNum     AS STRING * 8     '1-6   Utility Billing Account Number
  Amount      AS STRING * 10    '7-16  Decimal assumed (2.00 = 200) in record
  TenderType  AS STRING * 2     '17-18 CH for Check, CA for cash
  ChkNum      AS STRING * 8     '19-26 Check Number
  PaymentDate AS STRING * 10    '27-36 Payment Date. Formate: MM-DD-YYYY
  Blank       AS STRING * 42    '37-80 N/A
  CrLf        AS STRING * 2     '81-82 CR+LF
END TYPE

