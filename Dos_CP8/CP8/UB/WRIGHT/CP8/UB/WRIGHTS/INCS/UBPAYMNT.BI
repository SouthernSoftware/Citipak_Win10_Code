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

TYPE BookTotalType
  Count   AS LONG
  Cash    AS DOUBLE
  Check   AS DOUBLE
END TYPE

