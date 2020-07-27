TYPE TaxInfoType
  TaxAmt AS DOUBLE
  a AS INTEGER
END TYPE
TYPE InvoiceTaxInfoType
    AutoDist AS STRING * 1

    InvTax(1 TO 2) AS TaxInfoType
'    STAACCT  AS STRING * 16
'    STATAX   AS DOUBLE
'
'    CTYACCT  AS STRING * 16
'    CTYTAX   AS DOUBLE
END TYPE

