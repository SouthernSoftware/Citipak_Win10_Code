TYPE TAXGLAcctRecType
  TAXYEAR       AS INTEGER        'protected
  TaxDBAcct     AS STRING * 14
  TaxCRAcct     AS STRING * 14
  IntDBAcct     AS STRING * 14
  IntCRAcct     AS STRING * 14
  AdvDBAcct     AS STRING * 14
  AdvCRAcct     AS STRING * 14
  Fill1         AS STRING * 1     'protected
END TYPE

TYPE TaxAcctsType
  TaxAcct(1 TO 31)   AS TAXGLAcctRecType
  '1980 thru 2010  Inclusive
END TYPE

