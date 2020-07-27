
TYPE TBREKENCust
  AccountNumber     AS STRING * 5
  MapNumber         AS STRING * 14
  District          AS STRING * 2
  CustName          AS STRING * 35
  Address1          AS STRING * 35
  Address2          AS STRING * 35
  Address3          AS STRING * 35
  Zipcode           AS STRING * 5
  PropDesc1         AS STRING * 25
  PropDesc2         AS STRING * 25
  PropDesc3         AS STRING * 25
  PropDesc4         AS STRING * 25
  Zone              AS STRING * 2
  ClassCode         AS STRING * 2
  Tract             AS STRING * 18
  SqFtSize          AS STRING * 8
  Acres             AS STRING * 12
  Comment1          AS STRING * 20
  Comment2          AS STRING * 20
  Comment3          AS STRING * 20
  LandValue         AS STRING * 10
  BldgValue         AS STRING * 10
  LandUseValue      AS STRING * 10
  MPCurYrBillNumber AS STRING * 7       'Municipality Bills (Victoria/Kenbridge)
  MPCurYrTaxAmount  AS STRING * 10

END TYPE
