'

TYPE TBCust1Type
  ptaxamt0 AS SINGLE
  pintamt0 AS SINGLE
  padvamt0 AS SINGLE
  pcolamt0 AS SINGLE
  plateamt0 AS SINGLE
  ptaxnotice0 AS SINGLE
  ptaxamt1 AS SINGLE
  pintamt1 AS SINGLE
  padvamt1 AS SINGLE
  pcolamt1 AS SINGLE
  plateamt1 AS SINGLE
  ptaxnotice1 AS SINGLE
  ptaxamt2 AS SINGLE
  pintamt2 AS SINGLE
  padvamt2 AS SINGLE
  pcolamt2 AS SINGLE
  plateamt2 AS SINGLE
  ptaxnotice2 AS SINGLE
  ptaxamt3 AS SINGLE
  pintamt3 AS SINGLE
  padvamt3 AS SINGLE
  pcolamt3 AS SINGLE
  plateamt3 AS SINGLE
  ptaxnotice3 AS SINGLE
  ptaxamt4 AS SINGLE
  pintamt4 AS SINGLE
  padvamt4 AS SINGLE
  pcolamt4 AS SINGLE
  plateamt4 AS SINGLE
  ptaxnotice4 AS SINGLE
  ptaxamt5 AS SINGLE
  pintamt5 AS SINGLE
  padvamt5 AS SINGLE
  pcolamt5 AS SINGLE
  plateamt5 AS SINGLE
  ptaxnotice5 AS SINGLE
  ptaxamt6 AS SINGLE
  pintamt6 AS SINGLE
  padvamt6 AS SINGLE
  pcolamt6 AS SINGLE
  plateamt6 AS SINGLE
  ptaxnotice6 AS SINGLE
  ptaxamt7 AS SINGLE
  pintamt7 AS SINGLE
  padvamt7 AS SINGLE
  pcolamt7 AS SINGLE
  plateamt7 AS SINGLE
  ptaxnotice7 AS SINGLE
  ptaxamt8 AS SINGLE
  pintamt8 AS SINGLE
  padvamt8 AS SINGLE
  pcolamt8 AS SINGLE
  plateamt8 AS SINGLE
  ptaxnotice8 AS SINGLE
  ptaxamt9 AS SINGLE
  pintamt9 AS SINGLE
  padvamt9 AS SINGLE
  pcolamt9 AS SINGLE
  plateamt9 AS SINGLE
  ptaxnotice9 AS SINGLE
END TYPE

TYPE TBCust2Type
  D1 AS STRING * 40
  D2 AS STRING * 40
  D3 AS STRING * 40
  D4 AS STRING * 40
  D5 AS STRING * 40
  D6 AS STRING * 40
  D7 AS STRING * 40
  D8 AS STRING * 40
  D9 AS STRING * 40
  D10 AS STRING * 40
END TYPE

TYPE TBCustType
  acct      AS STRING * 8
  pin       AS STRING * 4
  firstnamE AS STRING * 25
  LName     AS STRING * 35
  Addr1     AS STRING * 35
  Addr2     AS STRING * 35
  CITY      AS STRING * 20
  State     AS STRING * 2
  ZIP       AS STRING * 10
  Phone     AS STRING * 12
  pdesc     AS STRING * 39
  Map       AS STRING * 22
  SubDiv    AS STRING * 35
  plot      AS SINGLE
  County    AS STRING * 35
  psize     AS SINGLE

  pland     AS DOUBLE
  pbldg     AS DOUBLE
  MH        AS DOUBLE          'Mobile Home    OR NC IS PERS PROP
  fm        AS DOUBLE          'Farm Equip     OR NC IS PERS RES
  ls        AS DOUBLE          'Pers Prop      OR NC IS PERS PUB
  op        AS SINGLE          'Mach Tools     OR NC IS PERSEXP
  me        AS SINGLE          'Merch Cap     OR NC IS SCEXP
  homeexp1  AS SINGLE
  useexp1   AS SINGLE
  CLate     AS SINGLE
  CTax      AS SINGLE
  CInterest AS SINGLE
  CAdv      AS SINGLE
  CCol      AS SINGLE
  CNotice   AS SINGLE
  ftr       AS SINGLE
  ltr       AS SINGLE
  SSN       AS STRING * 11
  late      AS STRING * 1
 END TYPE

TYPE TBCustTypeShift
  BLK       AS STRING * 341
  MH        AS DOUBLE          'Mobile Home    OR NC IS PERS PROP
  fm        AS DOUBLE          'Farm Equip     OR NC IS PERS RES
  ls        AS DOUBLE          'Pers Prop      OR NC IS PERS PUB
  op        AS SINGLE          'Mach Tools     OR NC IS PERSEXP
  me        AS SINGLE          'Merch Cap     OR NC IS SCEXP
  homeexp   AS SINGLE
  useexp    AS SINGLE
  CLate     AS SINGLE
  CTax      AS SINGLE
  CInterest AS SINGLE
  CAdv      AS SINGLE
  CCol      AS SINGLE
  CNotice   AS SINGLE
  ftr       AS SINGLE
  ltr       AS SINGLE
  SSN       AS STRING * 11
  late      AS STRING * 1
END TYPE

TYPE RAMAST
  BLK AS STRING * 17
  Name AS STRING * 35
  Add1 AS STRING * 30
  ADD2 AS STRING * 30
  CITY AS STRING * 31
  ZIP AS STRING * 5
  BLK1 AS STRING * 1
  Zip1 AS STRING * 4
  BLK2 AS STRING * 10
  DESC1 AS STRING * 25
  DESC2 AS STRING * 22
  DESC3 AS STRING * 28
  Map AS STRING * 17
  BLK4 AS STRING * 93
  CLASS AS STRING * 2
  BlkVal AS STRING * 10
  LandVal AS STRING * 10
  BldgVal AS STRING * 10
  BLK5 AS STRING * 295
  SSN AS STRING * 9
END TYPE

TYPE RABILL
  FYear AS STRING * 2
  TYear AS STRING * 2
  CustRec AS STRING * 5
  filler AS STRING * 1
  billrec AS STRING * 3
  filler2 AS STRING * 12
  DueDate AS STRING * 6
  TaxDue AS STRING * 11
  late AS STRING * 1
  filler3 AS STRING * 22
  AmtDue AS STRING * 11
  filler4 AS STRING * 89
END TYPE

TYPE Oldtbcustdat  '256 byte file
  LName     AS STRING * 30
  firstnamE AS STRING * 25
  Addr1     AS STRING * 30
  Addr2     AS STRING * 30
  CITY      AS STRING * 20
  State     AS STRING * 2
  ZIP       AS STRING * 5
  Phone     AS STRING * 12
  pdesc     AS STRING * 20
  Map       AS STRING * 15
  pland     AS DOUBLE
  mv        AS DOUBLE
  SrCit     AS DOUBLE
  TaxType   AS STRING * 2
  acct      AS STRING * 12
  CURBAL    AS DOUBLE
  PastDue   AS DOUBLE
  ftr       AS SINGLE
  ltr       AS SINGLE
END TYPE

TYPE OldTBCUST1DAT              '64 byte record tbcust1.dat
  CurYrTax    AS DOUBLE
  PastYrTax   AS DOUBLE
  LateList    AS SINGLE
  TaxType     AS STRING * 1
  SSN         AS STRING * 11
END TYPE

TYPE HamletTaxRec
  acct AS STRING * 9
  Name AS STRING * 30
  Add1 AS STRING * 30
  ADD2 AS STRING * 30
  ADD3 AS STRING * 30
  Add4 AS STRING * 30
  ZIP AS STRING * 9
  PropDesc AS STRING * 30
  Lots AS STRING * 5
  REAL AS STRING * 8
  Pers AS STRING * 7
  Discnt AS STRING * 5
  nu AS STRING * 43
  SSN AS STRING * 9
  nu1 AS STRING * 282
  SSN1 AS STRING * 9
END TYPE

TYPE HamletRecType
  TownShip AS STRING * 3
  CityCode AS STRING * 2
  LandFill AS STRING * 6
  SCEx AS STRING * 6
  OthEx AS STRING * 6
  ZipCode AS STRING * 6
  Name1 AS STRING * 30
  Name2 AS STRING * 30
  Add1 AS STRING * 30
  CITY AS STRING * 21
  flag AS STRING * 1
  late AS STRING * 1
  pin AS STRING * 15
  LandDesc AS STRING * 24
  Acres AS STRING * 8
  AorL AS STRING * 1              'A for Acres  L for Lots
  MHomes AS STRING * 2
  MtrVeh AS STRING * 2
  MortCode AS STRING * 2
  RealValue AS STRING * 8
  HouseHoldValue AS STRING * 8
  InventoryValue AS STRING * 8
  MachValue AS STRING * 8
  NOTUSED AS STRING * 8
  VehValue AS STRING * 8
  AcctNumb AS STRING * 8
  dogs AS STRING * 2
  Transferred AS STRING * 1
  CLASS AS STRING * 1
  DeedDate AS STRING * 6
  DeedBook AS STRING * 3
  DeedPage AS STRING * 3
  Notes AS STRING * 30
  PlatBook AS STRING * 6
  Recn AS STRING * 6
  Blank1 AS STRING * 6
  LandValue AS STRING * 8
  BldgValue AS STRING * 8
  OtherValue AS STRING * 8
  Dist AS STRING * 2
  FireDist AS STRING * 2
  DeedCode AS STRING * 1
  SSN AS STRING * 9
END TYPE

TYPE chilrealtype
  BLK1    AS STRING * 6
  CtyAcct AS STRING * 5
  NME     AS STRING * 30
  Add1    AS STRING * 30
  ADD2    AS STRING * 30
  ADD3    AS STRING * 30
  ZIP     AS STRING * 5
  BLK2    AS STRING * 4
  Map     AS STRING * 15
  Siz     AS STRING * 18
  DESC1   AS STRING * 25
  DESC2   AS STRING * 25
  DESC3   AS STRING * 25
  BLK3    AS STRING * 11
  LandVal AS STRING * 10
  BlkVal  AS STRING * 9
  BldgVal AS STRING * 9
  bk      AS STRING * 12
  PAGE    AS STRING * 5
END TYPE

TYPE BrookNealType
  Ticket AS STRING * 6
  AcctNo AS STRING * 8
  SSN    AS STRING * 9
  SSEx   AS STRING * 1
  Owner  AS STRING * 1
  LName  AS STRING * 45
  FName  AS STRING * 33
  MInt   AS STRING * 1
  Add1   AS STRING * 35
  ADD2   AS STRING * 35
  ADD3   AS STRING * 35
  CITY   AS STRING * 22
  State  AS STRING * 2
  ZIP    AS STRING * 5
  Zip1   AS STRING * 4
  SSN1   AS STRING * 9
  bk     AS STRING * 1
  CName  AS STRING * 79
  nu     AS STRING * 5
  PPValue AS STRING * 8
  PPRate AS STRING * 3
  PPAmt AS STRING * 9
  PPRebate AS STRING * 8
  PPNet AS STRING * 8
  MTValue AS STRING * 8
  MTRate AS STRING * 3
  MTAmt AS STRING * 9
  MCValue AS STRING * 8
  MCRate AS STRING * 3
  MCAmt AS STRING * 9
  MHValue AS STRING * 9
  MHRate AS STRING * 3
  MHAmt AS STRING * 8
  TicTotal AS STRING * 40
  VIN AS STRING * 22
  Seq AS STRING * 3
  Year AS STRING * 4
  MAKE AS STRING * 20
  MODEL AS STRING * 15
  Qual AS STRING * 1
  nu1 AS STRING * 8
  VALU AS STRING * 8
END TYPE

TYPE WarsawREType
  Parcel AS STRING * 20
  Name AS STRING * 40
  Add1 AS STRING * 35
  ADD2 AS STRING * 35
  CITY AS STRING * 20
  State AS STRING * 2
  ZIP AS STRING * 5
  BLK1 AS STRING * 17
  CLASS AS STRING * 1
  BLK2 AS STRING * 21
  Acres AS STRING * 8
  LandVal AS STRING * 22
  LandEx  AS STRING * 22
  BldgVal AS STRING * 22
  DESC1 AS STRING * 13
  DESC2 AS STRING * 15
END TYPE

TYPE ElktonPPType
  AcctNumber AS STRING * 8
  Name AS STRING * 35
  Add1 AS STRING * 35
  ADD2 AS STRING * 33
  CITY AS STRING * 36
  ZIP AS STRING * 10
  TaxDue AS STRING * 8
  BillNo AS STRING * 6
  BLK AS STRING * 1
  SS1 AS STRING * 11
  BLK1 AS STRING * 1
  SS2 AS STRING * 11
END TYPE

TYPE ElkType
  BLK AS STRING * 10
  Map AS STRING * 6
  BLK1 AS STRING * 4
  Deed AS STRING * 9
  BLK2 AS STRING * 3
  Nam AS STRING * 35
  Add1 AS STRING * 35
  ADD2 AS STRING * 36
  ZIP AS STRING * 5
  BLK3 AS STRING * 4
  CITY AS STRING * 35
  DESC1 AS STRING * 35
  DESC2 AS STRING * 35
  DESC3 AS STRING * 35
  Siz AS STRING * 8
  LandVal AS STRING * 10
  BldgVal AS STRING * 10
  DisVal AS STRING * 10
  TaxAmt AS STRING * 12
END TYPE

TYPE AlbertaREType
 BLK AS STRING * 5
 acct AS STRING * 10
 BLK1 AS STRING * 2
 NM   AS STRING * 35
 Add1 AS STRING * 30
 ADD2 AS STRING * 30
 ADD3 AS STRING * 31
 ZIP AS STRING * 5
 BLK2 AS STRING * 1
 ZIP4 AS STRING * 4
 BLK3 AS STRING * 10
 DESC1 AS STRING * 25
 DESC2 AS STRING * 25
 DESC3 AS STRING * 26
 Map   AS STRING * 16
 BLK5 AS STRING * 94
 CLASS AS STRING * 1
 BLK6 AS STRING * 10
 LAND AS STRING * 10
 BLDG AS STRING * 10
 BLK7 AS STRING * 16
 BOOK AS STRING * 4
 PAGE AS STRING * 4
 END TYPE

TYPE ALBERTAPPTYPE
  BLK AS STRING * 5
  acct AS STRING * 8
  BLK1 AS STRING * 4
  TYP  AS STRING * 1
  NM  AS STRING * 35
  Add1 AS STRING * 30
  ADD2 AS STRING * 30
  ADD3 AS STRING * 31
  ZIP AS STRING * 5
  BLK3 AS STRING * 4
  BLK4 AS STRING * 20
  SS1 AS STRING * 9
  SS2 AS STRING * 10
  BLK5 AS STRING * 433
  CLASS AS STRING * 2
  VIN AS STRING * 20
  BLK31 AS STRING * 8
  AMT AS STRING * 10
  BLK41 AS STRING * 11
  MAKE AS STRING * 10
  MODEL AS STRING * 15
  BODY AS STRING * 16
  YR AS STRING * 4
  BLK51 AS STRING * 57
  BLK39 AS STRING * 392
  DMV AS STRING * 1
END TYPE

TYPE tbctype
  LName     AS STRING * 30
  firstnamE AS STRING * 25
  Addr1     AS STRING * 30
  Addr2     AS STRING * 30
  CITY      AS STRING * 20
  State     AS STRING * 2
  ZIP       AS STRING * 5
  Phone     AS STRING * 12
  Prop      AS STRING * 20
  Map       AS STRING * 15
  REAL      AS DOUBLE
  Pers      AS DOUBLE
  SrCit     AS DOUBLE
  CarTag    AS INTEGER
  acct      AS STRING * 12
  CURBAL    AS DOUBLE
  PastDue   AS DOUBLE
  ftr       AS SINGLE
  ltr       AS SINGLE
END TYPE

TYPE TBCType1
  CurYrTax AS DOUBLE
  PastYrTax AS DOUBLE
  LateList AS SINGLE
  TaxType AS STRING * 1
  SSN AS STRING * 11
  Dog AS INTEGER
END TYPE

TYPE TBCType2
  ptaxamt0 AS SINGLE
  pintamt0 AS SINGLE
  ptaxamt1 AS SINGLE
  pintamt1 AS SINGLE
  ptaxamt2 AS SINGLE
  pintamt2 AS SINGLE
  ptaxamt3 AS SINGLE
  pintamt3 AS SINGLE
  ptaxamt4 AS SINGLE
  pintamt4 AS SINGLE
  ptaxamt5 AS SINGLE
  pintamt5 AS SINGLE
  ptaxamt6 AS SINGLE
  pintamt6 AS SINGLE
  ptaxamt7 AS SINGLE
  pintamt7 AS SINGLE
  ptaxamt8 AS SINGLE
  pintamt8 AS SINGLE
  ptaxamt9 AS SINGLE
  pintamt9 AS SINGLE
END TYPE

