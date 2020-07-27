TYPE TaxCustType
  Acct       AS LONG
  OPENDATE   AS INTEGER
'  FName      As String * 15
'  LName      As String * 25
  CustName   AS STRING * 50
  SName      AS STRING * 10
  HPHONE     AS STRING * 14
  WPHONE     AS STRING * 14
  CSSN       AS STRING * 11
  OSSN       AS STRING * 11
  Addr1      AS STRING * 35
  Addr2      AS STRING * 35
  City       AS STRING * 20
  State      AS STRING * 2
  Zip        AS STRING * 10
  Active     AS STRING * 1    'Y if Active N if Inactive
  Interest   AS STRING * 1    'Y/N to Charge Interest
  TaxExempt  AS STRING * 1    'Y/N to Charge Taxes Period
  Penalty    AS STRING * 1    'Y/N to Charge Penalty
  Employer   AS STRING * 25
  Bankrupt   AS STRING * 1    'Y/N to Charge Collect/Adv
  TownShip   AS STRING * 25
'032400
  LateNotice AS STRING * 1    'Y/N Allow late notice
'*  PAD1              As String * 202
'*Change Pad1 from 202 and added PrePayment Balance field
'*Also pointer to prepay transaction
'*added these 2 fields on 3/25/03 pks
  PrePayBal    AS DOUBLE
  PrePayTrans  AS LONG
'032900 for New Market Va to Track Who Has Been Submitted to DMV
'  DMV1999           As String * 1'remmed out on 3/2/05
'  DMV2000           As String * 1'remmed out on 3/2/05
'  DMV2001           As String * 1'remmed out on 3/2/05
'  DMV2002           As String * 1'remmed out on 3/2/05
  CountyAcctString  AS STRING * 18    'County Account in String Format when lo
  CountyAcct    AS LONG        'County Account Number to Link to County Record
  LastTrans     AS LONG        'Pointer to last transaction
  FirstPropRec  AS LONG        'Pointer to first property rec
  FirstPersRec  AS LONG        'Pointer to first personal rec
  PIN           AS LONG        'Cust internal id number.
  Deleted       AS INTEGER     'deleted flag
  FileVer       AS INTEGER     'this is the file struct version number
  OptSrchDesc   AS STRING * 15 '3/1/05
  ServiceAdd    AS STRING * 35
  DrvrsLic      AS STRING * 10
  DeliveryPt    AS STRING * 2
  PostalRt     AS STRING * 4
  Cycle        AS LONG
  CycleName    AS STRING * 20
  County4BillNum AS LONG 'used as option for billing
  County4BillName AS STRING * 20
  Pad1         AS STRING * 190  '*remainder after additional fields
END TYPE

TYPE PropertyRecType
    RealPin  AS STRING * 20
    PROPDATE AS INTEGER
    GISPOS   AS STRING * 20
    Map      AS STRING * 6
    BLOCK    AS STRING * 6
    LOTNUMB  AS STRING * 6
    LOTACRE  AS STRING * 1
    PropSize AS DOUBLE
    PROPDISC AS STRING * 1
    LateList AS STRING * 1
    OptRev1Chrg AS INTEGER
    OptRev2Chrg AS INTEGER
    OptRev3Chrg AS INTEGER
    TownShip AS STRING * 30
    MORTCODE AS STRING * 8
    PROPVALU AS DOUBLE
    EXMPSENI AS DOUBLE
    EXMPOTHR AS DOUBLE
    PROPNOT1 AS STRING * 31
    PROPNOT2 AS STRING * 31
    PROPNOT3 AS STRING * 31
    Fill1    AS STRING * 4
    CustPin  AS LONG
    NextRec  AS LONG
    LastYrPrinted AS INTEGER
    Deleted  AS INTEGER
    PropAddr AS STRING * 30
    InternalPin AS LONG
    LienYN AS STRING * 1
    LienDesc AS STRING * 30
    Mock AS STRING * 1
    Image AS STRING * 10
    OptSearch AS STRING * 20
    ICPDesc AS STRING * 15
    Blank  AS STRING * 55
END TYPE

TYPE NCPersonalRecType
   PropPin  AS STRING * 20
   PROPDATE AS INTEGER
   PersVal  AS DOUBLE
   MHValue  AS DOUBLE
   MCValue  AS DOUBLE
   CVALUE   AS DOUBLE
   MTValue  AS DOUBLE
   EXMPSENI AS DOUBLE
   EXMPOTHR AS DOUBLE
   DISCOV   AS STRING * 1
   LateList AS STRING * 1
   DESC1    AS STRING * 30
   DESC2    AS STRING * 30
   DESC3    AS STRING * 30
   Desc4    AS STRING * 30
   Desc5    AS STRING * 30
'end of form
   CustPin        AS LONG
   NextRec        AS LONG
   LastYrPrinted  AS INTEGER
   Deleted        AS INTEGER
   VehTaxYear     AS INTEGER
   DMVSubmitted   AS STRING * 1
   Blank          AS STRING * 117
   InternalPin    AS LONG
END TYPE

TYPE PersonalRecType
  PropPin  AS STRING * 20
  PROPDATE AS INTEGER
  PersVal  AS DOUBLE
  MHValue  AS DOUBLE
  MCValue  AS DOUBLE
  CVALUE   AS DOUBLE
  MTValue  AS DOUBLE
  EXMPSENI AS DOUBLE
  EXMPOTHR AS DOUBLE
  DISCOV   AS STRING * 1
  LateList AS STRING * 1
  DESC1    AS STRING * 30
  DESC2    AS STRING * 30
  DESC3    AS STRING * 30
  Desc4    AS STRING * 30
  Desc5    AS STRING * 30
  CustPin        AS LONG
  NextRec        AS LONG
  LastYrPrinted  AS INTEGER
  Deleted        AS INTEGER
  VehTaxYear     AS INTEGER
  DMVSubmitted   AS STRING * 1
  InternalPin    AS LONG
  TaxBillYear    AS INTEGER
  PPTRAYN        AS STRING * 1
  Prorate        AS STRING * 1
  ProrateVal     AS INTEGER
  Vin            AS STRING * 25
  MakeMod        AS STRING * 25
  Weight         AS DOUBLE
  ModYear        AS INTEGER
  OptRev1Chrg    AS INTEGER
  OptRev2Chrg    AS INTEGER
  OptRev3Chrg    AS INTEGER
  Blank          AS STRING * 105
END TYPE

