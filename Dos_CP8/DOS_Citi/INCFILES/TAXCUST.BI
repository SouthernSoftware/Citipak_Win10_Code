'Chief, I had to move the old Citipak TB structures out of this include file.
'They are now in the include file "TAXCNVRT.BI" This file is located in our
'include directory

CONST CustRecVerNO = -100

TYPE PINRecType
  pin AS LONG
END TYPE

TYPE TaxMasterType      'Master Default Information in Setup
  Name AS STRING * 35
  ADD1 AS STRING * 35
  ADD2 AS STRING * 35
  ADD3 AS STRING * 35
  TaxSt AS STRING * 2
  TaxForm AS STRING * 20
  CurRate AS SINGLE
  PastRate AS SINGLE
  PenRate AS SINGLE
  RcptPort AS INTEGER
  AcctgMethod AS STRING * 1
  RecptDef    AS STRING * 1
  Padding AS STRING * 252
END TYPE
TYPE TaxInterestDateType
  InterestDate AS INTEGER
END TYPE

TYPE Tax1997PPRateType
   Rate AS SINGLE
END TYPE

TYPE TaxValuesType
  Value    AS DOUBLE
  OthVal   AS DOUBLE
  ExmVal   AS DOUBLE
END TYPE

TYPE TaxCustType
  ACCT       AS LONG
  OPENDATE   AS INTEGER
  FName      AS STRING * 15
  LNAME      AS STRING * 25
  SName      AS STRING * 10
  HPHONE     AS STRING * 14
  WPHONE     AS STRING * 14
  CSSN       AS STRING * 11
  SSSN       AS STRING * 11
  Addr1      AS STRING * 35
  Addr2      AS STRING * 35
  CITY       AS STRING * 20
  State      AS STRING * 2
  ZIP        AS STRING * 10
  Active     AS STRING * 1    'Y if Active N if Inactive
  Interest   AS STRING * 1    'Y/N to Charge Interest
  TaxExempt  AS STRING * 1    'Y/N to Charge Taxes Period
  Penalty    AS STRING * 1    'Y/N to Charge Penalty
  Employer   AS STRING * 25
  Bankrupt   AS STRING * 1    'Y/N to Charge Collect/Adv
  TownShip   AS STRING * 25
'032400
  LateNotice AS STRING * 1    'Y/N Allow late notice
  PAD1              AS STRING * 202
'032900 for New Market Va to Track Who Has Been Submitted to DMV
  DMV1999           AS STRING * 1
  DMV2000           AS STRING * 1
  DMV2001           AS STRING * 1
  DMV2002           AS STRING * 1
  CountyAcctString  AS STRING * 18    'County Account in String Format when longer than 8 digit number
  CountyAcct    AS LONG        'County Account Number to Link to County Records if applic
  LastTrans     AS LONG        'Pointer to last transaction
  FirstPropRec  AS LONG        'Pointer to first property rec
  FirstPersRec  AS LONG        'Pointer to first personal rec
  pin           AS LONG        'Cust internal id number.
  Deleted       AS INTEGER     'deleted flag
  FileVer       AS INTEGER     'this is the file struct version number

END TYPE

TYPE WinRevSourceType
  Principle1    AS DOUBLE                 'Va Personal Prop
  Principle2    AS DOUBLE    'For Va Only     Mach/Tools
  Principle3    AS DOUBLE    'For Va Only     Merch Cap
  Principle4    AS DOUBLE    'For Va Only     Farm Equip
  Principle5    AS DOUBLE    'For Va Only     Mobile Homes
  Interest      AS DOUBLE
  Penalty       AS DOUBLE
  Collection    AS DOUBLE
  Future1       AS DOUBLE
  Future2       AS DOUBLE
  Principle1Pd  AS DOUBLE
  Principle2Pd  AS DOUBLE    'For Va Only
  Principle3Pd  AS DOUBLE    'For Va Only
  Principle4Pd  AS DOUBLE    'For Va Only
  Principle5Pd  AS DOUBLE    'For Va Only
  InterestPd    AS DOUBLE
  PenaltyPd     AS DOUBLE
  CollectionPd  AS DOUBLE
  Future1Pd     AS DOUBLE
  Future2Pd     AS DOUBLE
  RevOpt1       AS DOUBLE
  RevOpt1Pd     AS DOUBLE
  RevOpt2       AS DOUBLE
  RevOpt2Pd     AS DOUBLE
  RevOpt3       AS DOUBLE
  RevOpt3Pd     AS DOUBLE
  LateList      AS DOUBLE
  LateListPd    AS DOUBLE
  PrePaidAmt    AS DOUBLE
  PrePaidUsed   AS DOUBLE
  PrePaidBal    AS DOUBLE
  pad           AS STRING * 80
END TYPE



TYPE WasTaxTransactionType
  TransDate    AS INTEGER          'Transaction Date
  TaxYear      AS INTEGER          'Must Contain Full 4 digit Tax Year Here
  TranType     AS INTEGER          '1=Bill 2=Payment 3=Release 4=Interest
                                   '5=Penalty 6=Collection/Ad Cost Billing
                                   '7=AdjustmentDwnBill 8=MiscCost 9=AdjUpBill
                                   '10=DwnAdjPay 11=UpAdjPay
                                   '22=PrePayment 23=Refund Prepayment added 3-25-03
  BillType     AS STRING * 1       'R=Real P=Personal Property C=Combined (NC/
  Amount       AS DOUBLE           'Total Transaction Amount
  Revenue      AS WinRevSourceType    'See Revenue Source Type File above
  Description  AS STRING * 30      'Description of Transaction
  Posted2GL    AS STRING * 1       'I/F to G/L Yes or No
  CustomerRec  AS LONG             'Pointer Back to Customer Record
  LastTrans    AS LONG             'Points to Previous Trans in History
  'actually Previous pointer
  BelongTo     AS LONG             'Points to Record of Bill this Transaction
  DMVSubmitted AS STRING * 1       'Y if Sent to DMV
  DMVBatch     AS INTEGER          'Records which batch contained the DMV Tran
  Altered      AS INTEGER          'Flag <> 0  If TR altered at any time
   ' Padding      As String * 123     'Allow for Future Expansio
   'changed padding 123 above on 3-25-03 to allow flag to indicate
   'applied prepayment on regular payment transaction
  FromPrePay   AS STRING * 1       'Y if from Prepayment Balance
  Padding      AS STRING * 74     '
  PersPin      AS STRING * 20 'added for 2.05
  RealPin      AS STRING * 20 'added for 2.05
  CustPin      AS LONG 'added for 2.05
  InternalPin  AS LONG
  DiscXDate    AS INTEGER 'added for 2.05
  DiscAmt      AS DOUBLE 'added for 2.05
  OperNum      AS INTEGER
  CntyPara     AS STRING * 20
  CyclPara     AS STRING * 20
  TShpPara     AS STRING * 25
END TYPE


TYPE RevSourceType
  Principle1    AS DOUBLE                 'Va Personal Prop
  Principle2    AS DOUBLE    'For Va Only     Mach/Tools
  Principle3    AS DOUBLE    'For Va Only     Merch Cap
  Principle4    AS DOUBLE    'For Va Only     Farm Equip
  Principle5    AS DOUBLE    'For Va Only     Mobile Homes
  Interest      AS DOUBLE
  Penalty       AS DOUBLE
  Collection    AS DOUBLE
  Future1       AS DOUBLE
  Future2       AS DOUBLE
  Principle1Pd  AS DOUBLE
  Principle2Pd  AS DOUBLE    'For Va Only
  Principle3Pd  AS DOUBLE    'For Va Only
  Principle4Pd  AS DOUBLE    'For Va Only
  Principle5Pd  AS DOUBLE    'For Va Only
  InterestPd    AS DOUBLE
  PenaltyPd     AS DOUBLE
  CollectionPd  AS DOUBLE
  Future1Pd     AS DOUBLE
  Future2Pd     AS DOUBLE
END TYPE

TYPE TaxTransactionType
  TransDate    AS INTEGER          'Transaction Date
  TaxYear      AS INTEGER          'Must Contain Full 4 digit Tax Year Here
  TranType     AS INTEGER          '1=Bill 2=Payment 3=Release 4=Interest
                                   '5=Penalty 6=Collection/Ad Cost Billing
                                   '7=Adjustment
  BillType     AS STRING * 1       'R=Real P=Personal Property C=Combined (NC/GA)
  Amount       AS DOUBLE           'Total Transaction Amount
  Revenue      AS RevSourceType    'See Revenue Source Type File above
  Description  AS STRING * 30      'Description of Transaction
  Posted2GL    AS STRING * 1       'I/F to G/L Yes or No
  CustomerRec  AS LONG             'Pointer Back to Customer Record
  LastTrans    AS LONG             'Points to Previous Trans in History
  'actually Previous pointer
  BelongTo     AS LONG             'Points to Record of Bill this Transaction Belongs to : 'Will be 0 for Bill
  DMVSubmitted AS STRING * 1       'Y if Sent to DMV
  DMVBatch     AS INTEGER          'Records which batch contained the DMV Transmission
  Altered      AS INTEGER          'Flag <> 0  If TR altered at any time
  Padding      AS STRING * 123     'Allow for Future Expansion
END TYPE

TYPE InterestRecType
  CustRec            AS LONG                 'Acct #
  CustName           AS STRING * 40
  TaxYear            AS INTEGER
  Amount             AS DOUBLE
  BillNumber         AS STRING * 10
  CurYear            AS INTEGER
'end of form
  BillRec            AS LONG
  DelFlag            AS INTEGER
'  PropRec            AS LONG
  InfoTxt            AS STRING * 30
  Padding            AS STRING * 129
END TYPE

TYPE TaxMTransactionType
  Account      AS LONG
  TransDate    AS INTEGER
  TaxYear      AS INTEGER
  Desc         AS STRING * 30
  TaxAmount    AS DOUBLE
  IntAmount    AS DOUBLE
  AdColAmount  AS DOUBLE
  BillType     AS STRING * 1   'R=REAL P=PERS C=COMB
  SName        AS STRING * 30
  TName        AS STRING * 30
  MHTax        AS DOUBLE
  MCTax        AS DOUBLE
  FETax        AS DOUBLE
  MTTax        AS DOUBLE
  Deleted      AS INTEGER
  Padding      AS STRING * 94
END TYPE

TYPE MortCodeRecType
  MortCode AS STRING * 8
  BName    AS STRING * 32
  ADD1     AS STRING * 32
  ADD2     AS STRING * 32
  ADD3     AS STRING * 32
  Contact  AS STRING * 32
  PHONE    AS STRING * 14
  pad      AS STRING * 254
END TYPE


