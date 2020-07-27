Attribute VB_Name = "TaxSTRUCT"
Option Explicit

Type JGLAcctIdxType
  AcctNum As String * 14
  RecNo   As Integer
End Type


Type WinTAXGLAcctRecType
  TaxYear       As Integer        'protected
  TaxDBAcct     As String * 14
  TaxCRAcct     As String * 14
  IntDBAcct     As String * 14
  IntCRAcct     As String * 14
  AdvDBAcct     As String * 14
  AdvCRAcct     As String * 14
  Fill1         As String * 1     'protected
  LtLstDBAcct   As String * 14
  LtLstCRAcct   As String * 14
  Opt1DBAcct    As String * 14
  Opt1CRAcct    As String * 14
  Opt2DBAcct    As String * 14
  Opt2CRAcct    As String * 14
  Opt3DBAcct    As String * 14
  Opt3CRAcct    As String * 14
End Type

Type TaxAcctsType
  TaxAcct(1 To 51) As WinTAXGLAcctRecType
End Type

Type TaxGLPrePayType
  TaxDBAcct     As String * 14
  TaxCRAcct     As String * 14
  Filler        As String * 70
End Type

Type PINRecType
  PIN As Long
End Type

Type TaxMasterType      'Master Default Information in Setup
  Name As String * 35
  Add1 As String * 35
  Add2 As String * 35
  'ADD3 As String * 35
  'Change the add3 line to break out individual city,st,zip on 013103.
  City As String * 25
  'use taxst for state in address
  'State As String * 2
  Zip As String * 10
  TaxSt As String * 2
  'TaxForm As String * 20
  'Change taxform above to 2 byte integer
  TaxForm As Integer
  'add lateform 031303
  TaxYear As Integer
  LateForm As Integer
'  pad As String * 16  'left from taxform string of 20
'change above pad to use for following changes as of 3-28-03
'  pad     As String * 3
  WarnInt As String * 1  'Flag to Warn if interest not applied
'  DisFlag As String * 1  'set discount flag if want interest calculated
  MinBill As Double      'amount to not print bills
  'CurRate As Single
  'PastRate As Single
  'PenRate As Single
  'use the 3 rates above (12) for other stuff
 'change rcptport to pad up above - will set printer ports when sign on
  'RcptPort As Integer
  AcctgMethod As String * 1
  'add interface option 031301
  MinTxOpt As Integer '1/26/05 '1) if the taxpayer is charged nothing if
  'their tax bill is equal to or less than this amt...2) the taxpayer is charged at least this
  'amt even if they owe nothing
  TownState As String * 2 '1/26/05
  CurrYrInt As Double '1/26/05
  PastYrInt As Double '1/26/05
  PenPct As Double '1/26/05
  PenIdx As Integer
  CntrlDepYN As String * 1
  PriorYrMltRevYN As String * 1
  OverPayGLNum As String * 14
  PenPrncTaxYN As String * 1
  PenIntYN As String * 1
  PenAdvYN As String * 1
  PenLateLstYN As String * 1
  PenOpt1YN As String * 1
  PenOpt2YN As String * 1
  PenOpt3YN As String * 1
  IntPrncTaxYN As String * 1
  IntIntYN As String * 1
  IntAdvYN As String * 1
  IntLateLstYN As String * 1
  IntOpt1YN As String * 1
  IntOpt2YN As String * 1
  IntOpt3YN As String * 1
  OptRev1 As String * 35
  OptRev2 As String * 35
  OptRev3 As String * 35
  DiscXDate As Integer      'discount amount to calc on payment screen
  DisPct As Double
  OptSrchCust As String * 15
  OptSrchProp As String * 15
  CountyName(1 To 5) As String * 20
  CountyNum(1 To 5) As Integer
  UseCountyYN As String * 1
  RealPersSplit As String * 1
  CycleNum(1 To 5) As Long
  CycleName(1 To 5) As String * 20
  UseCyclesYN As String * 1
  CDCashGL  As String * 14
  CDSubGL  As String * 14
  Padding As String * 189
End Type

Type TaxInterestDateType
  InterestDate As Integer
End Type

Type Tax1997PPRateType
   Rate As Single
End Type

Type TaxValuesType
  Value    As Double
  OthVal   As Double
  ExmVal   As Double
End Type

Type TaxCustType
  Acct       As Long
  OPENDATE   As Integer
'  FName      As String * 15
'  LName      As String * 25
  CustName   As String * 50
  SName      As String * 10
  HPHONE     As String * 14
  WPHONE     As String * 14
  CSSN       As String * 11
  OSSN       As String * 11
  Addr1      As String * 35
  Addr2      As String * 35
  City       As String * 20
  State      As String * 2
  Zip        As String * 10
  Active     As String * 1    'Y if Active N if Inactive
  Interest   As String * 1    'Y/N to Charge Interest
  TaxExempt  As String * 1    'Y/N to Charge Taxes Period
  Penalty    As String * 1    'Y/N to Charge Penalty
  Employer   As String * 25
  Bankrupt   As String * 1    'Y/N to Charge Collect/Adv
  TownShip   As String * 25
'032400
  LateNotice As String * 1    'Y/N Allow late notice
'*  PAD1              As String * 202
'*Change Pad1 from 202 and added PrePayment Balance field
'*Also pointer to prepay transaction
'*added these 2 fields on 3/25/03 pks
  PrePayBal    As Double
  PrePayTrans  As Long
'032900 for New Market Va to Track Who Has Been Submitted to DMV
'  DMV1999           As String * 1'remmed out on 3/2/05
'  DMV2000           As String * 1'remmed out on 3/2/05
'  DMV2001           As String * 1'remmed out on 3/2/05
'  DMV2002           As String * 1'remmed out on 3/2/05
  CountyAcctString  As String * 18    'County Account in String Format when lo
  CountyAcct    As Long        'County Account Number to Link to County Record
  LastTrans     As Long        'Pointer to last transaction
  FirstPropRec  As Long        'Pointer to first property rec
  FirstPersRec  As Long        'Pointer to first personal rec
  PIN           As Long        'Cust internal id number.
  Deleted       As Integer     'deleted flag
  FileVer       As Integer     'this is the file struct version number
  OptSrchDesc   As String * 15 '3/1/05
  ServiceAdd    As String * 35
  DrvrsLic      As String * 10
  DeliveryPt      As String * 2
  PostalRt     As String * 4
  Cycle        As Long
  CycleName    As String * 20
  County4BillNum As Long 'used as option for billing
  County4BillName As String * 20
  Pad1         As String * 190  '*remainder after additional fields
End Type

Type HistRecInfoType
  TranRec    As Long
  TranType   As Integer
  TranDate   As Integer
  BelongTo   As Long
  Printed    As Integer
End Type

Type WinRevSourceType
  Principle1    As Double                 'Va Personal Prop
  Principle2    As Double    'For Va Only     Mach/Tools
  Principle3    As Double    'For Va Only     Merch Cap
  Principle4    As Double    'For Va Only     Farm Equip
  Principle5    As Double    'For Va Only     Mobile Homes
  Interest      As Double
  Penalty       As Double
  Collection    As Double
  Future1       As Double
  Future2       As Double
  Principle1Pd  As Double
  Principle2Pd  As Double    'For Va Only
  Principle3Pd  As Double    'For Va Only
  Principle4Pd  As Double    'For Va Only
  Principle5Pd  As Double    'For Va Only
  InterestPd    As Double
  PenaltyPd     As Double
  CollectionPd  As Double
  Future1Pd     As Double
  Future2Pd     As Double
  RevOpt1       As Double
  RevOpt1Pd     As Double
  RevOpt2       As Double
  RevOpt2Pd     As Double
  RevOpt3       As Double
  RevOpt3Pd     As Double
  LateList      As Double
  LateListPd    As Double
  PrePaidAmt    As Double
  PrePaidUsed   As Double
  PrePaidBal    As Double
  pad           As String * 80
End Type

Type TaxTransactionType
  TransDate    As Integer          'Transaction Date
  TaxYear      As Integer          'Must Contain Full 4 digit Tax Year Here
  TranType     As Integer          '1=Bill 2=Payment 3=Release 4=Interest
                                   '5=Penalty 6=Collection/Ad Cost Billing
                                   '7=AdjustmentDwnBill 8=MiscCost 9=AdjUpBill
                                   '10=DwnAdjPay 11=UpAdjPay
                                   '22=PrePayment 23=Refund Prepayment added 3-25-03
  BillType     As String * 1       'R=Real P=Personal Property C=Combined (NC/
  Amount       As Double           'Total Transaction Amount
  Revenue      As WinRevSourceType    'See Revenue Source Type File above
  Description  As String * 30      'Description of Transaction
  Posted2GL    As String * 1       'I/F to G/L Yes or No
  CustomerRec  As Long             'Pointer Back to Customer Record
  LastTrans    As Long             'Points to Previous Trans in History
  'actually Previous pointer
  BelongTo     As Long             'Points to Record of Bill this Transaction
  DMVSubmitted As String * 1       'Y if Sent to DMV
  DMVBatch     As Integer          'Records which batch contained the DMV Tran
  Altered      As Integer          'Flag <> 0  If TR altered at any time
' Padding      As String * 123     'Allow for Future Expansion
'changed padding 123 above on 3-25-03 to allow flag to indicate
'applied prepayment on regular payment transaction
  FromPrePay   As String * 1       'Y if from Prepayment Balance
  Padding      As String * 74     '
  PersPin      As String * 20 'added for 2.05
  RealPin      As String * 20 'added for 2.05
  CustPin      As Long 'added for 2.05
  InternalPin  As Long
  DiscXDate    As Integer 'added for 2.05
  DiscAmt      As Double 'added for 2.05
  OperNum      As Integer
  CntyPara     As String * 20
  CyclPara     As String * 20
  TShpPara     As String * 25
End Type
Type InterestRecType
  CustRec            As Long                 'Acct #
  CustName           As String * 40
  TaxYear            As Integer
  Amount             As Double
  BillNumber         As String * 10
  CurYear            As Integer
'end of form
  BillRec            As Long
  DelFlag            As Integer
'  PropRec            AS LONG
  InfoTxt            As String * 30
  NewOwnerName       As String * 50
  NewOwnerAcct       As Long
  CustPin            As Long
  RealPin            As String * 20
  PersPin            As String * 20
  Padding            As String * 129
End Type

Type TaxMTransactionType
  Account      As Long
  TransDate    As Integer
  TaxYear      As Integer
  Desc         As String * 30
  TaxAmount    As Double
  IntAmount    As Double
  AdColAmount  As Double
  LateList     As Double
  OptRev1      As Double
  OptRev2      As Double
  OptRev3      As Double
  BillType     As String * 1   'R=REAL P=PERS C=COMB
  SName        As String * 50
  TName        As String * 50
  RealRec      As Long
  PersRec      As Long
  BillNum      As Long
  Class        As String * 1
  Deleted      As Integer
  OverPayUsed  As Double 'if credit balance is applied
  Padding      As String * 94
End Type

Type MortCodeRecType
  MORTCODE As String * 8
  BName    As String * 32
  Add1     As String * 32
  Add2     As String * 32
  Add3     As String * 32
  Contact  As String * 32
  PHONE    As String * 14
'Add deleted field 021003
  Deleted  As Integer
  pad      As String * 252
End Type

Type MortRecType
    MORTCODE As String * 8
    MortRec  As Integer
End Type

Type PINSearchType
  PIN   As String * 20
  Cust  As Long
End Type

'This is Temporary File used for listing customers for selection
Type SortCustList
  Acct    As Long
  LName   As String * 25
  FName   As String * 15
  SSN     As String * 11
  PAddr   As String * 30
  PIN     As Long
End Type

Type SortStruct
  who As String * 14
  RecNum As Integer
End Type

Type PropertyRecType
    RealPin  As String * 20
    PROPDATE As Integer
    GISPOS   As String * 20
    Map      As String * 6
    BLOCK    As String * 6
    LOTNUMB  As String * 6
    LOTACRE  As String * 1
    PropSize As Double
    PROPDISC As String * 1
    LateList As String * 1
    OptRev1Chrg As Integer
    OptRev2Chrg As Integer
    OptRev3Chrg As Integer
    TownShip As String * 30
    MORTCODE As String * 8
    PROPVALU As Double
    EXMPSENI As Double
    EXMPOTHR As Double
    PROPNOT1 As String * 31
    PROPNOT2 As String * 31
    PROPNOT3 As String * 31
    Fill1    As String * 4
    CustPin  As Long
    NextRec  As Long
    LastYrPrinted As Integer
    Deleted  As Integer
    PropAddr As String * 30
    InternalPin As Long
    LienYN As String * 1
    LienDesc As String * 30
    Mock As String * 1
    Image As String * 10
    OptSearch As String * 20
    Blank  As String * 70
End Type

Type PersonalRecType
   PropPin  As String * 20
   PROPDATE As Integer
   PersVal  As Double
   MHVALUE  As Double
   MCVALUE  As Double
   CVALUE   As Double
   MTVALUE  As Double
   EXMPSENI As Double
   EXMPOTHR As Double
   DISCOV   As String * 1
   LateList As String * 1
   DESC1    As String * 30
   DESC2    As String * 30
   DESC3    As String * 30
   Desc4    As String * 30
   Desc5    As String * 30
'end of form
   CustPin        As Long
   NextRec        As Long
   LastYrPrinted  As Integer
   Deleted        As Integer
   VehTaxYear     As Integer
   DMVSubmitted   As String * 1
   Blank          As String * 117
   InternalPin    As Long
End Type

Type TaxBillType
     CustRec            As Long                 'Acct #
     CustName           As String * 40
     CustAdd1           As String * 35
     CustAdd2           As String * 35
     CustAdd3           As String * 35
     CustZip            As String * 10
     RDesc1             As String * 30
     RDesc2             As String * 30
     RealPin            As String * 20
     PersPin            As String * 20
     RealValue          As Double
     PersValue          As Double
     ExptValue          As Double
     RealTaxDue         As Double
     PersTaxDue         As Double
     LateTaxDue         As Double
     TotalBillDue       As Double
     BillNumber         As Long                          'Recpt #
     TaxYear            As Integer
     BillPrinted        As Integer            '-1 = printed
     RealPropRecord     As Long
     PersPropRecord     As Long
     PriorYrBalance     As Double
     RealTaxRate        As Double
     PersTaxRate        As Double
     CustPin            As Long         'additional Protection for relinking
     TownShip           As String * 20
     MORTCODE           As String * 2
     LotOrAcre          As String * 1
     LASize             As String * 9
     MortRec            As Integer
     CarShore           As Double
     RDesc3             As String * 30
     InternalPin        As Long 'added 5/12/05
     OptRevTax1         As Double 'added 5/12/05
     OptRevTax2         As Double 'added 5/12/05
     OptRevTax3         As Double 'added 5/12/05
     OverPayAmt         As Double 'added 5/24/05
     Padding            As String * 105
End Type

Type TaxBillInfoType
    TaxYear  As Integer
    BillNum  As Long
    REALRATE As Double
    PERSRATE As Double
    LATEPCT  As Double
    PRNORDER As String * 20
    CountyPara         As String * 20 'added 5/19/05
    TwnShpPara         As String * 30 'added 5/19/05
    SplitPara          As String * 30 'added 5/19/05
    CyclePara          As String * 20 'added 5/19/05
End Type

Type PaidOwedType
   AmtOwed   As Double
   AmtPaid   As Double
End Type

Type CustPayListType
   CustAcct     As Long
   LastPayRec  As Long
   NumPayRec   As Long
End Type

Type TaxPaymentRecType
    OperNum  As Integer
    payDate  As Integer
    CustAcct As Long
    CustName As String * 24
    CustAddr As String * 24
    AmtOwed  As Double
    TenderTY As String * 14
    CashAmt  As Double
    ChkAmt   As Double
    ChrgAmt  As Double
    DiscAmt  As Double
    AmtRecd  As Double
    Change   As Double
    Desc     As String * 20
    PaidOwed(1 To 7)  As PaidOwedType
    TotOwed  As Double
    AmtPaid  As Double
    TotPaid         As Double
    LastPayRec      As Long          'Pointer to first payment list record
    NumPayRec       As Integer       'Count of payment list records
    CustPin         As Long
    PrePayAmt As Double
End Type
'Type FLen2
'    V As String * 64
'End Type
Type txPayListType
  BillRec       As Long      'Pointer to bill trans rec this payment is for
  BillDate      As Integer 'added for 2.05
  DiscAmt       As Double  'added for 2.05
  DiscXDate     As Integer
  Principle1    As Double
  Interest1     As Double
  Collection    As Double
  LateList      As Double
  OptRev1       As Double
  OptRev2       As Double
  OptRev3       As Double
  TotPaid       As Double    'amount paid to this bill rec (parital payment)

  CustRec       As Long      'backup pointer to cust rec
  PrevListRec   As Long      'pointer to next paylist rec
  TaxYear       As Integer
  Description   As String * 30
  TotOwed       As Double
  PrePayAmt     As Double
End Type

Type GLFundIndexType                'Fund Index : 16 bytes
   FundNum     As String * 4        'Fund Number
   RecNum      As Integer           'Pointer to record
   '*****
End Type

Type GLFundRecType                  'Fund Record Type: 64 bytes
   Deleted     As Integer           'Deleted Flag
   FundNum     As String * 4        'Fund Code
   Title       As String * 30       'Fund Title
   Res         As String * 28       'Reserve for future needs
End Type


Type IFRecType
   AcctNum As String * 9      '9 AS tranacct$
   TRDATE As String * 8       '8 AS trandate$
   Desc As String * 20        '20 AS trandesc$
   CrAmt As Double            '8 AS cramt$
   DrAmt As Double            '8 AS dramt$
   Ref As String * 8          '8 AS detail$
   Src As String * 8          '8 AS source$
   Filler As String * 14      '4 AS nexttr$
   Posted As Integer
End Type


Type TXCustNameIdxType
   CustName As String * 50
   CustRec As Long
End Type

Type SrchNameIdxType
   SearchName As String * 10
   CustRec As Long
End Type

Type InternalPinType
  PIN As Long
End Type


Type PenaltyHandlingType
  PenIdx   As Integer 'of the 6th, 7th or 8th row choose one
  PenDesc  As String * 15 'user defined penalty name
  PenPct   As Double 'penalty percentage
  PenFlat  As Double 'penalty flat rate
  UsePct   As String * 1 'apply percentage only
  UseFlat  As String * 1 'apply flat rate only
  UseBoth  As String * 1 'apply either flat or pct in conjunction with
  'UseHigh or UseLow
  UseHigh  As String * 1 'if using UseBoth then take the higher of either
  'PenPct or PenFlat
  UseLow   As String * 1 'if using UseBoth the take the lower of either
  'PenPct or PenFlat
  AppToRev1 As String * 1 'apply penalty to this revenue
  Rev1Name  As String * 15 'description of Rev1
  AppToRev2 As String * 1
  Rev2Name  As String * 15
  AppToRev3 As String * 1
  Rev3Name  As String * 15
  AppToRev4 As String * 1
  Rev4Name  As String * 15
  AppToRev5 As String * 1
  Rev5Name  As String * 15
  AppToRev6 As String * 1
  Rev6Name  As String * 15
  AppToRev7 As String * 1
  Rev7Name  As String * 15
  AppToRev8 As String * 1
  Rev8Name  As String * 15
End Type

Type TempPayList 'use this as a temporary storage for
'the bills selected for payment but before the payment is saved
  BillRec       As Integer
  CustRec       As Long
  BillPtr       As Long
  BillDate      As Integer   'added for 2.05
End Type

Type TempTaxBillAddOn
  CustRec As Long
  CustName As String * 50
  Type As String * 50
  OldAmt As Double
  NewAmt As Double
End Type

Type TownshipType
  TownShip As String * 30
End Type

Type MessLineType
  Msg As String * 69
End Type

Type TaxMessRecType
  MessLine(1 To 15) As MessLineType
  TaxRec As Long
End Type

Type OptRevRateTablesType
  OptRevNum As Integer
  Desc As String * 20
  Type As String * 1
  StepType As String * 1 'pct or flat rate
  FromAmt(1 To 10) As Double
  ToAmt(1 To 10) As Double
  TaxFAmt(1 To 10) As Double
  TaxPAmt(1 To 10) As Double
  FlatAmt As Double
  Deleted As Boolean
  Cushion As String * 100
End Type

Type RealHistoryType 'new for 2.05
  InternalPin  As Long
  RealPin      As String * 20
  CustPin      As Long
  LastRec      As Long
  Cushion      As String * 80
End Type

Type ManualTaxListType 'designed for selecting a property to tax
  CustPin As Long
  RealPin As String * 20
  RealRec As Long
  CustName As String * 50
  RealAddr As String * 30
  RealVal As Double
  PersPin As String * 20
  PersRec As Long
End Type

Type OptCustIdxType
  OptDesc As String * 20
  CustRec As Long
  CustPin As Long
End Type

Type OptRealIdxType
  OptDesc As String * 20
  RealRec As Long
  RealPin As String * 20
End Type

Type SocSecIdxType
  SSNum As Double
  CustRec As Long
End Type

Type TAXLateLetterType
  Head1    As String * 40
  Head2    As String * 40
  Head3    As String * 40
  Head4    As String * 40
  Head5    As String * 40
  Body(1 To 20) As String * 75
End Type

Type LateListPrintType
  TownName As String * 35
  LateSeqNum As Long
  CustName As String * 50
  Addr1 As String * 35
  Addr2 As String * 35
  City As String * 20
  State As String * 2
  Zip As String * 10
  AdvDate As Integer
  payDate As Integer
  RealValue As Double
  PersValue As Double
  RealExemp As Double
  PersExemp As Double
  TaxYear As Integer
  PrincBal As Double
  IntBal As Double
  AdvBal As Double
  LateListBal As Double
  Opt1Bal As Double
  Opt2Bal As Double
  Opt3Bal As Double
  TotBal As Double
  CurrBal As Double
  PrevBal As Double
  CustAcct As Long
  LtrDate As Integer
  LtrType As String * 1
End Type

Type TxBill1DefaultsType
    TxtHead1 As String * 50
    TxtHead2 As String * 50
    txtOpt1 As String * 40
    TxtOpt2 As String * 40
    TxtOpt3 As String * 40
    TxtOpt4 As String * 40
    txtPgph0 As String * 125
    txtPgph1 As String * 125
    txtPgph2 As String * 125
    txtPgph3 As String * 125
    txtPgph4 As String * 125
    txtPgph5 As String * 125
    txtPgph6 As String * 125
    txtPgph7 As String * 125
    TxtOpt5 As String * 75
    txtHead4 As String * 40
    txtHead5 As String * 40
    txtHead6 As String * 40
    TxtOpt6 As String * 45
    TxtOpt7 As String * 75
    dologo As Integer  '0 for no 1 for yes
End Type

Type TaxBillExportRealType 'words are right set and numbers use formats
  TaxBillNum As String * 7
  CustName As String * 35
  Add1 As String * 35
  Add2 As String * 35
  Add3 As String * 35 'City, State, Zip
  TaxYear As String * 4
  CustAcct As String * 5
  MapNum As String * 14
  PropDesc1 As String * 25
  TAXRATE As String * 4
  LandVal As String * 10
  BldgVal As String * 10
  RealVal As String * 10
  CurrTaxAmt As String * 10
  PropDesc2 As String * 25
  PropDesc3 As String * 25
  TotTaxAmt As String * 10
End Type

Type TaxBillExportPersType
  CustName As String * 25
  Add1 As String * 25
  Add2 As String * 25
  City As String * 20
  State As String * 2
  Zip As String * 10
  CustAcct As String * 6
  SSN1 As String * 12
  SSN2 As String * 12
  DueDate As String * 10
  TotDue As String * 11
  LessRelief As String * 11
  NetDue As String * 11
  RepeatDesc As String * 20
  RepeatID As String * 17
  RepeatValue As String * 8
  RepeatTaxRate As String * 4
  RepeatTotTax As String * 9
  RepeatTaxRelief As String * 9
  RepeatNetTax As String * 9
End Type

