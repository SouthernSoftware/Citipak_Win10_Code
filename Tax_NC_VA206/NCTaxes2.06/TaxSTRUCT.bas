Attribute VB_Name = "TaxSTRUCT"
Option Explicit
Public Const TranUtilityBill = 1          '   1=Utility bill
Public Const TranLateCharge = 2           '   2=late charge
Public Const TranReconnectFee = 3         '   3=reconnect fee
Public Const TranBillPayment = 4          '   4=Bill Payment
Public Const TranAppliedDeposit = 5       '   5=Applied Deposit
Public Const TranPenaltyCharge = 6        '   6=Penalty Charge
Public Const TranDepositPayment = 7       '   7=Deposit Payment
Public Const TranDraftPayment = 8         '   8=Draft Payment
Public Const TranRefundDeposit = 9        '   9=Refund Deposit
Public Const TranBeginBalance = 10        '  10=Beginning Balance
Public Const TranUpwardAdjustment = 11    '  11=Bill Adjustments
Public Const TranDownwardAdjustment = 12  '  12=Bill Adjustments
Public Const TranMiscPayment = 99         '  99=Misc Payment
Public Const MaxRevsCnt = 15              '  Max num of Utility Revenues

Type JGLAcctIdxType
  AcctNum As String * 14
  RecNo   As Integer
End Type

Type GLSetupRecType
   UserName    As String * 30
   TotAcctLen  As Integer
   FundLen     As Integer
   AcctLen     As Integer
   DetLen      As Integer
   CashAcct    As String * 14
   APAcct      As String * 14
   EncAcct     As String * 14
   FBAcct      As String * 14
   FYBeg       As Integer
   FYEnd       As Integer
   NYBeg       As Integer
   NYEnd       As Integer
   CDCash      As String * 14  'new
   CDDue       As String * 14
   CDActive    As String * 1
   CRCashAcct  As String * 14
   DeptCode    As String * 1
   LPDate      As Integer
   HPDate      As Integer
   CDCashAcct  As String * 14
   CDsbCash    As String * 14
   APChkCode   As Integer
   POStop      As Boolean          'new 7/22/02 for potab on invoice entry
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
  ClassName(1 To 6) As String * 15
  OptSrchPers As String * 15 'added 8/16/06
  Padding As String * 98 ' 99 changed to 98 on 7/25/08
  AutoFillSrvAdd As String * 1 '7/25/08
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
  XFileNme As String * 8
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

Type FLen2
  V As String * 64
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
    ICPDesc As String * 15
    Blank  As String * 55
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
   InternalPin    As Long
   OptSearch      As String * 20 'added 8/16/06
   Blank          As String * 97
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
     MORTCODE           As String * 8
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
     SetDscvry2No       As String * 1 'added 12/5/06
     Padding            As String * 104
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
    XDate    As Integer 'added 9/20/05
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
    PayDate  As Integer
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
Type PayListType
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

Type GLAcctIndexType                'Account Index: 16 bytes
   AcctNum     As String * 14       'Formatted account Number string
   RecNum      As Integer           'Pointer to record
   '*****
End Type

Type GLAcctRecType                  'Account Record Type: ? bytes
   Deleted     As Integer           'Active Account Flag
   Num         As String * 14       'Formatted Account Number
   Title       As String * 30       'Account Description
   Typ         As String * 1        'Account Type
   FrstTran    As Long              'Pointer to First Trans
   LastTran    As Long              'Pointer to Last Trans
   PYAct       As Double            'Prior Year Actual
   BegBal      As Double            'Beginning Balance
   Bgt         As Double            'Budget Amount
   Bal         As Double            'Running Balance
   Encumb      As Double            'Encumbered Amount
   MTD         As Double            'Month to Date Bal (calc as needed)
   YTD         As Double            'Year to Date Bal (calc as needed)
   NYEst       As Double            'Bgt Estimate
   NYReq       As Double            'New Year Bgt Requested
   NYRec       As Double            'New Year Bgt Recommended
   NYApp       As Double            'New Year Bgt Approved
   FrstBTran   As Integer           'Pointer to First Budget Trans
   LastBTran   As Integer           'Pointer to Last Budget Trans
   FrstPTran   As Integer           'Pointer to First Budget Trans
   LastPTran   As Integer           'Pointer to Last Budget Trans
   'Res         AS STRING * 25       'Reserved for future needs
   Work        As Double            'Temp added 08/17/96 for closeout
   Res         As String * 17
   Marked      As Integer           '
End Type

Type GLDeptIndexType                'Dept Index
   DeptNum     As String * 8        'Dept Number
   RecNum      As Integer           'Pointer to record
   '*****
End Type

Type GLDeptRecType                  'Dept Record Type
   Deleted     As Integer           'Deleted Flag
   DeptNum     As String * 8        'Fund Code
   Title       As String * 30       'Fund Title
   Res         As String * 20       'Reserve for future needs
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


Type TranRecInfoType
    TranDate  As Integer
    TranRecNo As Long
End Type

Type MiscCodeRecType
    MiscCode As String * 7
    Description As String * 25
    GlAcctNumb As String * 14
    NotUsed As String * 18
End Type

Type CMTransRecType
    TransDate    As Integer
    TransAmount  As Double
    TransCash    As Double
    TransCheck   As Double
    TransAmtOwed As Double
    TransDesc    As String * 25
    TransSource  As Integer            '1-Misc 2-Utility 3-Tax 4-License
                                       '5-decal
    TransName    As String * 25
    TransAcctNum As Long               'Holds Master Acct Record Number in Module
    TransDetNum  As Long               'Holds Record Number of Transaction Detail in Module
    TransRevAmt(1 To 15) As Double
    TransOperNum As Long
    Trans2GL      As String * 1
    TransPad     As String * 25
End Type

Type RevSetUpType
    RevName As String * 15
    UseDep   As String * 1
    USERATE  As String * 1
    TAXRATE  As Single
    UseMtr   As String * 1
    DistOr   As Integer
    ProRate  As String * 1
End Type

Type SetUpAcctType
   RevName    As String * 15
   DebitAcct  As String * 14
   CreditAcct As String * 14
End Type

Type UBSetupRecType
    UTILNAME        As String * 35
    DEFCITY         As String * 18
    DEFSTATE        As String * 2
    ZIPCODE         As String * 10
    PreByBook       As String * 1
    RecpPort        As String * 1
    RECPDEFT        As String * 1
    EstRead         As String * 1
    BANKDFT         As String * 1
    UseSeq          As String * 1
    BILLCYCL        As String * 1
    DefLook         As String * 1
    MethAcct        As String * 1      'new 02-14-97
    SkipInactive    As String * 1
    SkipSeparator   As String * 1
    Make99File      As String * 1
    LowRead         As Integer
    HighRead        As Integer
    HHDEVICE        As String * 1    'P=PC3000 S=Sensus C=Syscom R=Radix N=Non
    Revenues(1 To 15) As RevSetUpType
    BillAcct(1 To 15) As SetUpAcctType
    PayAcct(1 To 15)  As SetUpAcctType
    DepAcct(1 To 15)  As SetUpAcctType
End Type
'Note:  if transaction is an adjustment then
'       CurRead field will contain the adjust amount

Type UBTransRecType
   TransDate              As Integer      '
   TransType              As Integer      '
   TransDesc              As String * 21  'may change
   TransAmt               As Double       'total revenue amount
   RevAmt(1 To 15)        As Double       'Revenue amounts
   TaxAmt(1 To 15)        As Single       'Tax Amounts
'01-20-97 Added meter types field to hold meter type at time of transaction
   MtrTypes(1 To 7)       As Integer
'*******************
   CurRead(1 To 7)        As Long         'Last/Current meter readings
   PrevRead(1 To 7)       As Long         'Previous readings
   EstRead(1 To 7)        As String * 1   'Y/N Flags for meter est's
   BillNumber             As Long         'Number on the bill that Printed
   ReadDate               As Integer
   BillDate               As Integer
   PastDueDate            As Integer
   DraftDate              As Integer      'mowasa & plymouths bills. Can be ch
'111398
   ProratePCT             As Integer
   Filler1                As String * 2
   'CustLocation          AS LONG         'Pointer to Location RecNo
   CustStatus             As String * 1   'Customer Status at Time of Transact
'102998
   PenAtBill              As Double       'calculated penalty at time of bill
   PayTypeCode            As Integer      'Payment Type:  1=Cash, 2=Check
   OperatorNumber         As Integer      '
   CustAcctNo             As Long         'Pointer to RecNo in ubcust.dat
   PrevTrans              As Long
   ActUsage               As Long         'Changed for wadesboro
   ActiveFlag             As Integer      'Valid transaction flag
   RunBalance             As Double
   CheckAmount            As Double
   CashAmount             As Double
   BillMsg                As String * 20
   ApplyDepFlag           As String * 1
   Posted2GL              As String * 1
   PrevDate               As Integer
   PenalFlag              As String * 1
   TaxExempt              As String * 1
   NONProfit              As String * 1
End Type

'Trans Types
Type UBXferInfoType
  DAcctNo   As String * 14
  DebitAmt  As Double
  DRecNo    As Integer       '**** Don't know if conversion is needed
  DTitle    As String * 30
  CAcctNo   As String * 14
  CreditAmt As Double
  CRecNo    As Integer       '**** Don't know if conversion is needed
  CTitle    As String * 30
End Type

Type GJXferRecType
    RevText    As String * 15
    BAcctInfo  As UBXferInfoType     'Billing Accounts
    PAcctInfo  As UBXferInfoType     'Payment Accounts
    DAcctInfo  As UBXferInfoType     'Deposit Accounts
End Type

'Type ARCatCodeRecType
'    CATCODE    As String * 5    'Not Used in Version 8.5 work2 directory
'    CODEDESC   As String * 35
'    Fee     As Single
'    REVGLNUM   As String * 14
'    CashAcct   As String * 14
'    ARGLACCT   As String * 14
'    CodeType   As String * 1    ' F=Flat M=Multiplier S=Step
'    Percent    As Single
'    Maximum    As Double
'    Extra      As String * 157
'End Type
'
'Type ARNewCatCodeRecType
'    CATCODE    As String * 5    'Not Used in Version 8.5 work2 directory
'    CodeType   As String * 1    ' F=Flat M=Multiplier S=Step
'    CODEDESC   As String * 35
'    Fee        As Single
'    BaseAmt1   As Single
'    Recpt1     As Double
'    Percent1   As Single
'    Maximum1   As Double
'    BaseAmt2   As Single
'    Recpt2     As Double
'    Percent2   As Single
'    Maximum2   As Double
'    BaseAmt3   As Single
'    Recpt3     As Double
'    Percent3   As Single
'    Maximum3   As Double
'    BaseAmt4   As Single
'    Recpt4     As Double
'    Percent4   As Single
'    Maximum4   As Double
'    BaseAmt5   As Single
'    Recpt5     As Double
'    Percent5   As Single
'    Maximum5   As Double
'    REVGLNUM   As String * 14
'    CashAcct   As String * 14
'    ARGLACCT   As String * 14
'    Extra      As String * 64
'End Type
'
'Type ARTransRecType
'    CustomerNumber      As String * 10
'    TransDate           As Integer
'    TransAmount         As Double
'    TransType           As Integer
'    TransDesc           As String * 35
'    CashAmount          As Double
'    ChkAmount           As Double
'    BalanceAfterTrans   As Double
'    Posted2GL           As String * 1
'    CatCodeRec          As Integer           'Place to Grab G/L Acct #'s
'    ExtraRoom           As String * 40
'    NextTrans           As Long
'End Type

Type CitiPassTempType
  usernum   As Integer
  UserName  As String * 15
  frommdl   As Integer   'this is to indicate to citipak ok to have file
End Type

Type CustNameIdxType
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

Type ReceiptPRNType
'This is for new local receipt setup file stored on each computer on
'drive c:\
  RcpPort   As String * 40
  PrnDefYN  As Integer
  CtlDefYN  As Integer
  PaymDate  As Integer    'For Changing Default Date During Daily Entry
  RValidate As Integer
  ZExtra    As String * 16
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
  PayDate As Integer
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
  NegYN As String * 1
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
    UseBarCode As Boolean
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
  LateListAmt As String * 10
  RealPin As String * 20
  ExemptAmt As String * 10 'added at the request of Beech Mt on 9/1/2010
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

Type OptPersIdxType
  OptDesc As String * 20
  PersRec As Long
  PersPin As String * 20
End Type

Type BillPrintZipIdxType
  TaxBillRec As Long
End Type

Type BillPrintMortIdxType
  TaxBillRec As Long
End Type

Type TaxBillTypeOld
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
     MORTCODE           As String * 8
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

Type AllowCountyEdit
  AllowCountyEditXDate As Integer '7/11/07
End Type


