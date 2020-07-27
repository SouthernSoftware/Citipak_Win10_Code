Attribute VB_Name = "TCSTRUCT"
Option Explicit

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
  Padding As String * 99
End Type

Type TaxInterestDateType
  InterestDate As Integer
End Type

Type TaxValuesType
  Value    As Double
  OthVal   As Double
  ExmVal   As Double
End Type

Type TaxCustType
  Acct       As Long
  OPENDATE   As Integer
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
    ICPDesc As String * 15
    BLDGVAL As Double
    Blank  As String * 100
End Type

Type PersonalRecType
   PropPin  As String * 20
   PROPDATE As Integer
   PersVal  As Double
   MHValue  As Double
   MCValue  As Double
   CVALUE   As Double
   MTValue  As Double
   EXMPSENI As Double
   EXMPOTHR As Double
   DISCOV   As String * 1
   LateList As String * 1
   DESC1    As String * 30
   DESC2    As String * 30
   DESC3    As String * 30
   Desc4    As String * 30
   Desc5    As String * 30
   CustPin        As Long
   NextRec        As Long
   LastYrPrinted  As Integer
   Deleted        As Integer
   VehTaxYear     As Integer
   DMVSubmitted   As String * 1
   InternalPin    As Long
   TaxBillYear    As Integer
   PPTRAYN        As String * 1
   Prorate        As String * 1
   ProrateVal     As Integer
   Vin            As String * 25
   MakeMod        As String * 25
   Weight         As Double
   ModYear        As Integer
   OptRev1Chrg    As Integer
   OptRev2Chrg    As Integer
   OptRev3Chrg    As Integer
   Blank          As String * 85
   OptSearch      As String * 20 'added 8/16/06
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

Type ConversionDataType
  CustName          As String * 50
  CountyAcctString  As String * 20    'County Account in String Format when lo
  CountyAcct        As Double       'County Account Number to Link to County Record
  Addr1             As String * 35
  Addr2             As String * 35
  City              As String * 20
  State             As String * 2
  Zip               As String * 10
  RPinNum            As String * 20
  PPinNum            As String * 20
  PEXMPSENI         As Double
  PEXMPOTHR         As Double
  PersVal           As Double
  MHValue           As Double
  MCValue           As Double
  CVALUE            As Double
  MTValue           As Double
  PDESC1            As String * 30
  PDESC2            As String * 30
  PDESC3            As String * 30
  PDESC4            As String * 30
  PDESC5            As String * 30
  REXMPSENI         As Double
  REXMPOTHR         As Double
  PROPVALU          As Double
  PropSize          As Double
  LOTNUMB           As String * 6
  LOTACRE           As String * 1
  RealAdd           As String * 30
  BLDGVAL           As Double
  PPTRAYN           As String * 1
  Vin               As String * 25
  MakeMod           As String * 25
  Weight            As Double
  ModYear           As Integer
  Map               As String * 6
  BLOCK             As String * 6
  GISPOS            As String * 20
  RDESC1            As String * 30
  RDESC2            As String * 30
  RDESC3            As String * 30
  CSSN              As String * 11
  OSSN              As String * 11
  OptSrchDesc       As String * 15
  SName             As String * 10
  HPHONE            As String * 14
  WPHONE            As String * 14
  ROptRev1Chrg       As Integer
  ROptRev2Chrg       As Integer
  ROptRev3Chrg       As Integer
  POptRev1Chrg       As Integer
  POptRev2Chrg       As Integer
  POptRev3Chrg       As Integer
  RealOptSearch      As String * 20
  RTownShip           As String * 25
  CTownShip           As String * 25
  RLateList          As String * 1
  Interest           As String * 1
  TaxExempt          As String * 1
  Penalty            As String * 1
  Bankrupt           As String * 1
  LateNotice         As String * 1
  ServiceAdd         As String * 35
  DrvrsLic           As String * 10
  DeliveryPt         As String * 2
  PostalRt           As String * 4
  Cycle              As Long
  CycleName          As String * 20
  County4BillNum     As Long
  County4BillName    As String * 20
  Prorate            As String * 1
  ProrateVal         As Integer
  MORTCODE           As String * 8
  LienYN             As String * 1
  LienDesc           As String * 30
  ICPDesc            As String * 15
  CustLateNotice     As String * 1
  PersOptSearch      As String * 20
  PLateList          As String * 1
End Type

Type TempConversionData
  CData As ConversionDataType
End Type

Type ConvSpreadsheet
  Field1 As String * 30
  Field2 As String * 30
  Field3 As String * 30
End Type

Type ConvResultsType
  CustName          As String * 50
  CountyAcctString  As String * 20    'County Account in String Format when lo
  CountyAcct        As Double       'County Account Number to Link to County Record
  RPinNum            As String * 20
  PPinNum            As String * 20
  PEXMPSENI         As Double
  PEXMPOTHR         As Double
  PersVal           As Double
  MHValue           As Double
  MCValue           As Double
  CVALUE            As Double
  MTValue           As Double
  REXMPSENI         As Double
  REXMPOTHR         As Double
  PROPVALU          As Double
  BLDGVAL           As Double
  Vin               As String * 25
  MakeMod           As String * 25
  Weight            As Double
  ModYear           As Integer
  PPTRAYN           As String * 1
End Type

Type ConvErrorType
  CustName          As String * 50
  CountyAcctString  As String * 20    'County Account in String Format when lo
  CountyAcct        As Double       'County Account Number to Link to County Record
  RPinNum            As String * 20
  PPinNum            As String * 20
  ErrorType         As Integer
  RealTot           As Double
  RealXTot          As Double
  PersTot           As Double
  PersXTot          As Double
  BLDGVAL           As Double
End Type
