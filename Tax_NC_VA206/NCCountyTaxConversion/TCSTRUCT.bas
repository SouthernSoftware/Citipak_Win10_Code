Attribute VB_Name = "TCSTRUCT"
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
   Blank          As String * 117
   InternalPin    As Long
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
  MHVALUE           As Double
  MCVALUE           As Double
  CVALUE            As Double
  MTVALUE           As Double
  PDESC1            As String * 30
  PDESC2            As String * 30
  PDESC3            As String * 30
  REXMPSENI         As Double
  REXMPOTHR         As Double
  PROPVALU          As Double
  PropSize          As Double
  LOTNUMB           As String * 6
  LOTACRE           As String * 1
  RealAdd           As String * 30
  Map               As String * 6
  BLOCK             As String * 6
  RDESC1            As String * 30
  RDESC2            As String * 30
  RDESC3            As String * 30
  CSSN              As String * 11
  OSSN              As String * 11
  OptSrchDesc       As String * 15
  SName             As String * 10
  OptRev1Chrg       As Integer
  OptRev2Chrg       As Integer
  OptRev3Chrg       As Integer
  County4BillName   As String * 20
  RealOptSearch     As String * 20
  LateList          As String * 1
  Cycle             As Long
  CycleName         As String * 20
  RTownShip         As String * 25
  CTownShip         As String * 25
  MORTCODE          As String * 8
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
  RPinNum           As String * 20
  PPinNum           As String * 20
  PEXMPSENI         As Double
  PEXMPOTHR         As Double
  PersVal           As Double
  MHVALUE           As Double
  MCVALUE           As Double
  CVALUE            As Double
  MTVALUE           As Double
  REXMPSENI         As Double
  REXMPOTHR         As Double
  PROPVALU          As Double
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
End Type

Type CitiPassTempType
  usernum   As Integer
  UserName  As String * 15
  frommdl   As Integer   'this is to indicate to citipak ok to have file
End Type



