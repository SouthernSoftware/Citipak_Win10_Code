Attribute VB_Name = "DC_STRUCT"
Option Explicit

Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)

Public Const DCCustFile = "DCCust.dat"
Public Const DCTranFile = "DCTrans.dat"
Public Const DCSetupFile = "DCSetup.dat"
Public Const DCVCodeFile = "DCCODE.dat"
Public Const DCVehFile = "DCVEH.dat"
Public Const DCData = "DCdata\"

Type DCSetupType
    DCTNNAME     As String * 30
    GLInterface  As String * 1
    AppType      As Integer
    DCVers       As String * 3
    Taxbalchk    As String * 1
    DefLook      As String * 1
    Filler       As String * 90  '128
End Type
  
Type struct
  who As String * 14
  RecNum As Integer
End Type

Type DCExpireDate
    ExpireDate As String * 10
End Type

Type DCCustRecType
    CUSTNUMB     As String * 10
    SORTNAME     As String * 10
    BILLNAME     As String * 35
    ADDRESS1     As String * 35
    ADDRESS2     As String * 35
    City         As String * 20
    State        As String * 2
    ZIPCODE      As String * 10
    SOSEC        As String * 11
    DRVLIC       As String * 12
    DATEOPED     As Integer
    CASHONLY     As String * 1
    resident     As String * 1
    Owner        As String * 1
    HPHONE       As String * 14
    WPHONE       As String * 14
    LICENSE      As String * 12
    Valid        As Integer
    AcctBal      As Double
    Deleted      As String * 1      'rem y=deleted :AnyThing Else is Non-Delet
    FirstTrans   As Long
    LastTrans    As Long
    FirstCar     As Long
    LastCar      As Long
    SocSec1      As String * 11
    OtherName    As String * 25
    RoomtoGrow   As String * 224
End Type

Type DCCatCodeRecType
    CATCODE    As String * 3
    CODEDESC   As String * 35
    APPNUMB    As Integer
    BILLCODE   As Integer
    REVGLNUM   As String * 14
    CashAcct   As String * 14
    Fee        As Single
    InactiveFlag As String * 1
    Extra      As String * 53
End Type

Type DCCustIDXRecType
    IDXName     As String * 10
    IDXRECORD   As Long
    ExtraRoom   As String * 52
End Type

Type ZipIndexType
    IDXName     As String * 10
    IDXRECORD   As Long
End Type

Type DCTempIDXRecType
    IDXRECORD   As Long
End Type

Type DCTransRecType
    CustomerNumber As String * 10
    TransDate As Integer
    TransAmount As Double
    TransType As Integer     '1-charge,2-pay,3-voidchrg,4-voidpay
    TRVinDesc As String * 40
    CashAmount As Double
    ChkAmount As Double
    BalanceAfterTrans As Double
    makemodel As String * 25
    StateTag As String * 35
    ExpireDate As Integer
    Sticker As String * 12
    NextTrans As Long
    OperNum   As Long
    GLInterfaced  As String * 1
    DecalCat As String * 5
    TransTender  As Integer     'Type: 1=Cash, 2=Check, 3=Cash/Check, 4=Charge
''added tendertype and 3,4 transtypes above and  chkbyte to prevent reconverting PS 7/8/05
    VoidFlag As String * 1   'Y if voided
    ChkByte  As String * 1   'this is chr$(1)
    ExtraDesc As String * 26   'added extra
    VehRecord As Long
    ExtraRoom As String * 48
End Type

Type DCEditPaymentRecType
  VehRecord As Long
  CustNumber As String * 10
  CustName As String * 35
  CustAddr As String * 35
  TranDate As Integer
  Amount As Double
  DecalCat As String * 5
  Sticker As String * 12
  VinDesc As String * 40
  ExpDate As Integer
  makemodel As String * 25
  StateTag As String * 35
  resident As String * 1
  Owner As String * 1
  PersBuss As String * 1
  PayDesc As String * 23
  CashAmt   As Double
  CheckAmt  As Double
  Change    As Double
  OperNum   As Long
  TransTender  As Integer     'Type: 1=Cash, 2=Check, 3=Cash/Check, 4=Charge
  VoidFlag As String * 1   'Y if voided
  Notes  As String * 39
  NewVeh As String * 1
End Type

Type DCVehType
  DecalCat As String * 5
  makemodel As String * 25
  StateTag As String * 35
  ExpireDate As Integer
  Sticker As String * 12
  Valid As String * 1         'y/n means is it current
  Active As String * 1        'y/n  n=deleted record
  Notes As String * 39
  PBFlag  As String * 1
  Desc As String * 40
  Fee As Single
  MasterRecord As Long
  NextRec As Long
  MoreRoom As String * 83
End Type

Type CMBankAcctRecType
    COMPACCT As String * 20
End Type

'This is for temporary Tax bill for Spruce Pine and Fairmont
Type ApplicationDefaultsType
  Head1    As String * 40
  Head2    As String * 40
  Head3    As String * 40
  Head4    As String * 40
  Head5    As String * 40
  Body(1 To 20) As String * 75
  dologo   As Integer  '0 for no 1 for yes
End Type

Type NoticeInfoType
  NoticeDate    As Integer         '1
  msgline       As String * 20        '2
  PrnCategory   As Integer         '6
  PRNORDER      As Integer         '7
  Printlogo     As Integer         '8
  PrnCnt        As Long
End Type

'Type GLSetupRecType                 'V205 added new fields noted below
'   UserName    As String * 30
'   TotAcctLen  As Integer
'   FundLen     As Integer
'   AcctLen     As Integer
'   DetLen      As Integer
'   CashAcct    As String * 14
'   APAcct      As String * 14
'   EncAcct     As String * 14
'   FBAcct      As String * 14
'   FYBeg       As Integer
'   FYEnd       As Integer
'   NYBeg       As Integer
'   NYEnd       As Integer
'   CDCash      As String * 14  'new
'   CDDue       As String * 14
'   CDActive    As String * 1
'   CRCashAcct  As String * 14
'   DeptCode    As String * 1
'   LPDate      As Integer
'   HPDate      As Integer
'   CDCashAcct  As String * 14
'   CDsbCash    As String * 14
'   APChkCode   As Integer
'   POStop      As Boolean          'new 7/22/02 for potab on invoice entry
' 'Fields added for V205
'   PSLFlag     As Integer   '1 for default to Yes, 0 for No
'   DupInvFlag  As Integer   '1 to allow duplicates, 0 for No
'   CRBank      As Integer   'banknum as default on entry
'   CDBank      As Integer   'banknum for default on entry
'   ChkBank     As Integer   'banknum for default on check printing
'   pad         As String * 20
'   ChkVer      As String * 4   ' for "V205"
'End Type

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
'edit the res added function rec pointer 6/11/04
   FNCTRec     As Long
   Res         As String * 12
   ChkByte     As String * 1    'this is updated at GASB34 conversion with chr$(1)
   Marked      As Integer           '
End Type

Type DistArrayType
   DistOrder As Integer
   DistCnt   As Integer
End Type
Type PayListType
  VehRec    As Long
  Listrec   As Long
End Type

Type CMTransRecType
    TransDate    As Integer
    TransAmount  As Double
    TransCash    As Double
    TransCheck   As Double
    TransAmtOwed As Double
    TransDesc    As String * 25
    TransSource  As Integer           '1-Misc 24-Util 27-UtilDep 31-Tax 131-Newtax 41-License 141-NewBL 51-decal
    ''''''''''''''''''''''''''''''''''201-void Misc 224-void util 227-void dep 241-void lic 231-void tax
    ''''''''''''''''''''''''''''''''''251-void Decal
    TransName    As String * 25
    TransAcctNum As Long               'Holds Master Acct Record Number in Mod
    TransDetNum  As Long               'Holds Record Number of Transaction Det
    TransRevAmt(1 To 15) As Double
    TransOperNum As Long
    Trans2GL     As String * 1
    TransTender  As Integer     'Type: 1=Cash, 2=Check, 3=Cash/Check, 4=Charge
'added charge 4 above and transvoid for new void payment procedure PS 4/14/04
    TransVoidNum As Long        'Voided trans link to record voided or void trans
    ChkByte      As String * 1
    TransPad     As String * 18
End Type

Type VATaxCustType
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

Type CustNameIdxType
   CustName As String * 50
   CustRec As Long
End Type



