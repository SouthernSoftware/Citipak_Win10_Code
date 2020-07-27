Attribute VB_Name = "BL_STRUCT"
Option Explicit
     Public Const BLData = "BLData\"
     Public Const BLCatCodeName = "ARCODE.DAT"
     Public Const BLCustFileName = "ARCUST.DAT"
     Public Const BLTransFileName = "ARTRANS.DAT"
     Public Const JGLAcctIdxFile = "GLACCT.IDX"
     Public Const AcctFileName = "GLACCT.DAT"
     Public Const CatCodeIdxName = "arcatcodeidx.dat"
     Public Const CustNameIdx = "arcustnameidx.dat"
     Public Const LicNumIdx = "arlicnumidx.dat"
     Public Const CustNumIdx = "arcustnumidx.dat"
     Public Const CustSearchNameIdx = "arsrhidx.dat"
     Public Const BLTransTempPost = "artmppst.dat"
     Public Const BLOperRecName = "CMOPER.DAT"
     Public Const BLPayFileName = "AREDPY"
     Public Const BLTownSetUpName = "artownsu.dat"
     Public Const BLTempCustRecName = "artmpcus.dat"
     Public Const BLTempPrintLicName = "artmplic.dat"
     Public Const BLTempPenaltyCharges = "artmppen.dat"
     Public Const BLLaserLetterName1 = "arlaser1.dat"
     Public Const BLLaserLetterName2 = "arlaser2.dat"
     Public Const BLLaserLetterName3 = "arlaser3.dat"
     Public Const BLLaserLetterName4 = "arlaser4.dat"
     Public Const BLLaserLetterName5 = "arlaser5.dat"
     
Type ARCustRecType
    CustNumb As String * 10
    SortName As String * 10
    BillName As String * 35
    ADDRESS1 As String * 35
    ADDRESS2 As String * 35
    City     As String * 20
    State    As String * 2
    ZipCode  As String * 10
    CustName As String * 35
    Contact  As String * 30
    BILLCAT1     As String * 5
    DESC1        As String * 35
    REV1         As Long
    Fee1         As Double 'latest charge
    FeeLicBal1   As Double
    FeeLicPay1   As Double
    BILLCAT2     As String * 5
    DESC2        As String * 35
    REV2         As Long
    Fee2         As Double
    FeeLicBal2   As Double
    FeeLicPay2   As Double
    BILLCAT3     As String * 5
    DESC3        As String * 35
    REV3         As Long
    Fee3         As Double
    FeeLicBal3   As Double
    FeeLicPay3   As Double
    BILLCAT4     As String * 5
    DESC4        As String * 35
    REV4         As Long
    Fee4         As Double
    FeeLicBal4   As Double
    FeeLicPay4   As Double
    BILLCAT5     As String * 5
    DESC5        As String * 35
    REV5         As Long
    Fee5         As Double
    FeeLicBal5   As Double
    FeeLicPay5   As Double
    IssuanceFee  As Double
    CustLocation As String * 1
    WPHONE       As String * 14
    FeeAmt       As Double
    LICENSE      As String * 12
    VALID        As Integer
    Inactive     As String * 1    '"Y" if account is inactive
    Prorate      As Integer       'prorate percentage
    AcctBal      As Double
    IssueLicense As String * 1    'y/n
    Deleted      As String * 1    '(yY)=deleted, anything else isn't
    FirstTrans   As Long
    LastTrans    As Long
    LicBal       As Double
    FeeBal       As Double
    PenBal       As Double
    RoomtoGrow   As String * 136
    ChkByte      As String * 1
    IssuanceBal  As Double
    IssuancePay  As Double
    ServAdd      As String * 35
    SSNFID       As String * 15
End Type

Type ARNewCatCodeRecType
    CatCode    As String * 5    'Not Used in Version 8.5 work2 directory
    CodeType   As String * 1    ' F=Flat M=Multiplier S=Step
    CODEDESC   As String * 35
    Fee        As Single
    BaseAmt1   As Single
    Recpt1     As Double
    Percent1   As Single
    Maximum1   As Double
    BaseAmt2   As Single
    Recpt2     As Double
    Percent2   As Single
    Maximum2   As Double
    BaseAmt3   As Single
    Recpt3     As Double
    Percent3   As Single
    Maximum3   As Double
    BaseAmt4   As Single
    Recpt4     As Double
    Percent4   As Single
    Maximum4   As Double
    BaseAmt5   As Single
    Recpt5     As Double
    Percent5   As Single
    Maximum5   As Double
    REVGLNUM   As Long
    CASHACCT   As Long
    ARGLACCT   As Long

    BaseAmt6   As Single
    Recpt6     As Double
    Percent6   As Single
    Maximum6   As Double
    RateStep   As Long
    Extra      As String * 36
End Type

Type ARCustIDXRecType
    IDXNAME     As String * 10
    IDXRECORD   As Integer
    ExtraRoom   As String * 52
End Type

Type ARTransRecType
    CustomerNumber      As String * 10
    TransDate           As Integer
    TransAmount         As Double
    TransType           As Integer
    TransDesc           As String * 35 '5
    CashAmount          As Double
    ChkAmount           As Double
    BalanceAfterTrans   As Double
    Posted2GL           As String * 1
    CatCodeRec1         As Long  '10         'Place to Grab G/L Acct #'s
    CatCodeRec2         As Long           'Place to Grab G/L Acct #'s
    CatCodeRec3         As Long           'Place to Grab G/L Acct #'s
    CatCodeRec4         As Long           'Place to Grab G/L Acct #'s
    CatCodeRec5         As Long           'Place to Grab G/L Acct #'s
    CatLicAmt1          As Double '15
    CatLicAmt2          As Double
    CatLicAmt3          As Double
    CatLicAmt4          As Double
    CatLicAmt5          As Double
    CatLicBal1          As Double '25
    CatLicBal2          As Double
    CatLicBal3          As Double
    CatLicBal4          As Double
    CatLicBal5          As Double
    PenBal              As Double
    LicBal              As Double
    IssBal           As Double
    FeeAmt              As Double
    LicAmt              As Double
    PenAmt              As Double
    IssAmt              As Double
    ExtraRoom           As String * 8
    NextTrans           As Long
    DetailTransType     As Integer 'used for reading transaction types inside BL program for reports...not GL
    'Codes for General Ledger:
    '1 = all non-penalty charges; 2 = all payments; 6 = all penalty charges; 13 = adjust payment down
    '23 = adjust billing down; 24 = adjust billing up
    'Codes for internal Business License:
    '101 = Charge Penalty ; 110 = Charge Lic; 201 = Pay Penalty; 210 = Pay Lic; 211 = Pay Lic and Penalty; 301 = Adjust Down Pen; 310 = Adjust Down Lic
    '311 = Adjust Down Pen and Lic; '401 = Adjust Up Pen; 410 = Adjust Up Lic; 411 = Adjust Up Lic and Penalty
End Type


Type AREditPaymentRecType
    TRANTYPE        As Integer
    TranDate        As Integer
    CustNumber      As String * 10
    CustName        As String * 35
    Add1            As String * 35
    City            As String * 25
    State           As String * 2
    ZipCode         As String * 10
    Amount          As Double
    CASHCHK         As String * 9
    CASHAMT         As Double
    CHKAMT          As Double
    CREDITAM        As Double
    AMTPAID         As Double
    CHANGE          As Double
    ISSUELIC        As String * 1
    SetFee          As String * 1
    ISSueFEE        As Double
    DESC            As String * 20
    LICDUE          As Double
    LICDUE1         As Double
    LICDUE2         As Double
    LICDUE3         As Double
    LICDUE4         As Double
    LICDUE5         As Double
    LICPAID         As Double
    LICPAID1        As Double
    LICPAID2        As Double
    LICPAID3        As Double
    LICPAID4        As Double
    LICPAID5        As Double
    TOTDUE          As Double
    TotPaid         As Double
    CatDesc1        As String * 35
    CatDesc2        As String * 35
    CatDesc3        As String * 35
    CatDesc4        As String * 35
    CatDesc5        As String * 35
    PENDUE          As Double
    PENPAID         As Double
    ISSDUE          As Double
    ISSPAID         As Double
End Type

Type CatCodeIdxType
  CatCodeRec As Integer
  CatCodeNum As String * 20
End Type

Type GLSetupRecType                 'still under const.
   UserName    As String * 30
   TotAcctLen  As Integer
   FundLen     As Integer
   AcctLen     As Integer
   DetLen      As Integer
   CASHACCT    As String * 14
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
   POStop      As Boolean    ' this is new 7/22/02 for tabstop on invoice entry
 'Fields added for V205
   PSLFlag     As Integer   '1 for default to Yes, 0 for No
   DupInvFlag  As Integer   '1 to allow duplicates, 0 for No
   CRBank      As Integer   'banknum as default on entry
   CDBank      As Integer   'banknum for default on entry
   ChkBank     As Integer   'banknum for default on check printingEnd Type
End Type

Type CustNameIdxType
   BillingName As String * 35
   CustRec As Integer
End Type

Type CustLicNumIdxType
   LicNum As String * 12
   CustRec As Integer
End Type

Type CustNumIdxType
   CustNumb As String * 10
   CustRec As Integer
End Type

Type CustSearchNameIdxType
   SortName As String * 10
   CustRec As Integer
End Type

Type TransIdxType
  TransWho As String * 35
  TransRecNum As Double
  TransAmt As Double
End Type

Type TempTransPostType
    CustomerNumber      As String * 10
    TransDate           As Integer
    TransAmount         As Double
    TransType           As Integer
    TransDesc           As String * 35
    BalanceAfterTrans   As Double
    Posted2GL           As String * 1
    CatCodeRec1         As Long           'Place to Grab G/L Acct #'s
    CatCodeRec2         As Long           'Place to Grab G/L Acct #'s
    CatCodeRec3         As Long           'Place to Grab G/L Acct #'s
    CatCodeRec4         As Long           'Place to Grab G/L Acct #'s
    CatCodeRec5         As Long           'Place to Grab G/L Acct #'s
    CatFee1             As Double
    CatFee2             As Double
    CatFee3             As Double
    CatFee4             As Double
    CatFee5             As Double
    CatFeeBal1             As Double
    CatFeeBal2             As Double
    CatFeeBal3             As Double
    CatFeeBal4             As Double
    CatFeeBal5             As Double
    LICENSE             As String * 12
    VALID               As Integer
    LicBal              As Double
    AcctBal             As Double
    PenBal              As Double
    Prev                As Long
    CreditUsed          As Boolean
    IssFee              As Double
    IssFeeBal           As Double
End Type

Type CMOperRecType
    OperatorNumber As Integer
    OperatorName As String * 30
    OperatorPassword As String * 8
    NotUsed As String * 24
End Type

Type MiscCodeRecType
    MiscCode As String * 7
    Description As String * 25
    GlAcctNumb As String * 14
    NotUsed As String * 18
End Type

Type RMReceiptRecType
    RecName As String * 25
    RecAddress As String * 25
    RecDesc As String * 25
    RecAmtOwed As Double
    RecPayType As Integer
    RecCashAmt As Double
    RecCheckAmt As Double
    RecChangeDue As Double
    RecDate As String * 10
    RecOperator As String * 2
    RecptNumber As Single
    RecBlank As String * 1
    RecBalance As Single
End Type

Type CMTransRecType
    TransDate    As Integer
    TransAmount  As Double
    TransCash    As Double
    TransCheck   As Double
    TransAmtOwed As Double
    TransDesc    As String * 25
    TransSource  As Integer            '1-Misc 2-Util 3-Tax 4-License 5-decal
    TransName    As String * 25
    TransAcctNum As Long               'Holds Master Acct Record Number in Module
    TransDetNum  As Long               'Holds Record Number of Transaction Detail in Module
    TransRevAmt(1 To 15) As Double
    TransOperNum As Long
    Trans2GL      As String * 1
    TransPad     As String * 25
End Type

Type CMConfigType
    TownName As String * 30
    CASHACCT As String * 14
    LPTPORT  As String * 1
    PrnDefYN As String * 1
    ENDMSG   As String * 30
End Type

Type TownSetUpType
    TownName As String * 38 'allow for TOWN OF
    TownAdd1 As String * 30
    TownAdd2 As String * 30
    Contact As String * 30
    City As String * 30
    State As String * 2
    ZipCode As String * 10
    TownPhone As String * 14
    AppForm As Integer
    DLQNotice As Integer
    SpareSpace As String * 60
    AppAdd1 As String * 30
    AppCity As String * 30
    AppState As String * 2
    AppPhone As String * 14
    AppAdminName As String * 25
    AppAdminTitle As String * 25 '17
    AppBaseFee(1 To 10) As Double
    AppCentsPer(1 To 4) As Double
    AppGrsRcpts(1 To 4) As Double '29
    AppFirstDay As String * 7
    AppLastDay As String * 7
    AppTownOf As String * 38
    AppZip As String * 10
    AppPct As Double
    AppGrsPct As Double
    AppDenom As Integer
    AppNumer As Integer '37
    AppColFee As Double
    AppPayBy As Integer
    AppDiscPct As Double
    AppDiscMonth As String * 9
    AppDiscDay As Integer
    AppPenMonth As String * 9
    AppPenDay As Integer
    AppFiscMonth As String * 9
    AppFiscDay As Integer
    AppMayorCouncil As String * 25
    AppWholeMonth As Integer
    AppWholeDay As Integer '52
    AppRetailMonth As Integer
    AppRetailDay As Integer
    AppFinMonth As Integer
    AppFinDay As Integer
    AppContMonth As Integer
    AppContDay As Integer
    AppRepairMonth As Integer
    AppRepairDay As Integer
    AppStartMonth As String * 9
    AppStartDay As Integer
    AppLicRetMonth As String * 9
    AppLicRetDay As Integer
    AppAdoptDate As Integer
    AppCityOrd As String * 40
    AppYrUpDown(1 To 10) As String * 4
    DlqTownName As String * 38
    DlqAdd1 As String * 30 '68
    DlqCity As String * 30
    DlqState As String * 2
    DlqZip As String * 10
    DlqPhone As String * 14
    DlqPhone2 As String * 14
    DlqFax As String * 14
    DlqAdminName As String * 25
    DlqAdminTitle As String * 25
    DlqFirstDay As String * 9
    DlqLastDay As String * 9
    DlqFirstHour As String * 9
    DlqLastHour As String * 9
    DlqClerkName As String * 25
    DlqMayorCouncil As String * 25 '82
    LicNumPermYN  As String * 3
    UseAmtPctYN   As String * 3
    PENREVGLNUM   As Long
    PENRECGLNUM   As Long
    PENCASHACCT   As Long
    IssFee        As Double
    AcctMeth      As String * 1
    LaserLtr      As String * 1
    GL2Cats       As String * 1
End Type

Type TempCustRecType
    CustRecNum As Integer
    AppType As Integer
    ThisYear As String * 4
    Fee(1 To 5) As Double
    CatCode(1 To 5) As String * 5
    CatDesc(1 To 5) As String * 35
    MiscNum As Double
    AmtPct As String * 3
    IssFee As Double
End Type

Type TempChargesType
    CustNumb As String * 10
    SortName As String * 10
    BillName As String * 35
    CustName As String * 35
    CustRecNum As Integer
    BILLCAT1     As String * 5
    DESC1        As String * 35
    REV1         As Double
    Fee1         As Double
    BILLCAT2     As String * 5
    DESC2        As String * 35
    REV2         As Double
    Fee2         As Double
    BILLCAT3     As String * 5
    DESC3        As String * 35
    REV3         As Double
    Fee3         As Double
    BILLCAT4     As String * 5
    DESC4        As String * 35
    REV4         As Double
    Fee4         As Double
    BILLCAT5     As String * 5
    DESC5        As String * 35
    REV5         As Double
    Fee5         As Double

'************
    FeeAmt       As Double
    Prorate      As Integer       'prorate percentage
    AcctBal      As Double

    LicBal       As Double
    FeeBal       As Double
End Type

Type TempLicPrintType
    LicNum       As Double
    RecNum       As Integer
    Head1        As String * 30
    Head2        As String * 30
    Head3        As String * 30
    Head4        As String * 30
    Issue        As String * 8
    Expire       As String * 8
    ThisYear     As String * 4
    SeqNum       As Integer
    FeeYN        As Boolean
    TBalYN       As Boolean
    Order        As String * 1
End Type

Type TempPenaltyCharges
    CustomerNumber      As String * 10
    TransDate           As Integer
    TransAmount         As Double
    PenAmt              As Double
End Type

Type JGLAcctIdxType
  AcctNum As String * 14
  RecNo   As Integer
End Type

'Type GLAcctRecType                  'Account Record Type: ? bytes
'   Deleted     As Integer           'Active Account Flag
'   Num         As String * 14       'Formatted Account Number
'   Title       As String * 30       'Account Description
'   Typ         As String * 1        'Account Type
'   FrstTran    As Long              'Pointer to First Trans
'   LastTran    As Long              'Pointer to Last Trans
'   PYAct       As Double            'Prior Year Actual
'   BegBal      As Double            'Beginning Balance
'   Bgt         As Double            'Budget Amount
'   Bal         As Double            'Running Balance
'   Encumb      As Double            'Encumbered Amount
'   MTD         As Double            'Month to Date Bal (calc as needed)
'   YTD         As Double            'Year to Date Bal (calc as needed)
'   NYEst       As Double            'Bgt Estimate
'   NYReq       As Double            'New Year Bgt Requested
'   NYRec       As Double            'New Year Bgt Recommended
'   NYApp       As Double            'New Year Bgt Approved
'   FrstBTran   As Integer           'Pointer to First Budget Trans
'   LastBTran   As Integer           'Pointer to Last Budget Trans
'   FrstPTran   As Integer           'Pointer to First Budget Trans
'   LastPTran   As Integer           'Pointer to Last Budget Trans
'   'Res         AS STRING * 25       'Reserved for future needs
'   Work        As Double            'Temp added 08/17/96 for closeout
'   Res         As String * 17
'   Marked      As Integer           '
'End Type

Type LaserLetterType1
  Header         As String * 50
  TownOf         As String * 38
  Address        As String * 30
  CityStateZip   As String * 44
  Phone          As String * 14
  Line1(0 To 11) As String * 80
End Type

Type LaserLetterType2
  TownOf         As String * 38
  Address        As String * 30
  CityStateZip   As String * 44
  Phone          As String * 14
  Line1(0 To 7) As String * 111
End Type

Type LaserLetterType3
  Line1(0 To 5) As String * 50
  Line2(0 To 3) As String * 80
  Signer        As String * 35
  Phone          As String * 14
End Type

Type LaserLetterType4 'Delinquent Notice
  Line1(0 To 3) As String * 50
  Line2(0 To 7) As String * 111
  Signer        As String * 35
  Phone          As String * 14
End Type

Type LaserLetterType5 'Application
  Header As String * 50
  Line1(0 To 13) As String * 111
  PrdBeg As Integer
  PrdEnd As Integer
  BusType(1 To 10) As String * 50
  TaxPer(1 To 10) As String * 30
  BLFee As String * 10
  OptFeeDesc As String * 30
  OptFee As String
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

Type CitiPassTempType
  usernum   As Integer
  UserName  As String * 15
  frommdl   As Integer   'this is to indicate to citipak ok to have file
End Type


