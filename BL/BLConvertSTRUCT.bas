Attribute VB_Name = "BLConvertSTRUCT"
Option Explicit
Type DOSARCustRecType2
    CUSTNUMB     As String * 10
    SORTNAME     As String * 10
    BILLNAME     As String * 35
    ADDRESS1     As String * 35
    ADDRESS2     As String * 35
    CITY         As String * 20
    STATE        As String * 2
    ZIPCODE      As String * 10
    CustName     As String * 35
    BILLCAT      As String * 5
    SOSEC        As String * 11
    DRVLIC       As String * 12
    DATEOPED     As Integer
    BILLCMT      As String * 20
    PAYCMT       As String * 20
    CASHONLY     As String * 1
    APPNUMB      As Integer
    BILLFORM     As Integer
    HPHONE       As String * 14
    WPHONE       As String * 14
    FeeAmt       As Double
    LICENSE      As String * 12
    VALID        As Integer
    AcctBal      As Double
    OldFirstTrans   As Integer
    OldLastTrans    As Integer
    Deleted      As String * 1      'rem y=deleted :AnyThing Else is Non-Deleted
    FirstTrans   As Long
    LastTrans    As Long
    IssueLicense As String * 1      ' rem y/n
    IssuanceFee  As Currency        ' Alabama Specific
    CustLocation As String * 1      ' Alabama Specific
    RoomtoGrow   As String * 164
End Type

Type DOSARCustRecType5
    CUSTNUMB     As String * 10
    SORTNAME     As String * 10
    BILLNAME     As String * 35
    ADDRESS1     As String * 35
    ADDRESS2     As String * 35
    CITY         As String * 20
    STATE        As String * 2
    ZIPCODE      As String * 10
    CustName     As String * 35
    Contact      As String * 30

    BILLCAT1     As String * 5
    DESC1        As String * 15
    REV1         As Long
    Fee1         As Double
    BILLCAT2     As String * 5
    DESC2        As String * 15
    REV2         As Long
    Fee2         As Double
    BILLCAT3     As String * 5
    DESC3        As String * 15
    REV3         As Long
    Fee3         As Double
    BILLCAT4     As String * 5
    DESC4        As String * 15
    REV4         As Long
    Fee4         As Double
    BILLCAT5     As String * 5
    DESC5        As String * 15
    REV5         As Long
    Fee5         As Double

    IssuanceFee  As Currency
    CustLocation As String * 1
    WPHONE       As String * 14
    FeeAmt       As Double
    LICENSE      As String * 12
    VALID        As Integer

    AcctBal      As Double
    IssueLicense As String * 1    'rem y/n
    Deleted      As String * 1    'rem y=deleted :AnyThing Else is Non-Deleted
    FirstTrans   As Long
    LastTrans    As Long
    RoomtoGrow   As String * 164
End Type

Type DOSARCustRecType4                 'used in verion 8.5 work2dir
    CUSTNUMB     As String * 10
    SORTNAME     As String * 10
    BILLNAME     As String * 35
    ADDRESS1     As String * 35
    ADDRESS2     As String * 35
    CITY         As String * 20
    STATE        As String * 2
    ZIPCODE      As String * 10
    CustName     As String * 35
    Contact      As String * 30

    BILLCAT1     As String * 5
    DESC1        As String * 15
    REV1         As Long
    Fee1         As Double
    BILLCAT2     As String * 5
    DESC2        As String * 15
    REV2         As Long
    Fee2         As Double
    BILLCAT3     As String * 5
    DESC3        As String * 15
    REV3         As Long
    Fee3         As Double
    BILLCAT4     As String * 5
    DESC4        As String * 15
    REV4         As Long
    Fee4         As Double
    BILLCAT5     As String * 5
    DESC5        As String * 15
    REV5         As Long
    Fee5         As Double

    IssuanceFee  As Currency
    CustLocation As String * 1
    WPHONE       As String * 14
    FeeAmt       As Double
    LICENSE      As String * 12
    VALID        As Integer
    'dodo AS STRING * 7
    AcctBal      As Double
    IssueLicense As String * 1    'rem y/n
    Deleted      As String * 1    'rem y=deleted :AnyThing Else is Non-Deleted
    FirstTrans   As Long
    LastTrans    As Long
    RoomtoGrow   As String * 164
End Type

Type DOSARCustRecType3
    CUSTNUMB     As String * 10
    SORTNAME     As String * 10
    BILLNAME     As String * 35
    ADDRESS1     As String * 35
    ADDRESS2     As String * 35
    CITY         As String * 20
    STATE        As String * 2
    ZIPCODE      As String * 10
    CustName     As String * 35
    BILLCAT      As String * 5
    SOSEC        As String * 11
    DRVLIC       As String * 12
    DATEOPED     As Integer
    BILLCMT      As String * 20
    PAYCMT       As String * 20
    CASHONLY     As String * 1
    APPNUMB      As Integer
    BILLFORM     As Integer
    HPHONE       As String * 14
    WPHONE       As String * 14
    FeeAmt       As Double
    LICENSE      As String * 12
    VALID        As Integer
    AcctBal      As Double
    OldFirstTrans   As Integer
    OldLastTrans    As Integer
    Deleted      As String * 1      'rem y=deleted :AnyThing Else is Non-Deleted
    FirstTrans   As Long
    LastTrans    As Long
    IssueLicense As String * 1      ' rem y/n
    IssuanceFee  As Currency        ' Alabama Specific
    CustLocation As String * 1      ' Alabama Specific
    Contact As String * 30
    RoomtoGrow   As String * 134
End Type

Type DosARCustRecType
    CUSTNUMB As String * 10
    SORTNAME As String * 10
    BILLNAME As String * 35
    ADDRESS1 As String * 35
    ADDRESS2 As String * 35
    CITY     As String * 20
    STATE    As String * 2
    ZIPCODE  As String * 10
    CustName As String * 35
    Contact  As String * 30

    BILLCAT1     As String * 5
    DESC1        As String * 15
    REV1         As Long
    Fee1         As Double
    BILLCAT2     As String * 5
    DESC2        As String * 15
    REV2         As Long
    Fee2         As Double
    BILLCAT3     As String * 5
    DESC3        As String * 15
    REV3         As Long
    Fee3         As Double
    BILLCAT4     As String * 5
    DESC4        As String * 15
    REV4         As Long
    Fee4         As Double
    BILLCAT5     As String * 5
    DESC5        As String * 15
    REV5         As Long
    Fee5         As Double

'************
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
'************
End Type

Type ARCustRecType
    CUSTNUMB As String * 10
    SORTNAME As String * 10
    BILLNAME As String * 35
    ADDRESS1 As String * 35
    ADDRESS2 As String * 35
    CITY     As String * 20
    STATE    As String * 2
    ZIPCODE  As String * 10
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

'************
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
    
'************
End Type


Type DosARNewCatCodeRecType
    CATCODE    As String * 5    'Not Used in Version 8.5 work2 directory
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
    REVGLNUM   As String * 14
    CASHACCT   As String * 14
    ARGLACCT   As String * 14

    BaseAmt6   As Single
    Recpt6     As Double
    Percent6   As Single
    Maximum6   As Double
    RateStep   As Long
    Extra      As String * 36
End Type

Type ARNewCatCodeRecType
    CATCODE    As String * 5    'Not Used in Version 8.5 work2 directory
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


Type AREditBegBalRecType
    CustNumber                  As String * 10
    CustName                    As String * 35
    TranDate                    As Integer
    Amount                      As Double
    ISSUELIC                    As String * 1       'Y/N Answer
    SetFee                      As String * 1       'Y/N Answer
    LicAmt                      As Double
    PenAmt                      As Double
    TRType                      As Integer
    TDesc                       As String * 20
    Extra                       As String * 4
End Type

Type AREditPaymentRecType
    TRANTYPE        As String * 15
    TranDate        As Integer
    CustNumber      As String * 10
    CustName        As String * 35
    ADD1            As String * 35
    CITY            As String * 25
    STATE           As String * 2
    ZIPCODE         As String * 10
    Amount          As Double
    CASHCHK         As String * 9
    CASHAMT         As Double
    CHKAMT          As Double
    AMTPAID         As Double
    CHANGE          As Double
    ISSUELIC        As String * 1
    SetFee          As String * 1
    ISSueFEE        As Double
    DESC            As String * 20
    LICDUE          As Double
    FEEDUE          As Double
    PENDUE          As Double
    LICPAID         As Double
    FEEPAID         As Double
    PENPAID         As Double
    TOTDUE          As Double
    TOTPAID         As Double
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
   CUSTNUMB As String * 10
   CustRec As Integer
End Type

Type CustSearchNameIdxType
   SORTNAME As String * 10
   CustRec As Integer
End Type

Type TransIdxType
  TransWho As String * 35
  TransRecNum As Double
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

Type DosARNewCatCodeRecType3
    CATCODE    As String * 3    'Not Used in Version 8.5 work2 directory
    CODEDESC   As String * 35
    APPNUMB    As Integer       'Not Used
    BILLCODE   As Integer       'Not Used
    REVGLNUM   As String * 14
    CASHACCT   As String * 14
    ALCATCODE  As String * 5    'Alabama Code Specific
    ARGLACCT   As String * 14
    Extra      As String * 39
End Type

Type DosARNewCatCodeRecType2
    CATCODE    As String * 3    'Not Used in Version 8.5 work2 directory
    CODEDESC   As String * 35
    APPNUMB    As Integer       'Not Used
    BILLCODE   As Integer       'Not Used
    REVGLNUM   As String * 14
    CASHACCT   As String * 14
    ALCATCODE  As String * 5    'Alabama Code Specific
    ARGLACCT   As String * 14
    Extra      As String * 39
End Type
