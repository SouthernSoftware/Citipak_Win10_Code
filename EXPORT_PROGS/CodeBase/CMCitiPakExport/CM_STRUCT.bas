Attribute VB_Name = "CM_STRUCT"
Option Explicit


Public Const CMUBSysFile = "UBSETUP.dat"
Public Const CMTranFile = "CMTrans.dat"
Public Const CMCodeFile = "CMMISCCD.DAT"

Type MiscCodeRecType
    MiscCode As String * 7
    Description As String * 25
    GlAcctNumb As String * 14
    InactiveFlag As String * 1
    NotUsed As String * 17
End Type

Type CMTransRecTypeII
    TransDate    As Integer
    TransAmount  As Double
    TransCash    As Double
    TransCheck   As Double
    TransAmtOwed As Double
    TransDesc    As String * 25
    TransSource  As Integer
    '1-Misc 24-Util 27-UtilDep 31-Tax 131-Newtax 41-License 141-NewBL 51-decal
    '''''''''''''''''''''''''''''''''
    '201-void Misc 224-void util 227-void dep 241-void lic 231-void tax
    '''''''''''''''''''''''''''''''''
    '251-void Decal
    TransName    As String * 25
    TransAcctNum As Long               'Holds Master Acct Record Number in Mod
    TransDetNum  As Long               'Holds Record Number of Transaction Det
    TransRevAmt(1 To 15) As Double     'if tr=1 then value is misc code pointer
    TransOperNum As Long
    Trans2GL     As String * 1
    TransTender  As Integer     'Type: 1=Cash, 2=Check, 3=Cash/Check, 4=Charge
'added charge 4 above and transvoid for new void payment procedure PS 4/14/04
    TransVoidNum As Long        'Voided trans link to record voided or void trans
    ChkByte      As String * 1
    TransPad     As String * 18
End Type

Type SetUpAcctType
   RevName    As String * 15
   DebitAcct  As String * 14
   CreditAcct As String * 14
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

Type UBSetupRecType
    UTILNAME        As String * 35
    DEFCITY         As String * 18
    DEFSTATE        As String * 2
    ZIPCODE         As String * 10
    PreByBook       As String * 1
    'RecpPort        As String * 1 ' change to LockBoxDef on 1/25/05
    LockBoxDef      As String * 1  ' 6 for 6digit acct, 8 for 8digit acct file type struc
    RECPDEFT        As String * 1
    ESTREAD         As String * 1
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
    HHDEVICE        As String * 1    'P=PC3000 S=Sensus C=Syscom R=Radix N=None
    Revenues(1 To 15) As RevSetUpType
    BillAcct(1 To 15) As SetUpAcctType
    PayAcct(1 To 15)  As SetUpAcctType
    DepAcct(1 To 15)  As SetUpAcctType
End Type

