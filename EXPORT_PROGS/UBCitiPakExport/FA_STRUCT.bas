Attribute VB_Name = "FA_STRUCT"
Option Explicit
   
Public Const FASetUpFileName = "FASETUP.DAT"
Public Const FAItemFileName = "FAITEMS.DAT"
Public Const FAAssetCodeName = "FACODES.DAT"
Public Const FADeptCodeName = "FADEPTCD.DAT"
Public Const FAFundCodeName = "FAFUNDCD.DAT"
Public Const FAYearEndName = "FAYEAR.DAT"
Public Const FADeprEditName = "FADPREDT.DAT"
Public Const FADprHistFileName = "FADPRHIST.DAT"
Public Const TempDprFileName = "FATEMPDPR.DAT"
Public Const TempDispDateName = "FATEMPDISPDATE.DAT"
Public Const FAItemFile = "FAITEMS.DAT"
Public Const FACodeFile = "FACODES.DAT"
Public Const FAData = "fadata\"


Type FAItemRecType
    ItemTag  As String * 20
    ISTATUS  As String * 1
    DEPYN    As String * 1
    AQURDATE As Integer
    IDESC1   As String * 30
    IDESC2   As String * 30
    GLAcct   As String * 14
    IDEPT    As Integer
    ASSETCODE As String * 4
    ILIFE    As Double
    ORGCOST  As Double
    DEP2DATE As Double
    CURRVAL  As Double
    CDEPDATE As Integer
    DispDate As Integer
    VENDOR   As String * 30
    SERIALNO As String * 30
    ITEMMFG  As String * 30
    CONTACT  As String * 30
    ITEMLOC  As String * 30
    EOLDATE  As Integer
    VHCLMAKE As String * 20
    VHCLMODL As String * 20
    VHCLVIN As String * 20
    VHCLTAG As String * 10
    VHCLCOLR As String * 10
    WARRXDAT As Integer
    Fill1     As String * 32
    PHONE As String * 14
'    FileVer   As Integer
    FundNum As Integer  'new for Windows
    DisposAmt As Double  'new for Windows
    LastDprRec As Long
    LifeLeft As Integer
    PONum As String * 15
    CheckNum As String * 10
    DsplFlag As Integer '0 = Still In Use, 1 = Tagged for disposal, 2 = Disposed Of
    DsplMethod As String * 10
End Type

Type FAAssetCodeRecType
    ASSETCODE      As String * 4
    AssetStatus    As String * 10
    AssetDesc      As String * 20
End Type

Type FAYearEndType
    LastYear  As String * 4
    CurYear As String * 4
End Type

Type FADepFileType
    AssetRecord As Long
    CurYrDep As Double
    PctFlag  As Integer
    CurrYear As String * 4
    DprDay As Integer
End Type

Type FASetupRecType
    TOWNNAME As String * 25
    Pct1St   As Double
    PRate1St As String * 1
    Filler1  As String * 94
    DeprType As String * 25 'new for Windows
End Type

Type FADeptCodeType
    DeptDesc As String * 25
    DeptNum   As Integer
End Type

Type FAFundCodeType
    FundDesc As String * 25
    FundNum   As Integer
End Type

'Type PRNSetupRecType
'    Printer As String * 20
'    RPT(1 To 18) As Integer
'    CheckType As Integer
'End Type

Type TagNumbSortIdxType
  TagNumb As String * 20   '14
  DataRecNum As Integer    '2
End Type

Type ACNumbSortIdxType 'asset Code
  AssNumb As String * 20   '14
  AssRecNum As Integer    '2
End Type

Type DeptNumbSortIdxType 'dept Code
  DeptNumb As String * 20   '14
  DeptRecNum As Integer    '2
  DeptIdxDesc As String * 25
End Type

Type FundNumbSortIdxType 'fund Code
  FundNumb As String * 20   '14
  FundRecNum As Integer    '2
  FundIdxDesc As String * 25
End Type

Type DprSortIdxType
  DprNumb As String * 20
  DprRecNum As Long
End Type

Type TempVHCLDataType
  VHCLMAKE As String * 20
  VHCLMODL As String * 20
  VHCLVIN As String * 20
  VHCLTAG As String * 10
  VHCLCOLR As String * 10
  OPENFLAG As Integer
  ThisRec As Long
End Type

Type DprHistType
  PrevDprRec As Long
  ThisDept As Integer
  DprAmt As Double
  DprYear As String * 4
  ItemTag As String * 20
  DprToDate As Double
  ThisDesc1 As String * 30
  BookTotal As Double
  OrigCost As Double
  Life As Double
  PurchYear As String * 4
  LifeLeft As Integer
  SoSoftFlag As Boolean
End Type
  
Type TempDisposedOfDate
  DsplDate As Integer
End Type

Type PrePostDsplType
  DisposAmt As Double
  DsplMethod As String * 10
  ThisRec As Integer
  Deleted As Boolean
End Type

'Type CitiPassTempType
'  usernum   As Integer
'  UserName  As String * 15
'  frommdl   As Integer   'this is to indicate to citipak ok to have file
'End Type
