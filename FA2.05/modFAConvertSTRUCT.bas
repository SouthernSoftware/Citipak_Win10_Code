Attribute VB_Name = "modFAConvertSTRUCT"
Option Explicit

Type DosFAItemRecType
    ITEMTAG  As String * 20
    ISTATUS  As String * 1
    DEPYN    As String * 1
    AQURDATE As Integer
    IDESC1   As String * 30
    IDESC2   As String * 30
    GLACCT   As String * 14
    IDEPT    As String * 4
    ASSETCODE As String * 4
    CODEREC  As Integer
    ILIFE    As Double
    ORGCOST  As Double
    DEP2DATE As Double
    CURRVAL  As Double
    CDEPDATE As Integer
    DISPDATE As Integer
    VENDOR   As String * 30
    SERIALNO As String * 30
    ITEMMFG  As String * 30
    CONTACT  As String * 30
    ITEMLOC  As String * 30
    EOLDate  As Integer
    Fill1     As String * 86
    FileVer   As Integer        ' "0" in ver1   "2" in ver2
End Type

Type FAItemRecType
    ITEMTAG  As String * 20
    ISTATUS  As String * 1
    DEPYN    As String * 1
    AQURDATE As Integer
    IDESC1   As String * 30
    IDESC2   As String * 30
    GLACCT   As String * 14
    IDEPT    As Integer
    ASSETCODE As String * 4
    ILIFE    As Double
    ORGCOST  As Double
    DEP2DATE As Double
    CURRVAL  As Double
    CDEPDATE As Integer
    DISPDATE As Integer
    VENDOR   As String * 30
    SERIALNO As String * 30
    ITEMMFG  As String * 30
    CONTACT  As String * 30
    ITEMLOC  As String * 30
    EOLDate  As Integer
    VHCLMAKE As String * 20
    VHCLMODL As String * 20
    VHCLVIN As String * 20
    VHCLTAG As String * 10
    VHCLCOLR As String * 10
    WARRXDAT As Integer
    Fill1     As String * 46
    FundNum As Integer  'new for Windows
    DisposAmt As Double  'new for Windows
    LastDprRec As Long
    LifeLeft As Integer
    PONum As String * 15
    CheckNum As String * 10
    DsplFlag As Integer '0 = Still In Use, 1 = Tagged for disposal, 2 = Disposed Of
    DsplMethod As String * 10
End Type


Type DosFAAssetCodeRecType
    ASSETCODE      As String * 4
    AssetStatus    As String * 10
    AssetDesc      As String * 20
End Type

Type FAAssetCodeRecType
    ASSETCODE      As String * 4
    AssetStatus    As String * 10
    AssetDesc      As String * 20
End Type


Type DosFAYearEndType
    LastYear  As String * 4
    CurYear As String * 4
End Type

Type FAYearEndType
    LastYear  As String * 4
    CurYear As String * 4
End Type


Type DosFADepFileType
    AssetRecord As Long
    CurYrDep As Double
    PctFlag  As Integer
End Type

Type FADepFileType
    AssetRecord As Long
    CurYrDep As Double
    PctFlag  As Integer
End Type


Type DosFASetupRecType
    TownName As String * 25
    Pct1St   As Integer
    PRate1St As String * 1
    Filler1  As String * 100
End Type

Type FASetupRecType
    TownName As String * 25
    Pct1St   As Double
    PRate1St As String * 1
    Filler1  As String * 94
    DeprType As String * 25 'new for Windows
End Type

Type TagNumbSortIdxType
    TagNumb As String * 20   '14
    DataRecNum As Integer    '2
End Type

Type FADeptCodeType
    DeptDesc As String * 25
    DeptNum   As Integer
End Type

Type DeptNumbSortIdxType 'dept Code
  DeptNumb As String * 20   '14
  DeptRecNum As Integer    '2
  DeptIdxDesc As String * 25
End Type

Type ACNumbSortIdxType 'asset Code
  AssNumb As String * 20   '14
  AssRecNum As Integer    '2
End Type

Type DosFAItemRecTypeV1
    ITEMTAG   As String * 20
    ISTATUS   As String * 1
    AQURDATE  As Integer
    IDESC1    As String * 30
    IDESC2    As String * 30
    IDESC3    As String * 30
    GLACCT    As String * 14
    IDEPT     As String * 4
    ASSETCODE As String * 4
    CODEREC   As Integer   'protected
    ILIFE     As Double
    ORGCOST   As Double
    DEP2DATE  As Double
    CDEPDATE  As Integer
    DISPDATE  As Integer
    VENDOR    As String * 30
    SERIALNO  As String * 30
    ITEMMFG   As String * 30
    CONTACT   As String * 30
    Fill1     As String * 99
End Type

