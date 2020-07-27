Attribute VB_Name = "FA_STRUCT"
Option Explicit

Const FAItemFile = "FAITEMS.DAT"
Const FACodeFile = "FACODES.DAT"

Type FAItemRecTypeV1
    ITEMTAG   As String * 20
    ISTATUS   As String * 1
    AQURDATE  As Integer
    IDESC1    As String * 30
    IDESC2    As String * 30
    IDESC3    As String * 30
    GLACCT    As String * 14
    IDEPT     As String * 4
    ASSETCODE As String * 4
    CodeRec   As Integer   'protected
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

Type FAItemRecType
    ITEMTAG  As String * 20
    ISTATUS  As String * 1
    DEPYN    As String * 1
    AQURDATE As Integer
    IDESC1   As String * 30
    IDESC2   As String * 30
    GLACCT   As String * 14
    IDEPT    As String * 4
    ASSETCODE As String * 4
    CodeRec  As Integer
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
    EOLDATE  As Integer
    Fill1     As String * 86
    FileVer   As Integer        ' "0" in ver1   "2" in ver2
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
End Type

Type FASetupRecType
    TownName As String * 25
    Pct1St   As Integer
    PRate1St As String * 1
    Filler1  As String * 100
End Type

Type UnitFileRecType
    UFEMPR   As String * 35
    UFATTN   As String * 35
    UFADDR1  As String * 35
    UFADDR2  As String * 35
    UFCITY   As String * 25
    UFSTATE  As String * 2
    UFZIP    As String * 10
    UFFEDID  As String * 14
    UFSTAID  As String * 14
    UFRETID  As String * 14
    UFRETIDL As String * 14
    ESCRTYPE As Integer
    TAXWBASE As Double
    BBTCNTNO As String * 12
    BBTBATCH As String * 12
    USEACH   As String * 1
    IMMDNUM  As String * 9
    IMMONUM  As String * 9
    DBANK    As String * 23
    OBANK    As String * 23
    FileVer  As Double
    BankDraft As String * 1
    '********added 11/11/02
    ESCRemitNum As String * 20
    ESCEmplrNum As String * 20

End Type

Type TaxRetLiabType
   Acct As String * 14
End Type

Type RegDSysFileRecType
    USEIMP   As String * 1
    CashAcct As String * 14
    IDRACCT  As String * 14
    ICRACCT  As String * 14
    Liab(1 To 5) As TaxRetLiabType
    CITIDIR  As String * 48
    SplitFlag As String * 1
    EXPMETHD As String * 1
    FRNGRATE As Double
    FRNGEXP  As String * 7
    FRNGDR   As String * 14
    FRNGCR   As String * 14
    INDRATE  As Double
    INDEXP   As String * 7
    INDDR    As String * 14
    INDCR    As String * 14
    SOCEXP   As String * 14
    SOCLIAB  As String * 14
    MEDEXP   As String * 14
    MEDLIAB  As String * 14
    RETEXP   As String * 14
    RETLIAB  As String * 14
    AcctCnt  As Integer
    GLActLen As Integer
    CheckStyle As Integer
    GLCheckYN As String * 1
'    VAC2SICK As String * 1
End Type

Type PRNSetupRecType
    Printer As String * 20
    RPT(1 To 18) As Integer
'    RPT(1 To 19) As Integer '8/13 went from 18 to 19
    'when "Checks by Number" report was added
    CheckType As Integer
End Type

Type GLTransRecType                 'Transaction Record: 96 bytes
   AcctRec     As Integer           'Pointer to Acct Record
   AcctNum     As String * 14       'Formatted Acct Number string
   TrDate      As Integer           'Date2Num function
   Desc        As String * 20       'Transaction Description
   Ref         As String * 8        'Document Reference
   DrAmt       As Double            'Debit Amount
   CrAmt       As Double            'Credit Amount
   Src         As String * 8        'Module Source Code
   NextTran    As Long              'Pointer to Next Trans
   Res         As String * 20       'Reserved for future needs
   Marked      As Integer
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

Type FAITEMS
    ITEMTAG  As String * 20
    ISTATUS  As String * 1
    DEPYN    As String * 1
    AQURDATE As Integer
    IDESC1   As String * 30
    IDESC2   As String * 30
    GLACCT   As String * 14
    IDEPT    As String * 4
    ASETCODE As String * 4
    CodeRec  As Integer
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
    EOLDATE  As Integer
End Type

Type JGLAcctIdxType
  AcctNum As String * 14
  RecNo   As Integer
End Type

Type GLAcctIdxType
  AcctNum As Double
  RecNo   As Single
End Type

Type GLSetupRecType                 'still under const.
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
   POStop      As Boolean    ' this is new 7/22/02 for tabstop on invoice entry
End Type

Type FLen2
  V As String * 64
End Type

Type Struct
  who As String * 14
  RecNum As Integer
End Type

