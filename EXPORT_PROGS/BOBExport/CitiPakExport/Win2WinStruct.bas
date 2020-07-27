Attribute VB_Name = "Win2WinStruct"
Option Explicit

Type EmpDedType
    DPct        As String * 7
    DAmt        As Double
    DOTI        As String * 1
End Type

Type EmpWageDistType
    DAcct       As String * 14
    DAmt        As Double
End Type

Type OldWinEmpData2Type
    EmpNo    As String * 10
    EmpSSN   As String * 11
    EMPLNAME As String * 24
    EMPFNAME As String * 24
    EmpAddr1 As String * 36
    EMPADDR2 As String * 36
    EmpCity  As String * 24
    EmpState As String * 2
    EmpZip   As String * 10
    EMPBDAY  As Integer
    EMPGENDR As String * 6
    EMPRACE  As String * 14
    EMPRETNO As String * 16
    EMPRETTP As String * 24
    DRAFTCOD As String * 1
    EMPDDACC As String * 20
    PRENOTED As String * 1
    BANKNAME As String * 33
    BANKLOC  As String * 30
    TRANSIT  As String * 9
    EMPJOB   As String * 28
    EMPWCCLS As String * 12
    EMPSTATS As String * 10
    EMPBCODE As Double
    EMPPTYPE As String * 10
    EMPPFREQ As String * 16
    EMPPRATE As Double
    EMPORATE As Double
    EMPHDATE As Integer
    EMPRDATE As Integer
    EMPTDATE As Integer
    EMPFEDX  As String * 1
    EMPFEDO2 As String * 1
    EMPFEDO1 As Double
    EMPFEDS  As String * 1
    EMPFEDA  As Integer       'num of allowance
    EMPFEDAA As Double
    EMPSTAX  As String * 1
    EMPSTAO2 As String * 1
    EMPSTAO1 As Double
    EMPSTAS  As String * 1
    EMPSTAA  As Integer       '
    EMPSTAAA As Double
    EMPSOCX  As String * 1
    EMPMEDX  As String * 1
    EMPEIC   As String * 1

    EmpDed(1 To 50)  As EmpDedType

    EMPEACT1 As String * 14
    EMPEAMT1 As Double
    EMPEACT2 As String * 14
    EMPEAMT2 As Double
    EMPEACT3 As String * 14
    EMPEAMT3 As Double

    EMPHP    As String * 1

    EDist(1 To 8)   As EmpWageDistType

    EMPVACE  As Double
    EMPVUSED As Double
    EMPVBAL  As Double

    EMPSLE   As Double
    EMPSLUSE As Double
    EMPSLBAL As Double

    EMPCTE   As Double
    EMPCTUSE As Double
    EMPCTBAL As Double

    PERERN   As Double
    PerUsed  As Double
    PERBAL   As Double

    HOLERN   As Double
    HolUsed  As Double
    HOLBAL   As Double

    LeaveTbl As Integer

    ExcludeESC  As String * 1
    UseLife     As String * 1

    LastTransRec As Integer
    EmpPin       As Integer
    Deleted      As Integer

    LDTDate      As Integer      'last test date
    CDTDate      As Integer      'current test date
    InprocFlag   As Integer      'in process flag

    Unused       As String * 43
    CheckType    As Integer
End Type

Type EmpData2Type
    EmpNo    As String * 10
    EmpSSN   As String * 11
    EMPLNAME As String * 24
    EMPFNAME As String * 24
    EmpAddr1 As String * 36
    EMPADDR2 As String * 36
    EmpCity  As String * 24
    EmpState As String * 2
    EmpZip   As String * 10
    EMPBDAY  As Integer
    EMPGENDR As String * 6
    EMPRACE  As String * 14
    EMPRETNO As String * 16
    EMPRETTP As String * 24
    DRAFTCOD As String * 1
    EMPDDACC As String * 20
    PRENOTED As String * 1
    BANKNAME As String * 33
    BANKLOC  As String * 30
    TRANSIT  As String * 9
    EMPJOB   As String * 28
    EMPWCCLS As String * 12
    EMPSTATS As String * 10
    EMPBCODE As Double
    EMPPTYPE As String * 10
    EMPPFREQ As String * 16
    EMPPRATE As Double
    EMPORATE As Double
    EMPHDATE As Integer
    EMPRDATE As Integer
    EMPTDATE As Integer
    EMPFEDX  As String * 1
    EMPFEDO2 As String * 1
    EMPFEDO1 As Double
    EMPFEDS  As String * 1
    EMPFEDA  As Integer       'num of allowance
    EMPFEDAA As Double
    EMPSTAX  As String * 1
    EMPSTAO2 As String * 1
    EMPSTAO1 As Double
    EMPSTAS  As String * 1
    EMPSTAA  As Integer       '
    EMPSTAAA As Double
    EMPSOCX  As String * 1
    EMPMEDX  As String * 1
    EMPEIC   As String * 1

    EmpDed(1 To 50)  As EmpDedType

    EMPEACT1 As String * 14
    EMPEAMT1 As Double
    EMPEACT2 As String * 14
    EMPEAMT2 As Double
    EMPEACT3 As String * 14
    EMPEAMT3 As Double

    EMPHP    As String * 1

    EDist(1 To 8)   As EmpWageDistType

    EMPVACE  As Double
    EMPVUSED As Double
    EMPVBAL  As Double

    EMPSLE   As Double
    EMPSLUSE As Double
    EMPSLBAL As Double

    EMPCTE   As Double
    EMPCTUSE As Double
    EMPCTBAL As Double

    PERERN   As Double
    PerUsed  As Double
    PERBAL   As Double

    HOLERN   As Double
    HolUsed  As Double
    HOLBAL   As Double

    LeaveTbl As Integer

    ExcludeESC  As String * 1
    UseLife     As String * 1

    LastTransRec As Integer
    EmpPin       As Integer
    Deleted      As Integer

    LDTDate      As Integer      'last test date
    CDTDate      As Integer      'current test date
    InprocFlag   As Integer      'in process flag

    Unused       As String * 43
    CheckType    As Integer
    '*******added 11/11/02
    YN401K   As String * 1
    PrimeDept As String * 6
    HomePhone As String * 14
    EmrgncyCntctName As String * 48
    EmrgncyCntctPhnNum As String * 14
    EmrgncyCntctRelation As String * 16
    
End Type


Type OldWinUnitFileRecType
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




