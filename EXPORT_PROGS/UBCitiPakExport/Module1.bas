Attribute VB_Name = "ConvertSTRUCT"
Option Explicit

'Type OldWinEmpDedType
'    DPct        As String * 7
'    DAmt        As Double
'    DOTI        As String * 1
'End Type

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
    EmrgncyCntctRelation As String * 14
    
End Type

'Type DosEmpData3Type
'    Data1RecNum     As Integer
'    YTDGrossPay     As Double
'    YTDSocGrossPay  As Double
'    YTDMedGrossPay  As Double
'    YTDFedGrossPay  As Double
'    YTDStaGrossPay  As Double
'    YTDOTPay        As Double
'    YTDRegPay       As Double
'    YTDNet          As Double
'    YTDSocial       As Double
'    YTDMedicare     As Double
'    YTDFederal      As Double
'    YTDState        As Double
'    YTDRetire       As Double
'    YTDDAmt(1 To 12) As Double
'    YTDDAmtT        As Double
'    YTDEarn1        As Double        'e
'    YTDEarn2        As Double
'    YTDEarn3        As Double
'    YTDEarnT        As Double
'    YTDEIC          As Double
'    YTDOther2       As Double
'End Type
'
'Type EmpData3Type
'    Data1RecNum     As Integer
'    YTDGrossPay     As Double
'    YTDSocGrossPay  As Double
'    YTDMedGrossPay  As Double
'    YTDFedGrossPay  As Double
'    YTDStaGrossPay  As Double
'    YTDOTPay        As Double
'    YTDRegPay       As Double
'    YTDNet          As Double
'    YTDSocial       As Double
'    YTDMedicare     As Double
'    YTDFederal      As Double
'    YTDState        As Double
'    YTDRetire       As Double
''    YTDDAmt(1 To 12) As Double
'    YTDDAmt(1 To 50) As Double
'    YTDDAmtT        As Double
'    YTDEarn1        As Double        'e
'    YTDEarn2        As Double
'    YTDEarn3        As Double
'    YTDEarnT        As Double
'    YTDEIC          As Double
'    YTDOther2       As Double
'End Type


'Type DosPRNSetupRecType
'    RPT1     As Integer
'    RPT2     As Integer
'    RPT3     As Integer
'    RPT4     As Integer
'    RPT5     As Integer
'    RPT6     As Integer
'    RPT7     As Integer
'    RPT8     As Integer
'    RPT9     As Integer
'    RPT10    As Integer
'    RPT11    As Integer
'    RPT12    As Integer
'    RPT13    As Integer
'    RPT14    As Integer
'    RPT15    As Integer
'    RPT16    As Integer
'End Type

'Type PRNSetupRecType
'    Printer As String * 20
'    RPT(1 To 18) As Integer
''    RPT(1 To 19) As Integer '8/14 added 1 to the array
'    'for new report "Checks by Number"
'    CheckType As Integer
'End Type

'Type EIC1RecType
'    EIC1OVR0 As Double
'    EIC1NVR0 As Double
'    EIC1AMT0 As Double
'    EIC1OVR1 As Double
'    EIC1NVR1 As Double
'    EIC1AMT1 As Double
'    EIC1OVR2 As Double
'    EIC1NVR2 As Double
'    EIC1AMT2 As Double
'    EIC1LESS As Double
'    EIC1EXES As Double
'End Type

'Type EICRecType
'    EIC(1 To 2) As EIC1RecType
'End Type
'
'Type TaxRetLiabType
'   Acct As String * 14
'End Type

'Type RegDSysFileRecType
'    USEIMP   As String * 1
'    CASHACCT As String * 14
'    IDRACCT  As String * 14
'    ICRACCT  As String * 14
'    Liab(1 To 5) As TaxRetLiabType
'    CITIDIR  As String * 48
'    SplitFlag As String * 1
'    EXPMETHD As String * 1
'    FRNGRATE As Double
'    FRNGEXP  As String * 7
'    FRNGDR   As String * 14
'    FRNGCR   As String * 14
'    INDRATE  As Double
'    INDEXP   As String * 7
'    INDDR    As String * 14
'    INDCR    As String * 14
'    SOCEXP   As String * 14
'    SOCLIAB  As String * 14
'    MEDEXP   As String * 14
'    MEDLIAB  As String * 14
'    RETEXP   As String * 14
'    RETLIAB  As String * 14
'    AcctCnt  As Integer
'    GLActLen As Integer
'    CheckStyle As Integer
'    GLCheckYN As String * 1
''    VAC2SICK As String * 1
'End Type

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


'Type RetireRecType
'    TYPEDES1 As String * 20
'    TYPEWH1  As Double
'    TYPEM1   As Double
'    TYPEOT1  As String * 1
'    TYPETD1  As String * 1
'End Type
'
'Type OldDedCodeRecType
'    DCDESC1  As String * 10
'    DCACCT1  As String * 14
'    DCFWT1   As String * 1
'    DCSWT1   As String * 1
'    DCSOC1   As String * 1
'    DCMED1   As String * 1
'End Type
'
'Type DedCodeRecType
'    DCDESC1  As String * 10
'    DCACCT1  As String * 14
'    DCFWT1   As String * 1
'    DCSWT1   As String * 1
'    DCSOC1   As String * 1
'    DCMED1   As String * 1
'End Type
'
'Type TransWageDistType                      'Transaction Wage Distributions
'    DAcct       As String * 14              'G/L Account (Dept)
'    DRHrs       As Double                   'Reg Hours Distributed
'    DOHrs       As Double                   'OT Hours Distributed
'    DPct        As Double                   'Distribution Percent
'    DRWage      As Double                   'Reg Wage Distributed
'    DOWage      As Double                   'OT Wage Distributed
'End Type
'
'Type TransEarnDistType                      'Additional Earnings
'    EAcct       As String * 14              'Default Add'l Earings Distribution Account (G/L)
'    EAmt        As Double                   'Default Add'l Earings Amount
'End Type
'
'Type TransRecType
'    TActive         As Integer             '1     'Active Transaction Flag
'    PrevTransRec    As Integer             '2     'Pointer to employee's prev trans
'    EmpPin          As Integer             '3     'Pointer to employee rec
'    PaySFlag        As String * 1          '4     'Pay Salary Flag in time trans
'    CheckNum        As Long                '5     'Payroll Check Number
'    PayPdStart      As Integer             '6     'Start of Pay Period
'    PayPdEnd        As Integer             '7     'End of Pay Period
'    CheckDate       As Integer   'yea      '8     'Date checks written
'    PostDate        As Integer             '9     'Date Transaction are posted
'    PayType         As String * 1          '10    'Salaried or Hourly
'    BaseRate        As Double              '11    'Base Rate or Salary Amt
'    OTRate          As Double              '12    'Overtime Rate
'    RegHrsWork      As Double             '13    'Hours worked this period
'    VacUsed          As Double             '14    'vacation used this period
'    SickUsed         As Double             '15    'Sick hours used this period
'    CompUsed         As Double             '16    'comp hours used this period
'
'    HOLHOURS         As Double             '17    'holiday hours used this period
'    PerHours         As Double             '17a
'
'    RegHrsPaid       As Double             '18    'sum of reg hours paid
'    OTHours          As Double             '19    'OT hours this period
'    OTHrsPaid        As Double             '20    'OT hours paid this period
'    OT2COMP          As Double             '21    'Hours to comp time
'    TDist(1 To 8)    As TransWageDistType  '      'Above TransWageDistType                              'wage distributions
'                              '22 23 24 25 26 27 28 29
'    TotRegWage     As Double               '30    'Total Reg Wage distributions
'    TotOTWage      As Double               '31    'Total OT Wage distributions
'    GrossWage      As Double               '32    'Reg Wage + OT Wage
'    EAmt(1 To 3) As Double                 '      'Add Earnings amounts
'                                     '33 34 35
'    EDist(1 To 6)   As TransEarnDistType   '2     'Add Earnings distribitions (G/L) accs
'    TotAdditEarn    As Double              '      'Total Additional Earnings
'    GrossPay        As Double              '      'Add Earnings + GrossWage
'    SocGrossPay     As Double              '      'Social Security Gross
'    MedGrossPay     As Double              '      'Medicare Gross
'    FedGrossPay     As Double              '      'Federal Gross
'    StaGrossPay     As Double              '      'State Gross
'    SocTaxAmt       As Double              '      'Social Security Tax W/H
'    MedTaxAmt       As Double              '      'Medicare Tax W/H
'    FedTaxAmt       As Double              '      'Fed Tax W/H
'    StaTaxAmt       As Double              '      'State Tax W/H
'    TotTaxAmt       As Double              '      'Total Taxes W/H
'    RetireAmt       As Double              '      'Retirement W/H
'    DAmt(1 To 50) As Double              '      'Voluntary Deduction amounts / pcts
'    TotDedAmt As Double                    '      'Total Voluntary Deductions
'    EICAmt     As Double                   '
'    NetPay    As Double                    '
'    PeriodHistRec As Integer    'not used        '      'YTD Totals?
'    MatchRetAmt     As Double              '      'Employer's Retirement Match
'    MatchSocAmt     As Double              '      'Employer's Social Secity Match
'    MatchMedAmt     As Double              '      'Employer's Medicare Match
'    RetGrossPay     As Double              '      'Retirement Gross
'    TaxFring        As Double              '      'Taxable Fringe
'    Pad1            As String * 14
''-----------------------------
''can sum these for a report from transaction history file
'End Type
'
'Type DosTransRecType
'    TActive         As Integer             '1     'Active Transaction Flag
'    PrevTransRec    As Integer             '2     'Pointer to employee's prev trans
'    EmpPin          As Integer             '3     'Pointer to employee rec
'    PaySFlag        As String * 1          '4     'Pay Salary Flag in time trans
'    CheckNum        As Long                '5     'Payroll Check Number
'    PayPdStart      As Integer             '6     'Start of Pay Period
'    PayPdEnd        As Integer             '7     'End of Pay Period
'    CheckDate       As Integer   'yea      '8     'Date checks written
'    PostDate        As Integer             '9     'Date Transaction are posted
'    PayType         As String * 1          '10    'Salaried or Hourly
'    BaseRate        As Double              '11    'Base Rate or Salary Amt
'    OTRate          As Double              '12    'Overtime Rate
'    RegHrsWork      As Double             '13    'Hours worked this period
'    VacUsed          As Double             '14    'vacation used this period
'    SickUsed         As Double             '15    'Sick hours used this period
'    CompUsed         As Double             '16    'comp hours used this period
'
'    HOLHOURS         As Double             '17    'holiday hours used this period
'    PerHours         As Double             '17a
'
'    RegHrsPaid       As Double             '18    'sum of reg hours paid
'    OTHours          As Double             '19    'OT hours this period
'    OTHrsPaid        As Double             '20    'OT hours paid this period
'    OT2COMP          As Double             '21    'Hours to comp time
'    TDist(1 To 8)    As TransWageDistType  '      'Above TransWageDistType                              'wage distributions
'                              '22 23 24 25 26 27 28 29
'    TotRegWage     As Double               '30    'Total Reg Wage distributions
'    TotOTWage      As Double               '31    'Total OT Wage distributions
'    GrossWage      As Double               '32    'Reg Wage + OT Wage
'    EAmt(1 To 3) As Double                 '      'Add Earnings amounts
'                                     '33 34 35
'    EDist(1 To 6)   As TransEarnDistType   '2     'Add Earnings distribitions (G/L) accs
'    TotAdditEarn    As Double              '      'Total Additional Earnings
'    GrossPay        As Double              '      'Add Earnings + GrossWage
'    SocGrossPay     As Double              '      'Social Security Gross
'    MedGrossPay     As Double              '      'Medicare Gross
'    FedGrossPay     As Double              '      'Federal Gross
'    StaGrossPay     As Double              '      'State Gross
'    SocTaxAmt       As Double              '      'Social Security Tax W/H
'    MedTaxAmt       As Double              '      'Medicare Tax W/H
'    FedTaxAmt       As Double              '      'Fed Tax W/H
'    StaTaxAmt       As Double              '      'State Tax W/H
'    TotTaxAmt       As Double              '      'Total Taxes W/H
'    RetireAmt       As Double              '      'Retirement W/H
'    DAmt(1 To 12) As Double               '      'Voluntary Deduction amounts / pcts
'    TotDedAmt As Double                    '      'Total Voluntary Deductions
'    EICAmt     As Double                   '
'    NetPay    As Double                    '
'    PeriodHistRec As Integer    'not used        '      'YTD Totals?
'    MatchRetAmt     As Double              '      'Employer's Retirement Match
'    MatchSocAmt     As Double              '      'Employer's Social Secity Match
'    MatchMedAmt     As Double              '      'Employer's Medicare Match
'    RetGrossPay     As Double              '      'Retirement Gross
'    TaxFring        As Double              '      'Taxable Fringe
'    Pad1            As String * 38
''-----------------------------
''can sum these for a report from transaction history file
'End Type
'
'Type PRDEDType
'   DCode       As String * 10
'   DAmt        As Double
'   YTDDAmt     As Double
'End Type
'
'
'Type DosPRCheckRecType
'   CActive       As Integer
'   CheckNum      As Long
'   CheckDate     As Integer
'
'   EmpName       As String * 33
'   EmpNo         As String * 10
'   EmpSSN        As String * 11
'
''=-=-=-=-=-
'   EmpAddr1 As String * 36
'   EmpCity  As String * 24
'   EmpState As String * 2
'   EmpZip   As String * 10
''-=-=-=-=-=-=
'
'   PayEndDate    As Integer
'   BaseRate      As Double
'   GrossPay      As Double
'   FedTaxAmt     As Double
'   StaTaxAmt     As Double
'   MedTaxAmt     As Double
'   SocTaxAmt     As Double
'   TotDedAmt     As Double
'
''added
'   RetireAmt     As Double
'
'   NetPay        As Double
'   YTDGrossPay   As Double
'   YTDFederal    As Double
'   YTDState      As Double
'   YTDSocial     As Double
'   YTDMedicare   As Double
'   YTDTotDed     As Double
'   YTDNetPay     As Double
'
''added
'   YTDRetire     As Double
'
'   VactBal       As Double   '
'   SickBal       As Double   '
'   CompBal       As Double
'
''added-----
'   CompEarn      As Double
'   RegHrsWork    As Double
'   OTHrsPaid     As Double
'   TotRegWage    As Double
'   VacUsed       As Double
'   SickUsed      As Double
'   CompUsed      As Double
'   HolUsed       As Double
'   PerUsed       As Double
'
'   RegHrsPaid    As Double
'   TotOTWage     As Double
'
'   AEarn(1 To 3) As PRDEDType
'
'   TotAdditEarn  As Double
'
'   EICAmt        As Double
'   TaxFring      As Double
'
''----------
'  CDED(1 To 12) As PRDEDType
''    CDED(1 To 50) As PRDEDType
'   DDFlag        As Integer
'End Type
'
'Type PRCheckRecType
'   CActive       As Integer
'   CheckNum      As Long
'   CheckDate     As Integer
'
'   EmpName       As String * 33
'   EmpNo         As String * 10
'   EmpSSN        As String * 11
'
''=-=-=-=-=-
'   EmpAddr1 As String * 36
'   EmpCity  As String * 24
'   EmpState As String * 2
'   EmpZip   As String * 10
''-=-=-=-=-=-=
'
'   PayEndDate    As Integer
'   BaseRate      As Double
'   GrossPay      As Double
'   FedTaxAmt     As Double
'   StaTaxAmt     As Double
'   MedTaxAmt     As Double
'   SocTaxAmt     As Double
'   TotDedAmt     As Double
'
''added
'   RetireAmt     As Double
'
'   NetPay        As Double
'   YTDGrossPay   As Double
'   YTDFederal    As Double
'   YTDState      As Double
'   YTDSocial     As Double
'   YTDMedicare   As Double
'   YTDTotDed     As Double
'   YTDNetPay     As Double
'
''added
'   YTDRetire     As Double
'
'   VactBal       As Double   '
'   SickBal       As Double   '
'   CompBal       As Double
'
''added-----
'   CompEarn      As Double
'   RegHrsWork    As Double
'   OTHrsPaid     As Double
'   TotRegWage    As Double
'   VacUsed       As Double
'   SickUsed      As Double
'   CompUsed      As Double
'   HolUsed       As Double
'   PerUsed       As Double
'
'   RegHrsPaid    As Double
'   TotOTWage     As Double
'
'   AEarn(1 To 3) As PRDEDType
'
'   TotAdditEarn  As Double
'
'   EICAmt        As Double
'   TaxFring      As Double
'
''----------
'  CDED(1 To 50) As PRDEDType
'   DDFlag        As Integer
'End Type
'
'Type DraftInfoFileName
'    BANKDEST As String * 9
'    BANKORIG As String * 9
'    BANKNAME As String * 23
'    BANKLOC  As String * 23
'    FEDPREFX As String * 1
'    FEDID As String * 9
'End Type
'
'Type W2DedType
'    CHKDED  As String * 20
'    AMTBOX  As String * 3
'    DedCode As String * 4
'End Type
'
'Type DosW2SetUpType
'                           'only two options per check box  15c 15g
'    ExtrYear As Integer
''    CHKDED0  AS STRING * 20
''    AMTBOX0  AS STRING * 3
''    RETCODE  AS STRING * 4
'
'    Deds(0 To 12) As W2DedType
'
'End Type
'
'Type W2SetUpType
'                           'only two options per check box  15c 15g
'    ExtrYear As Integer
''    CHKDED0  AS STRING * 20
''    AMTBOX0  AS STRING * 3
''    RETCODE  AS STRING * 4
'
'    Deds(0 To 50) As W2DedType
'
'End Type
'
'Type PeriodDefaultRecType
'    PACTIVE  As Integer
'    PERBEG   As Integer
'    PEREND   As Integer
'    USEDEF   As String * 1
'
'    PAYWK    As String * 1
'    PAYBIWK  As String * 1
'    PAYSEMIM As String * 1
'    PAYMO    As String * 1
'    PAYQTR   As String * 1
'    PAYSEMIA As String * 1
'    PAYANNL  As String * 1
'
'    UseDed(1 To 50)   As String * 1
'    USEAE1   As String * 1
'    USEAE2   As String * 1
'    USEAE3   As String * 1
'    MACTIVE  As Integer
'End Type
'
'Type DosPeriodDefaultRecType
'    PACTIVE  As Integer
'    PERBEG   As Integer
'    PEREND   As Integer
'    USEDEF   As String * 1
'
'    PAYWK    As String * 1
'    PAYBIWK  As String * 1
'    PAYSEMIM As String * 1
'    PAYMO    As String * 1
'    PAYQTR   As String * 1
'    PAYSEMIA As String * 1
'    PAYANNL  As String * 1
'
'    UseDed(1 To 12)   As String * 1
'    USEAE1   As String * 1
'    USEAE2   As String * 1
'    USEAE3   As String * 1
'    MACTIVE  As Integer
'End Type

