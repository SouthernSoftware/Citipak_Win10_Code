Attribute VB_Name = "W2STRUCT"
Option Explicit

Type PeriodDefaultRecType
    PACTIVE  As Integer
    PERBEG   As Integer
    PEREND   As Integer
    USEDEF   As String * 1

    PAYWK    As String * 1
    PAYBIWK  As String * 1
    PAYSEMIM As String * 1
    PAYMO    As String * 1
    PAYQTR   As String * 1
    PAYSEMIA As String * 1
    PAYANNL  As String * 1

    UseDed(1 To 50)   As String * 1
    USEAE1   As String * 1
    USEAE2   As String * 1
    USEAE3   As String * 1
    MACTIVE  As Integer
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
'    FileVer  As Double 'commented out 8/28/03 because of Wrightsville Beach
    FileVer As String * 7
    LMT401YN As String * 1
    BankDraft As String * 1
    '********added 11/11/02
    ESCRemitNum As String * 20
    ESCEmplrNum As String * 20
    '********added 8/31/04
    GMatch401K  As Double
    LMatch401K  As Double
    SSNOnCheck  As String * 1

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
End Type

'two records in file
Type EIC1RecType
    EIC1OVR0 As Double
    EIC1NVR0 As Double
    EIC1AMT0 As Double
    EIC1OVR1 As Double
    EIC1NVR1 As Double
    EIC1AMT1 As Double
    EIC1OVR2 As Double
    EIC1NVR2 As Double
    EIC1AMT2 As Double
    EIC1LESS As Double
    EIC1EXES As Double
End Type

Type EICRecType
    EIC(1 To 2) As EIC1RecType
End Type
'---------------------------------

Type RetireRecType
    TYPEDES1 As String * 20
    TYPEWH1  As Double
    TYPEM1   As Double
    TYPEOT1  As String * 1
    TYPETD1  As String * 1
End Type
'---------------------------------

Type PRPPDraftInfoType
    EmpRec     As Long
    DraftDate  As Integer
    NetPay     As Double
End Type

Type PRDraftRecType
    BANKDEST As String * 9
    BANKORIG As String * 9
    BANKNAME As String * 23
    BANKLOC  As String * 23
    FEDPREFX As String * 1
    FEDID As String * 9
End Type

Type DraftInfoFileName
    BANKDEST As String * 9
    BANKORIG As String * 9
    BANKNAME As String * 23
    BANKLOC  As String * 23
    FEDPREFX As String * 1
    FEDID As String * 9
End Type

Type UBDraftPayRecType
    CustAcctNum   As Long
    DraftAmt      As Double
End Type

Type ErnCodeRecType
    ERNCODE1 As String * 10
    ERNFWT1  As String * 1
    ERNSWT1  As String * 1
    ERNSOC1  As String * 1
    ERNMED1  As String * 1
    ERNRET1  As String * 1
    EarnYN   As String * 1
    Pad      As String * 10
End Type

Type DedCodeRecType
    DCDESC1  As String * 10
    DCACCT1  As String * 14
    DCFWT1   As String * 1
    DCSWT1   As String * 1
    DCSOC1   As String * 1
    DCMED1   As String * 1
End Type

Type StateTaxRecType
'--- Single

    TAX101 As Double
    TAX102 As Double
    TAX103 As Double
    TAX104 As Double
    TAX105 As Double
    TAX106 As Double
    TAX107 As Double
    STS(1 To 3, 1 To 12) As Double

'--- Married

    TAX201 As Double
    TAX202 As Double
    TAX203 As Double
    TAX204 As Double
    TAX205 As Double
    TAX206 As Double
    TAX207 As Double
    STM(1 To 3, 1 To 12) As Double

'--- Head of House

    TAX301 As Double
    TAX302 As Double
    TAX303 As Double
    TAX304 As Double
    TAX305 As Double
    TAX306 As Double
    TAX307 As Double
    STH(1 To 3, 1 To 12) As Double

'--- Georgia Page 4

    TAX401 As Double
    TAX402 As Double
    TAX403 As Double
    TAX404 As Double
    TAX405 As Double
    TAX406 As Double
    TAX407 As Double
    ST4(1 To 3, 1 To 12) As Double

End Type

Type FederalTaxRecType
'--- single
    FTSEMPSS As Double
    FTSEMRSS As Double
    FTSSSMW  As Double
    FTSEMPM  As Double
    FTSEMRM  As Double
    FTSMMW   As Double
    FTSSDAA  As Double

    FTS(1 To 3, 1 To 10) As Double

'--- married
    FTMEMPSS As Double
    FTMEMRSS As Double
    FTMSSMW  As Double
    FTMEMPM  As Double
    FTMEMRM  As Double
    FTMMMW   As Double
    FTMSDAA  As Double

    FTM(1 To 3, 1 To 10) As Double
End Type

'npayroll
Type EmployeeIndexType
    DataRecNum    As Integer
End Type

Type EmpNumType
    EmpNum        As String * 10
End Type

Type EmpDedType
    DPct        As String * 7
    DAmt        As Double
    DOTI        As String * 1
End Type

Type EmpWageDistType
    DAcct       As String * 14
    DAmt        As Double
End Type

Type EmpData1Type
    EmpNo       As String * 10     '1-10
    EmpLName    As String * 24     '11-34
    EmpFName    As String * 24     '35-58
    Data1RecNum As Integer         '59-60
    TransRecNum As Integer         '61-62
    Deleted     As Integer         '63-64
    'KillFlag    AS STRING * 1
    'FillerPad   AS STRING * 1
End Type

Type EmpData2Type         'new emp 2 rec
    EmpNo    As String * 10
    EmpSSN   As String * 11
    EmpLName As String * 24
    EmpFName As String * 24
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
'new
    DRAFTCOD As String * 1
    EMPDDACC As String * 20
    PRENOTED As String * 1
    BANKNAME As String * 33
    BANKLOC  As String * 30
    TRANSIT  As String * 9
'new
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
'------ EMPMA2
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

'    EmpDed(1 To 12)  As EmpDedType
    EmpDed(1 To 50)  As EmpDedType

'------ page 3
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

'------ Misc
    LastTransRec As Integer
    EmpPin       As Integer
    Deleted      As Integer
    'for new leave table stuff
' PreNoteFlag     AS INTEGER

'    Unused       AS STRING * 51

    LDTDate      As Integer      'last test date
    CDTDate      As Integer      'current test date
    InprocFlag   As Integer      'in process flag

    Unused       As String * 43
    CheckType    As Integer
    '******added 11/12/2002
    YN401K   As String * 1
    PrimeDept As String * 6
    HomePhone As String * 14
    EmrgncyCntctName As String * 48
    EmrgncyCntctPhnNum As String * 14
    EmrgncyCntctRelation As String * 16
    '*******added 8/31/04
    Comment      As String * 25
    
End Type

Type EmpData3Type
    Data1RecNum     As Integer
    YTDGrossPay     As Double
    YTDSocGrossPay  As Double
    YTDMedGrossPay  As Double
    YTDFedGrossPay  As Double
    YTDStaGrossPay  As Double
    YTDOTPay        As Double
    YTDRegPay       As Double
    YTDNet          As Double
    YTDSocial       As Double
    YTDMedicare     As Double
    YTDFederal      As Double
    YTDState        As Double
    YTDRetire       As Double
'    YTDDAmt(1 To 12) As Double
    YTDDAmt(1 To 50) As Double
    YTDDAmtT        As Double
    YTDEarn1        As Double        'e
    YTDEarn2        As Double
    YTDEarn3        As Double
    YTDEarnT        As Double
    YTDEIC          As Double
    YTDOther2       As Double
End Type

Type NameSortIdxType
  EmpName As String * 14
  DataRecNum As Integer
End Type

Type NumbSortIdxType
  EmpNumb As String * 14   '14
  DataRecNum As Integer    '2
End Type

Type EmployeeIdxType
  DataRecNum As Integer
End Type

'npayroll

Type TransWageDistType                      'Transaction Wage Distributions
    DAcct       As String * 14              'G/L Account (Dept)
    DRHrs       As Double                   'Reg Hours Distributed
    DOHrs       As Double                   'OT Hours Distributed
    DPct        As Double                   'Distribution Percent
    DRWage      As Double                   'Reg Wage Distributed
    DOWage      As Double                   'OT Wage Distributed
End Type

Type TransEarnDistType                      'Additional Earnings
    EAcct       As String * 14              'Default Add'l Earings Distribution Account (G/L)
    EAmt        As Double                   'Default Add'l Earings Amount
End Type

''''''''''''''
Type TransRecType
    TActive         As Integer             '1     'Active Transaction Flag
    PrevTransRec    As Integer             '2     'Pointer to employee's prev trans
    EmpPin          As Integer             '3     'Pointer to employee rec
    PaySFlag        As String * 1          '4     'Pay Salary Flag in time trans
    CheckNum        As Long                '5     'Payroll Check Number
    PayPdStart      As Integer             '6     'Start of Pay Period
    PayPdEnd        As Integer             '7     'End of Pay Period
    CheckDate       As Integer   'yea      '8     'Date checks written
    PostDate        As Integer             '9     'Date Transaction are posted
    PayType         As String * 1          '10    'Salaried or Hourly
    BaseRate        As Double              '11    'Base Rate or Salary Amt
    OTRate          As Double              '12    'Overtime Rate
    RegHrsWork      As Double             '13    'Hours worked this period
    VacUsed          As Double             '14    'vacation used this period
    SickUsed         As Double             '15    'Sick hours used this period
    CompUsed         As Double             '16    'comp hours used this period

    HOLHOURS         As Double             '17    'holiday hours used this period
    PerHours         As Double             '17a

    RegHrsPaid       As Double             '18    'sum of reg hours paid
    OTHours          As Double             '19    'OT hours this period
    OTHrsPaid        As Double             '20    'OT hours paid this period
    OT2Comp          As Double             '21    'Hours to comp time
    TDist(1 To 8)    As TransWageDistType  '      'Above TransWageDistType                              'wage distributions
                              '22 23 24 25 26 27 28 29
    TotRegWage     As Double               '30    'Total Reg Wage distributions
    TotOTWage      As Double               '31    'Total OT Wage distributions
    GrossWage      As Double               '32    'Reg Wage + OT Wage
    EAmt(1 To 3) As Double                 '      'Add Earnings amounts
                                     '33 34 35
    EDist(1 To 6)   As TransEarnDistType   '2     'Add Earnings distribitions (G/L) accs
    TotAdditEarn    As Double              '      'Total Additional Earnings
    GrossPay        As Double              '      'Add Earnings + GrossWage
    SocGrossPay     As Double              '      'Social Security Gross
    MedGrossPay     As Double              '      'Medicare Gross
    FedGrossPay     As Double              '      'Federal Gross
    StaGrossPay     As Double              '      'State Gross
    SocTaxAmt       As Double              '      'Social Security Tax W/H
    MedTaxAmt       As Double              '      'Medicare Tax W/H
    FedTaxAmt       As Double              '      'Fed Tax W/H
    StaTaxAmt       As Double              '      'State Tax W/H
    TotTaxAmt       As Double              '      'Total Taxes W/H
    RetireAmt       As Double              '      'Retirement W/H
    DAmt(1 To 50) As Double              '      'Voluntary Deduction amounts / pcts
    TotDedAmt As Double                    '      'Total Voluntary Deductions
    EICAmt     As Double                   '
    NetPay    As Double                    '
    PeriodHistRec As Integer    'not used        '      'YTD Totals?
    MatchRetAmt     As Double              '      'Employer's Retirement Match
    MatchSocAmt     As Double              '      'Employer's Social Secity Match
    MatchMedAmt     As Double              '      'Employer's Medicare Match
    RetGrossPay     As Double              '      'Retirement Gross
    TaxFring        As Double              '      'Taxable Fringe
    Less401k(1 To 3)  As Boolean
    Pad1            As String * 8
'-----------------------------
'can sum these for a report from transaction history file
End Type


Type W2DedType
    CHKDED  As String * 20
    AMTBOX  As String * 3
    DedCode As String * 4
End Type

Type W2SetUpType
                           'only two options per check box  15c 15g
    ExtrYear As Integer

    Deds(0 To 51) As W2DedType '11/16/04 changed from 50 to 51

End Type

Type W2FormType                      '  ===  ========
    FEDWAGE  As Double               '    1  FEDWAGE
    FEDTAXWH As Double               '    2  FEDTAXWH
    SOCWAGE  As Double               '    3  SOCWAGE
    SOCTAXWH As Double               '    4  SOCTAXWH
    MedWages As Double               '    5  MEDWAGES
    MEDTAXWH As Double               '    6  MEDTAXWH
    SocTips  As Double               '    7  SOCTIPS
    ALLOCTIP As Double               '    8  ALLOCTIP
    AdvEIC   As Double               '    9  ADVEIC
    DEPNDCAR As Double               '   10  DEPNDCAR
    NQPLAN   As Double               '   11  NQPLAN
    BENFBOX1 As Double               '   12  BENFBOX1
    BOX13AMT As Double               '   13  BOX13AMT
    BOX13TXt As String * 4           '   14  BOX13TXT
    BOX14AMT As Double               '   15  BOX14AMT
    BOX14TXT As String * 4           '   16  BOX14TXT
    BOX13AM1 As Double               '   17  BOX13AM1
    BOX13TX1 As String * 4           '   18  BOX13TX1
    BOX13AM2 As Double               '   added for Fall 04
    BOX13TX2 As String * 4           '   added for Fall 04
    BOX13AM3 As Double               '   added for Fall 04
    BOX13TX3 As String * 4           '   added for Fall 04
    BOX14AM1 As Double               '   19  BOX14AM1
    BOX14TX1 As String * 4           '   20  BOX14TX1
    BOX15A   As String * 1           '   21  BOX15A
    BOX15B   As String * 1           '   22  BOX15B
    BOX15c   As String * 1           '   23  BOX15C
    BOX15D   As String * 1           '   24  BOX15D
    BOX15E   As String * 1           '   25  BOX15E
    BOX15F   As String * 1           '   26  BOX15F
    BOX15G   As String * 1           '   27  BOX15G
    State    As String * 2           '   28  STATE
    STAWAGE  As Double               '   29  STAWAGE
    STATAXWH As Double               '   30  STATAXWH
    LOCALNAM As String * 16          '   31  LOCALNAM
    LocWage  As Double               '   32  LOCWAGE
    LOCALTAX As Double               '   33  LOCALTAX
    W3DfCmp1 As Double               'added for Fall04 update
    W3DfCmp2 As Double               'added for Fall04 update
    W3DfCmp3 As Double               'added for Fall04 update
    W3DfCmp4 As Double               'added for Fall04 update
End Type

Type W2ReprintIdxType
    RECNO As Long
    CONTNUM As Long
End Type

Type EarnNoMatchType
    EarnYN       As String * 1
    Pad         As String * 10
End Type

Type W2ElectronicSubRA
    EINNum As String * 9
    PersIDNum As String * 17
    ResubID As String * 1
    ReSubWFID As String * 6
    SftwrCode As String * 2
    CmpnyName As String * 57
    LocAddr As String * 22
    DelAddr As String * 22
    City As String * 22
    State As String * 2
    Zip As String * 5
    ZipExt As String * 4
    SubmttrName As String * 57
    SubLocAddr As String * 22
    SubDelAddr As String * 22
    SubCity As String * 22
    SubState As String * 2
    SubZip As String * 5
    SubZipExt As String * 4
    ContactName As String * 27
    CntctPhone As String * 15
    CntPhnExt As String * 5
    CntEMail As String * 40
    CntFAX As String * 10
    CntMethod As String * 1
    PrepCode As String * 1
End Type

Type W2ElectronicSubRE
  TaxYear As String * 4
  AgentCode As String * 1
  EmprAgntEIN As String * 9
  EINAgent As String * 9
  TermBusInd As String * 1
  OthEIN As String * 9
  EmprName As String * 57
  EmprLocAddr As String * 22
  EmprDelAddr As String * 22
  EmprCity As String * 22
  EmprState As String * 2
  EmprZip As String * 5
  EmprZipX As String * 4
  ThrdSckPay As String * 1
End Type

Type W2ElectronicSubRW
  EmpSSN As String * 9
  EmpFName As String * 15
  EmpMName As String * 15
  EmpLName As String * 20
  EmpSuffix As String * 4
  EmpAdd1 As String * 22
  EmpAdd2 As String * 22
  EmpCity As String * 22
  EmpState As String * 2
  EmpZip As String * 5
  EmpZipX As String * 4
  WageTips As String * 11
  FedTax As String * 11
  SSWages As String * 11
  SSTax As String * 11
  MedWages As String * 11
  MedTax As String * 11
  SSTips As String * 11
  AdvEIC As String * 11
  DepCare As String * 11
  Defr401k As String * 11
  Defr403b As String * 11
  Defr408k6 As String * 11
  Defr457b As String * 11
  Defr501c18D As String * 11
  NQPlan457 As String * 11
  NQPNot457 As String * 11
  LifeIns As String * 11
  NonStaStcks As String * 11
  StatuEmp As String * 1
  RetPlan As String * 1
  ThrdSckPay As String * 1
  ThrdSckAmt As String * 11
  RONum As String * 5
  Roth401K As String * 11
End Type

Type W2ElectronicSubRO
  AllocTips As String * 11
  TaxOnTips As String * 11
  MedSavings As String * 11
  RetAcct As String * 11
  AdoptionX As String * 11
  UnSSLife As String * 11
  UnMedLife As String * 11
  RecNum As Integer
End Type

Type W2ElectronicSubRU
  NumOfROs As String * 7
  AllocTips As String * 15
  TaxOnTips As String * 15
  MedSavings As String * 15
  RetAcct As String * 15
  AdoptionX As String * 15
  UnSSLife As String * 15
  UnMedLife As String * 15
End Type

Type W2ElectronicSubRT
  NumOfRWS As String * 7
  WagesTips As String * 15
  FedTax As String * 15
  SocWages As String * 15
  SocTax As String * 15
  MedWages As String * 15
  MedTax As String * 15
  SocTips As String * 15
  AdvEIC As String * 15
  DepCare As String * 15
  Defr401k As String * 15
  Defr403b As String * 15
  Defr408k6 As String * 15
  Defr457b As String * 15
  Defr501c18D As String * 15
  NQPlan457 As String * 15
  NQPNot457 As String * 15
  GrpTerm As String * 15
  ThrdTaxPay As String * 15
  NonStatStk As String * 15
  Roth401K As String * 15
End Type

Type W2ElectronicSubRF
  NumOfRWS As String * 9
End Type

Type W3FormType
  Control    As String * 18
  Payer      As String * 17
  ThirdSck   As Boolean
  EstabNum   As String * 22
  EmpIDNum   As String * 30
  EmpName    As String * 29
  Add1       As String * 29
  Add2       As String * 29
  City       As String * 29
  State1     As String * 2
  Zip        As String * 5
  ZipX       As String * 4
  OtherEIN   As String * 25
  State2     As String * 2
  StateID    As String * 25
  Contact    As String * 26
  Email      As String * 31
  Phone      As String * 15
  Fax        As String * 15
  Box13Sck   As String * 32
  Box14Sck   As Double
  LocWage    As Double
  LocTax     As Double
End Type

Type CitiPassTempType
  usernum   As Integer
  UserName  As String * 15
  frommdl   As Integer   'this is to indicate to citipak ok to have file
End Type
