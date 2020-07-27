Attribute VB_Name = "PRCheck_STRUCT"
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

Type DraftInfoFileName
    BANKDEST As String * 9
    BANKORIG As String * 9
    BankName As String * 23
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
    EMPLNAME    As String * 24     '11-34
    EMPFNAME    As String * 24     '35-58
    Data1RecNum As Integer         '59-60
    TransRecNum As Integer         '61-62
    Deleted     As Integer         '63-64
    'KillFlag    AS STRING * 1
    'FillerPad   AS STRING * 1
End Type

Type EmpData2Type         'new emp 2 rec
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
'new
    DRAFTCOD As String * 1
    EMPDDACC As String * 20
    PRENOTED As String * 1
    BankName As String * 33
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
    '*******added 11/11/02
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

Type ManualTransRecType
    PDSTART  As Integer           ' 1
    PDEND    As Integer           '
    CHKDATE  As Integer           '
    CheckNum As Long              '
    RegHrs   As Double            ' 5
    SICKHRS  As Double            '  6
    VACHRS   As Double             ' 7
    COMPHRS  As Double             ' 8
    PERSHRS  As Double
    HOLHOURS As Double             ' 9
    OTHRSPD  As Double             ' 10

    DISTACT1 As String * 14         '11
    WAGEAMT1 As Double              '12
    DISTACT2 As String * 14         '13
    WAGEAMT2 As Double              '14
    DISTACT3 As String * 14         '15
    WAGEAMT3 As Double              '16
    DISTACT4 As String * 14         '17
    WAGEAMT4 As Double              '18

    GrossPay As Double              '19
    RegWage  As Double              '20
    OTWage   As Double              '21
    FEDTAX   As Double              '22
    STATAX   As Double              '23
    SOCTAX   As Double              '24
    MEDTAX   As Double              '25
    RETAMT   As Double              '26

    DAmt(1 To 50)  As Double        '27-38

    TOTDED   As Double              '39
    EIC      As Double              '40
    NetPay   As Double              '41
    FedGross As Double              '42
    STAGROSS As Double              '43
    SocGross As Double              '44
    MedGross As Double              '45
    RETGROSS As Double              '46
    TOTTAX   As Double              '47  this is hidden and protected
End Type

Type DistWageRptType                        'For the register?
    Acct       As String * 14
    Pct        As Double
    RHrs       As Double
    OHrs       As Double
    RWage      As Double
    OWage      As Double
    AddEarn    As Double
    GrossPay   As Double
    MATSocAmt  As Double
    MATMedAmt  As Double
    MATRetAmt  As Double
End Type

Type YTDFundMnthType
  RegHrs    As Double
  OTHrs     As Double
  RegWage   As Double
  OTWage    As Double
'  VACHRS    AS DOUBLE
'  SICKHRS   AS DOUBLE
'  HOLHRS    AS DOUBLE
'  CTHrs     AS DOUBLE
'  CWHrs     AS DOUBLE
End Type


Type YTDFundRptType
  FundNum As Long
  Mths(1 To 12) As YTDFundMnthType
End Type

Type QtrDateType
  LDate   As Integer
  HDate   As Integer
End Type

Type K401RptType
'062498 Added RetType & EmpNum
  RetType As String * 1
  EmpNum  As String * 10
  EmpName As String * 32
  SSN     As String * 11
  VAmt    As String * 11
  LAmt    As String * 11
  MAmt    As String * 11
  Batch   As String * 11
  HDate   As String * 10
  Gross   As String * 10
End Type

Type DedRptType
  EmpName As String * 35
  SSN     As String * 11
  DAmt    As String * 23
  CrLf    As String * 2
End Type

Type EmpTransHistRptType
  EmpName        As String * 35
  SSN            As String * 11
  CheckNum       As Long
  CheckDate      As Integer   'yea
  GrossPay       As Double    'this is add earnings and wages combined
  'GrossWage      AS DOUBLE   'this is also in the transaction history
  TotDedAmt      As Double
  NetPay         As Double
End Type

Type EmpHistFormType
    FirstEmp     As Long
    LastEmp      As Long
    StartDate    As Integer
    EndDate      As Integer
    SumOnly      As String * 1
End Type

Type EmpHistoryRptType
    RegHrs       As Double
    VACHRS       As Double
    SICKHRS      As Double
    HOLHRS       As Double
    COMPHRS      As Double
    TotalHrs     As Double
'*******
    PHrs         As Double
'*******
    TOTHrs       As Double

    TOTPaid      As Double
    TotEIC       As Double
    TRegWage     As Double
    TOTWage      As Double
    GPay         As Double
    SSTax        As Double
    MTax         As Double
    FTax         As Double
    STax         As Double
    RETTOT       As Double
    RETMATT      As Double
    TNetPay      As Double
End Type

Type GrossWageRptType
    EmpNo        As String * 11
    EmpName      As String * 21
    GrossPay     As Double
    FedGross     As Double
    FEDTAX       As Double
    SocGross     As Double
    SOCTAX       As Double
    MedGross     As Double
    MEDTAX       As Double
    'StaGross     AS DOUBLE
    EIC          As Double
    STATAX       As Double
End Type

Type GrossFundsType
   FundNum As Integer
   DAmt    As Double
'   FedGross AS DOUBLE
'   FedTax   AS DOUBLE
'   StaGross AS DOUBLE
'   StaTax   AS DOUBLE
'   MedGross AS DOUBLE
'   MedTax   AS DOUBLE
'   SocGross AS DOUBLE
'   SocTax   AS DOUBLE
'   RetGross AS DOUBLE
'   RetTax   AS DOUBLE
End Type


Type ESCGrossWageRptType
    EmpName      As String * 31
    ESSN         As String * 11
    GrossPay     As Double
End Type

Type ESC2DiskRecType1
    Qtr          As String * 9     '1 - 9
    Fill1        As String * 13    '10-22
    ESSN         As String * 16    '23-38
    EName        As String * 24    '39-62
    Seasonal     As String * 1     '63
    GPay         As String * 15    '64-78
    CrLf         As String * 2     '79- 80
End Type

Type EQWRptRecType
  ENumb    As String * 11
  EName    As String * 21
  GPay     As String * 11
  FedGr    As String * 11
  FedTx    As String * 11
  SocGr    As String * 11
  SocTx    As String * 11
  MedGr    As String * 11
  MedTx    As String * 11
  'StaGr    AS STRING * 11
  EIC      As String * 11
  StaTx    As String * 11
  CrLf     As String * 2
End Type

Type ESCMAG2DiskType
   Blank1       As String * 1
   SSN          As String * 9
   LastName     As String * 12
   EmpInitials  As String * 2
   EmpWages     As String * 9
   SeasInd      As String * 1   '"N"
   RemitNumb    As String * 6
   EmplrAcct    As String * 7   'Employer Account Number
   BranchAcct   As String * 7
   RQuarter     As String * 1
   RYear        As String * 4
   EmplrName    As String * 20
   Blank2       As String * 1
   CrLf         As String * 2
End Type

Type CompSortType
  CompCode As String * 12
  RecNo    As Long
End Type

Type EmpSortType
  EmpNo As String * 14
  RecNo As Integer
End Type
 
Type LeaveEntryType
    YEARS   As Integer
    EARN    As Double
End Type

Type LeaveRecType
  VacMax   As Double
  VEntry(1 To 20)  As LeaveEntryType
  SICKMAX  As Double
  SEntry(1 To 20)  As LeaveEntryType
  PerMax   As Double
  PEntry(1 To 20)  As LeaveEntryType
  HolMax   As Double
  HEntry(1 To 20)  As LeaveEntryType
End Type

Type DosPRNSetupRecType
    RPT1     As Integer
    RPT2     As Integer
    RPT3     As Integer
    RPT4     As Integer
    RPT5     As Integer
    RPT6     As Integer
    RPT7     As Integer
    RPT8     As Integer
    RPT9     As Integer
    RPT10    As Integer
    RPT11    As Integer
    RPT12    As Integer
    RPT13    As Integer
    RPT14    As Integer
    RPT15    As Integer
    RPT16    As Integer
End Type

Type PRNSetupRecType
    Printer As String * 20
    RPT(1 To 18) As Integer
    CheckType As Integer
End Type


Type PRDEDType
   DCode       As String * 10
   DAmt        As Double
   YTDDAmt     As Double
End Type

Type PRCheckRecType
   CActive       As Integer
   CheckNum      As Long
   CheckDate     As Integer

   EmpName       As String * 33
   EmpNo         As String * 10
   EmpSSN        As String * 11

'=-=-=-=-=-
   EmpAddr1 As String * 36
   EmpCity  As String * 24
   EmpState As String * 2
   EmpZip   As String * 10
'-=-=-=-=-=-=

   PayEndDate    As Integer
   BaseRate      As Double
   GrossPay      As Double
   FedTaxAmt     As Double
   StaTaxAmt     As Double
   MedTaxAmt     As Double
   SocTaxAmt     As Double
   TotDedAmt     As Double

'added
   RetireAmt     As Double

   NetPay        As Double
   YTDGrossPay   As Double
   YTDFederal    As Double
   YTDState      As Double
   YTDSocial     As Double
   YTDMedicare   As Double
   YTDTotDed     As Double
   YTDNetPay     As Double

'added
   YTDRetire     As Double

   VactBal       As Double   '
   SickBal       As Double   '
   CompBal       As Double

'added-----
   CompEarn      As Double
   RegHrsWork    As Double
   OTHrsPaid     As Double
   TotRegWage    As Double
   VacUsed       As Double
   SickUsed      As Double
   CompUsed      As Double
   HolUsed       As Double
   PerUsed       As Double

   RegHrsPaid    As Double
   TotOTWage     As Double

   AEarn(1 To 3) As PRDEDType

   TotAdditEarn  As Double

   EICAmt        As Double
   TaxFring      As Double

'----------
   CDED(1 To 50) As PRDEDType
   DDFlag        As Integer
End Type

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

Type ScrnCalcType
   REGEARN As Double
   OTEARN As Double
   ALTEARN1 As Double
   ALTEARN2 As Double
   ALTEARN3 As Double
'   DED1 As Double
'   DED2 As Double
'   DED3 As Double
'   DED4 As Double
'   DED5 As Double
'   DED6 As Double
'   DED7 As Double
'   DED8 As Double
'   DED9 As Double
'   DED10 As Double
'   DED11 As Double
'   DED12 As Double
   Ded(1 To 50) As Double
   SOCTAX As Double
   MEDTAX As Double
   FEDTAX As Double
   STATAX As Double
   RETIRE As Double
   GrossPay As Double
   TOTDED As Double
   EIC As Double
   NetPay As Double
  
End Type

Type HourlyIDistType
    DAcct    As String * 14    '  10 WAGEDST1 AS STRING * 14
    DRHrs    As Double         '  11 D1REGHRS AS DOUBLE
    DOHrs    As Double         '  12 D1OTHRS  AS DOUBLE
End Type

Type HourlyInputType
    WORKHRS  As Double         '  1  WORKHRS  AS DOUBLE
    VACHRS   As Double         '  2  VACHRS   AS DOUBLE
    SICKHRS  As Double         '  3  SICKHRS  AS DOUBLE
    HOLHRS   As Double         '  4  HOLHRS   AS DOUBLE
    COMPHRS  As Double         '  5  COMPHRS  AS DOUBLE

'022498 New for Wrightsville beech
    PerHRS   As Double         '  6
'*********************

    TOTHRSPD As Double         '  6  TOTHRSPD AS DOUBLE
    OTWORKED As Double         '  7  OTWORKED AS DOUBLE
    OTHRSPD  As Double         '  8  OTHRSPD  AS DOUBLE
    OT2Comp  As Double         '  9  OT2COMP  AS DOUBLE

    HDist(1 To 8)  As HourlyIDistType

    DTREGHRS As Double         '  34 TOTREGHR AS DOUBLE
    DTOTHRS  As Double         '  35 TOTOTHRS AS DOUBLE

    ALTEARN1 As Double         '  36 ALTEARN1 AS DOUBLE
    ALTEARN2 As Double         '  37 ALTEARN2 AS DOUBLE
    ALTEARN3 As Double         '  38 ALTEARN3 AS DOUBLE

    TaxFring  As Double

    AERNDST1 As String * 14    '  39 AERNDST1 AS STRING * 14
    AERNAMT1 As Double         '  40 AERNAMT1 AS DOUBLE
    AERNDST2 As String * 14    '  41 AERNDST2 AS STRING * 14
    AERNAMT2 As Double         '  42 AERNAMT2 AS DOUBLE
    AERNDST3 As String * 14    '  43 AERNDST3 AS STRING * 14
    AERNAMT3 As Double         '  44 AERNAMT3 AS DOUBLE
    AERNDST4 As String * 14    '  45 AERNDST4 AS STRING * 14
    AERNAMT4 As Double         '  46 AERNAMT4 AS DOUBLE
    AERNDST5 As String * 14    '  47 AERNDST5 AS STRING * 14
    AERNAMT5 As Double         '  48 AERNAMT5 AS DOUBLE
    AERNDST6 As String * 14    '  49 AERNDST6 AS STRING * 14
    AERNAMT6 As Double         '  50 AERNAMT6 AS DOUBLE
    TOTAERN  As Double         '  51 TOTAERN  AS DOUBLE
    TOTEADST As Double         '  52 TOTEADST AS DOUBLE
End Type

Type SalaryIDistType
    DAcct    As String * 14    '  10 WAGEDST1 AS STRING * 14
    DPct     As Double         '  11 D1REGHRS AS DOUBLE
End Type

Type SalaryInputType
    PAYSAL   As String * 1     '  1
    VACHRS   As Double         '  2
    SICKHRS  As Double         '  3
    HOLHRS   As Double         '  4
'022498 New for Wrightsville beech
    PerHRS   As Double         '  5
    COMPHRS  As Double 'added 9/1/04
    Hrs2Cmp  As Double 'added 9/1/04
'*********************************
    SDist(1 To 8)   As SalaryIDistType

    TOTOTHRS As Double         '  22

    ALTEARN1 As Double         '  23
    ALTEARN2 As Double         '  24
    ALTEARN3 As Double         '  25

    TaxFring As Double         '  26

    AERNDST1 As String * 14    '  27
    AERNAMT1 As Double         '  28
    AERNDST2 As String * 14    '  29
    AERNAMT2 As Double         '  30
    AERNDST3 As String * 14    '  31
    AERNAMT3 As Double         '  32
    AERNDST4 As String * 14    '  33
    AERNAMT4 As Double         '  34
    AERNDST5 As String * 14    '  35
    AERNAMT5 As Double         '  36
    AERNDST6 As String * 14    '  37
    AERNAMT6 As Double         '  38

    TOTAERN  As Double         '  39
    TOTEADST As Double         '  40
End Type

Type GLIFDataType12              'Hummm.
   TranAcct   As String * 12     ' 12 AS ARTRANACCT$
   TranDate   As String * 8      ' 8 AS artrandate$
   TranDesc   As String * 20     ' 19 AS artrandesc$
   CrAmt      As Double          ' 8 AS ARCRAMT$
   DrAmt      As Double          ' 8 AS ARDRAMT$
   Detail     As Double          ' 8 AS ardetail$
   Source     As String * 8      ' 8 AS arsource$
   rRNEXTTR   As Single          ' 4 AS ARNEXTTR$
   Fill       As String * 8
   FromFlag   As String * 1
End Type

Type SplitDedsType
   Acct       As String * 14
   FEDTAX     As Double
   STATAX     As Double
   MEDTAX     As Double
   SOCTAX     As Double
   RetTax     As Double
   EICPay     As Double
'   DedAmts(1 To 12) As Double
   DedAmts(1 To 50) As Double

End Type

Type GLIFDataType14
   TranAcct   As String * 14   'For New G/L
   TranDate   As String * 8
   TranDesc   As String * 20
   CrAmt      As Double
   DrAmt      As Double
   Detail     As Double
   Source     As String * 8
   rRNEXTTR   As Single
   Fill       As String * 6    'Adj. from 8 to 6
   FromFlag   As String * 1
End Type

Type FundType
   FundCode   As String * 14
   Credit     As Double
   Debit      As Double
   Net        As Double
End Type

Type AcctSumType
   FundCode   As String * 14
   Credit     As Double
   Debit      As Double
   'Net        AS DOUBLE
End Type

Type JGLAcctIdxType
    AcctNum As String * 14
    RecNo   As Integer
End Type

Type GLAcctIdxType
    AcctNum As Double
    RecNo   As Single
End Type

'Type GLSetupRecType                 'still under const.
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
'End Type

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


Type OSChkRecType
   ChkNum   As Single        '4 AS chknum$
   CHKDATE  As String * 8    '8 AS chkdate$
   Desc     As String * 30   '30 AS chkdesc$
   Amt      As Single        '4 AS chkamt$
   Src      As Integer       '2 AS CHKSOURCE$
   Cleared  As Integer       'added by JB
   BankCode As Integer       '09/07/96 Gate City's
   filler   As String * 12 '14   '16 AS nul$
End Type

'Type GLTransRecType                 'Transaction Record: 96 bytes
'   AcctRec     As Integer           'Pointer to Acct Record
'   AcctNum     As String * 14       'Formatted Acct Number string
'   TrDate      As Integer           'Date2Num function
'   Desc        As String * 20       'Transaction Description
'   Ref         As String * 8        'Document Reference
'   DrAmt       As Double            'Debit Amount
'   CrAmt       As Double            'Credit Amount
'   Src         As String * 8        'Module Source Code
'   NextTran    As Long              'Pointer to Next Trans
'   Res         As String * 20       'Reserved for future needs
'   Marked      As Integer
'End Type

Type TrEditRecType                  'Experimental GJ edit record:
   Deleted     As Integer           'Deleted transaction flag
   Posted      As Integer           'Posted flag
   AcctRec     As Integer           'Pointer to acct rec
   AcctNum     As String * 14       'Formatted Acct number string
   AcctName    As String * 30       'Account Title
   TrDate      As Integer           'Date2Num
   DrAmt       As Double            'Transaction Debit Amount
   CrAmt       As Double            'Transaction Credit Amount
   EType       As String * 1        'Entry Type (Debit/Credit)
   Desc        As String * 20       'Transaction Description
   Ref         As String * 8        'Document Reference #
   Src         As String * 8        'Module Source Code
   'Res         AS STRING *          'Reserve for future needs
End Type

Type TrSortType                     'Used for sorting trans in history rpt
   TrDate     As Integer            'Transaction Date
   Record     As Long               'Pointer to transaction record
End Type

Type TrSortType1                    'Used for sorting trans in history rpt
   TrDate     As String * 12             'Transaction Date
   Record     As Long               'Pointer to transaction record
End Type

Type IFRecType
   AcctNum As String * 9      '9 AS tranacct$
   TrDate As String * 8       '8 AS trandate$
   Desc As String * 20        '20 AS trandesc$
   CrAmt As Double            '8 AS cramt$
   DrAmt As Double            '8 AS dramt$
   Ref As String * 8          '8 AS detail$
   Src As String * 8          '8 AS source$
   filler As String * 14      '4 AS nexttr$
   Posted As Integer
End Type

Type GLFBAdjRecType
   AcctNum  As String * 16
   AdjAmt As Double
End Type

Type GLFundCloseRecType
   FundNum As String * 3
End Type

Type BankRecType   '128 bytes
   Deleted As Integer
   BankNum As Integer
   BankName As String * 25
   BankAcct As String * 25
   GLAcct As String * 25
   Pad As String * 49
End Type

Type GLSumSetupRecType
   Beg1  As String * 6
   End1  As String * 6
   Desc1 As String * 30
   Beg2  As String * 6
   End2  As String * 6
   Desc2 As String * 30
   Beg3  As String * 6
   End3  As String * 6
   Desc3 As String * 30
   Beg4  As String * 6
   End4  As String * 6
   Desc4 As String * 30
   Beg5  As String * 6
   End5  As String * 6
   Desc5 As String * 30
   Beg6  As String * 6
   End6  As String * 6
   Desc6 As String * 30
   Beg7  As String * 6
   End7  As String * 6
   Desc7 As String * 30
   Beg8  As String * 6
   End8  As String * 6
   Desc8 As String * 30
   Beg9  As String * 6
   End9  As String * 6
   Desc9 As String * 30
   Pad   As String * 75
End Type

Type InvTaxAcct
    AcctNo  As String * 16
    TaxAmt  As Double
End Type

Type GLFundRecType                  'Fund Record Type: 64 bytes
   Deleted     As Integer           'Deleted Flag
   FundNum     As String * 4        'Fund Code
   Title       As String * 30       'Fund Title
   Res         As String * 28       'Reserve for future needs
End Type

Type GLAcctIndexType                'Account Index: 16 bytes
   AcctNum     As String * 14       'Formatted account Number string
   RecNum      As Integer           'Pointer to record
End Type

Type GLFundIndexType                'Fund Index : 16 bytes
   FundNum     As String * 4        'Fund Number
   RecNum      As Integer           'Pointer to record
End Type

Type EarnNoMatchType
  EarnYN       As String * 1
  Pad         As String * 10
End Type

Type VoidDEDType
   DedGLNum       As String * 14
   DAmt        As Double
   DedDesc     As String * 12
End Type

Type VoidCheckType
  EmpNum      As String * 10  'employee number
  NumOfAccts  As Integer      'number of wage accounts for this employee
  CheckNum    As Long         'used to match up the correct check
  CheckDate   As Integer      'used to match up the correct check
  CheckAmt    As Double       'used to match up the correct check
  TransRec   As Double  'Pointer to transaction record for this check
  DedData(1 To 50) As VoidDEDType
  PRNet       As Double
  PRNetGL     As String * 14
  FEDWHAmt    As Double
  FEDWHGL     As String * 14
  STAWHAmt    As Double
  STAWHGL     As String * 14
  SOCWHAmt    As Double
  SOCWHGL     As String * 14
  SOCMATCRAmt As Double
  SOCMATCRGL  As String * 14
  SOCMATDBAmt As Double
  SOCMATDBGL  As String * 14
  MEDWHAmt    As Double
  MEDWHGL     As String * 14
  MEDMATCRAmt As Double
  MEDMATCRGL  As String * 14
  MEDMATDBAmt As Double
  MEDMATDBGL  As String * 14
  RETWHAmt    As Double
  RETWHGL     As String * 14
  RETMATCRAmt As Double
  RETMATCRGL  As String * 14
  RETMATDBAmt As Double
  RETMATDBGL  As String * 14
  WagesAmt    As Double
  WagesGL     As String * 14
  VoidFlag    As Boolean 'False = has not been voided
  PPEAmt      As Double
  PPEGL       As String * 14
  PPETotAmt   As Double
  PPETotGL    As String * 14
  Type        As String * 1 '(C)entral Depository, (P)ool, (I)mprest
  Pad         As String * 40
End Type

Type CitiPassTempType
  usernum   As Integer
  UserName  As String * 15
  frommdl   As Integer   'this is to indicate to citipak ok to have file
End Type

Type PrintOptType
  PrintOpt  As Integer
End Type

