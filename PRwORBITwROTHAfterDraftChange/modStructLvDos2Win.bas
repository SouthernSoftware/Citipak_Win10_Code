Attribute VB_Name = "modStructLvWin2Win"
Option Explicit

Type LeaveEntryType
  YEARS   As Integer
  EARN    As Double
End Type

Type OldLeaveRecType
  VacMax   As Double
  VEntry(1 To 20)  As LeaveEntryType
  SICKMAX  As Double
  SEntry(1 To 20)  As LeaveEntryType
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
    Less401k        As Boolean
    Pad1            As String * 12
'-----------------------------
'can sum these for a report from transaction history file
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

End Type

