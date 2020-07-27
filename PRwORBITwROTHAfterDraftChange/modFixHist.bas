Attribute VB_Name = "Module1"
DefInt A-Z

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
    Less401k(1 To 3)  As Boolean
    Voided          As String * 1 'added 12/17/04
    VoidRec         As Long
    Pad1            As String * 3
'-----------------------------
'can sum these for a report from transaction history file
End Type

Public Function Date2Num%(TheDate$)
 'useful function throughout program...
 'takes a string date and converts into a number based on 12/31/1979
  Date2Num% = DateDiff("d", "12/31/1979", (TheDate$))
End Function

Public Function MakeRegDate(ByVal DateNumb) As String
  Dim Month As Integer, ThisDate As String
  'function does the opposite of Date2Num
  If DateNumb = -32767 Then
    MakeRegDate = "%%%%%%%%%% "
  Else
    MakeRegDate = Format(DateAdd("d", (DateNumb), "12-31-1979"), "mm/dd/yyyy")
  End If
End Function

