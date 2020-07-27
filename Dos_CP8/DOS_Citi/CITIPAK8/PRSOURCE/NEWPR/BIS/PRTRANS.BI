'npayroll
'greetings from the Klingon Empire.

TYPE TransWageDistType                      'Transaction Wage Distributions
    DAcct       AS STRING * 14              'G/L Account (Dept)
    DRHrs       AS DOUBLE                   'Reg Hours Distributed
    DOHrs       AS DOUBLE                   'OT Hours Distributed
    DPct        AS DOUBLE                   'Distribution Percent
    DRWage      AS DOUBLE                   'Reg Wage Distributed
    DOWage      AS DOUBLE                   'OT Wage Distributed
END TYPE

TYPE TransEarnDistType                      'Additional Earnings
    EAcct       AS STRING * 14              'Default Add'l Earings Distribution Account (G/L)
    EAmt        AS DOUBLE                   'Default Add'l Earings Amount
END TYPE

TYPE TransRecType
    TActive         AS INTEGER             '1     'Active Transaction Flag
    PrevTransRec    AS INTEGER             '2     'Pointer to employee's prev trans
    EmpPin          AS INTEGER             '3     'Pointer to employee rec
    PaySFlag        AS STRING * 1          '4     'Pay Salary Flag in time trans
    CHECKNUM        AS LONG                '5     'Payroll Check Number
    PayPdStart      AS INTEGER             '6     'Start of Pay Period
    PayPdEnd        AS INTEGER             '7     'End of Pay Period
    CheckDate       AS INTEGER   'yea      '8     'Date checks written
    PostDate        AS INTEGER             '9     'Date Transaction are posted
    PayType         AS STRING * 1          '10    'Salaried or Hourly
    BaseRate        AS DOUBLE              '11    'Base Rate or Salary Amt
    OTRate          AS DOUBLE              '12    'Overtime Rate
    RegHrsWork       AS DOUBLE             '13    'Hours worked this period
    VacUsed          AS DOUBLE             '14    'vacation used this period
    SickUsed         AS DOUBLE             '15    'Sick hours used this period
    CompUsed         AS DOUBLE             '16    'comp hours used this period

    HOLHOURS         AS DOUBLE             '17    'holiday hours used this period
    PerHours         AS DOUBLE             '17a

    RegHrsPaid       AS DOUBLE             '18    'sum of reg hours paid
    OTHours          AS DOUBLE             '19    'OT hours this period
    OTHrsPaid        AS DOUBLE             '20    'OT hours paid this period
    OT2Comp          AS DOUBLE             '21    'Hours to comp time
    TDist(1 TO 8)    AS TransWageDistType  '      'Above TransWageDistType                              'wage distributions
                              '22 23 24 25 26 27 28 29
    TotRegWage     AS DOUBLE               '30    'Total Reg Wage distributions
    TotOTWage      AS DOUBLE               '31    'Total OT Wage distributions
    GrossWage      AS DOUBLE               '32    'Reg Wage + OT Wage
    EAmt(1 TO 3) AS DOUBLE                 '      'Add Earnings amounts
                                     '33 34 35
    EDist(1 TO 6)   AS TransEarnDistType   '2     'Add Earnings distribitions (G/L) accs
    TotAdditEarn    AS DOUBLE              '      'Total Additional Earnings
    GROSSPAY        AS DOUBLE              '      'Add Earnings + GrossWage
    SocGrossPay     AS DOUBLE              '      'Social Security Gross
    MedGrossPay     AS DOUBLE              '      'Medicare Gross
    FedGrossPay     AS DOUBLE              '      'Federal Gross
    StaGrossPay     AS DOUBLE              '      'State Gross
    SocTaxAmt       AS DOUBLE              '      'Social Security Tax W/H
    MedTaxAmt       AS DOUBLE              '      'Medicare Tax W/H
    FedTaxAmt       AS DOUBLE              '      'Fed Tax W/H
    StaTaxAmt       AS DOUBLE              '      'State Tax W/H
    TotTaxAmt       AS DOUBLE              '      'Total Taxes W/H
    RetireAmt       AS DOUBLE              '      'Retirement W/H
    DAMT(1 TO 12) AS DOUBLE                '      'Voluntary Deduction amounts / pcts
    TotDedAmt AS DOUBLE                    '      'Total Voluntary Deductions
    EICAmt     AS DOUBLE                   '
    NETPAY    AS DOUBLE                    '
    PeriodHistRec AS INTEGER               '      'YTD Totals?
    MatchRetAmt     AS DOUBLE              '      'Employer's Retirement Match
    MatchSocAmt     AS DOUBLE              '      'Employer's Social Secity Match
    MatchMedAmt     AS DOUBLE              '      'Employer's Medicare Match
    RetGrossPay     AS DOUBLE              '      'Retirement Gross
    TaxFring        AS DOUBLE              '      'Taxable Fringe
    Pad1            AS STRING * 38
'-----------------------------
'can sum these for a report from transaction history file
END TYPE

TYPE ManualTransRecType
    PDSTART  AS INTEGER           ' 1
    PDEND    AS INTEGER           '
    CHKDATE  AS INTEGER           '
    CHECKNUM AS LONG           '
    REGHRS   AS DOUBLE            ' 5
    SICKHRS  AS DOUBLE            '  6
    VACHRS   AS DOUBLE             ' 7
    COMPHRS  AS DOUBLE             ' 8
    HOLHOURS AS DOUBLE             ' 9
    OTHRSPD  AS DOUBLE             ' 10

    DISTACT1 AS STRING * 14         '11
    WAGEAMT1 AS DOUBLE              '12
    DISTACT2 AS STRING * 14         '13
    WAGEAMT2 AS DOUBLE              '14
    DISTACT3 AS STRING * 14         '15
    WAGEAMT3 AS DOUBLE              '16
    DISTACT4 AS STRING * 14         '17
    WAGEAMT4 AS DOUBLE              '18

    GROSSPAY AS DOUBLE              '19
    REGWAGE  AS DOUBLE              '20
    OTWAGE   AS DOUBLE              '21
    FEDTAX   AS DOUBLE              '22
    STATAX   AS DOUBLE              '23
    SOCTAX   AS DOUBLE              '24
    MEDTAX   AS DOUBLE              '25
    RETAMT   AS DOUBLE              '26

    DAMT(1 TO 12)  AS DOUBLE        '27-38

    TOTDED   AS DOUBLE              '39
    EIC      AS DOUBLE              '40
    NETPAY   AS DOUBLE              '41
    FEDGROSS AS DOUBLE              '42
    STAGROSS AS DOUBLE              '43
    SOCGROSS AS DOUBLE              '44
    MEDGROSS AS DOUBLE              '45
    RETGROSS AS DOUBLE              '46
    TOTTAX   AS DOUBLE              '47  this is hidden and protected
END TYPE

TYPE DistWageRptType                        'For the register?
    Acct       AS STRING * 14
    Pct        AS DOUBLE
    RHrs       AS DOUBLE
    OHrs       AS DOUBLE
    RWage      AS DOUBLE
    OWage      AS DOUBLE
    AddEarn    AS DOUBLE
    GROSSPAY   AS DOUBLE
    MATSocAmt  AS DOUBLE
    MATMedAmt  AS DOUBLE
    MATRetAmt  AS DOUBLE
END TYPE

