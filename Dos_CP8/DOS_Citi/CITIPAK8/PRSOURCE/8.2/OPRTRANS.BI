'greetings from the Klingon Empire.

TYPE TransRecType1
    TActive         AS INTEGER              'Active Transaction Flag
    PrevTransRec    AS INTEGER              'Pointer to employee's prev trans
    EmpPin          AS INTEGER              'Pointer to employee rec
    
    PaySFlag        AS STRING * 1           'Pay Salary Flag in time trans

    CHECKNUM        AS LONG                 'Payroll Check Number
    PayPdStart      AS INTEGER              'Start of Pay Period
    PayPdEnd        AS INTEGER              'End of Pay Period
    CheckDate       AS INTEGER   'yea       'Date checks written
    PostDate        AS INTEGER              'Date Transaction are posted
    
    PayType         AS STRING * 1           'Salaried or Hourly
    BaseRate        AS DOUBLE               'Base Rate or Salary Amt
    OTRate          AS DOUBLE               'Overtime Rate
                                            
    RegHrsWork       AS DOUBLE              'Hours worked this period
    VacUsed          AS DOUBLE              'vacation used this period
    SickUsed         AS DOUBLE              'Sick hours used this period
    CompUsed         AS DOUBLE              'comp hours used this period
    HOLHOURS         AS DOUBLE              'holiday hours used this period
    RegHrsPaid       AS DOUBLE              'sum of reg hours paid

    OTHours          AS DOUBLE              'OT hours this period
    OTHrsPaid        AS DOUBLE              'OT hours paid this period
    OT2Comp          AS DOUBLE              'Hours to comp time

    TDist(1 TO 8)    AS TransWageDistType   'Above TransWageDistType                                    'wage distributions
    
    TotRegWage     AS DOUBLE                'Total Reg Wage distributions
    TotOTWage      AS DOUBLE                'Total OT Wage distributions

    GrossWage      AS DOUBLE                'Reg Wage + OT Wage

    EAmt(1 TO 3) AS DOUBLE                  'Add Earnings amounts

    EDist(1 TO 6)   AS TransEarnDistType    'Add Earnings distribitions (G/L) accts

    TotAdditEarn    AS DOUBLE               'Total Additional Earnings

    GROSSPAY        AS DOUBLE               'Add Earnings + GrossWage

    SocGrossPay     AS DOUBLE               'Social Security Gross
    MedGrossPay     AS DOUBLE               'Medicare Gross
    FedGrossPay     AS DOUBLE               'Federal Gross
    StaGrossPay     AS DOUBLE               'State Gross

    SocTaxAmt       AS DOUBLE               'Social Security Tax W/H
    MedTaxAmt       AS DOUBLE               'Medicare Tax W/H
    FedTaxAmt       AS DOUBLE               'Fed Tax W/H
    StaTaxAmt       AS DOUBLE               'State Tax W/H
    TotTaxAmt       AS DOUBLE               'Total Taxes W/H
    RetireAmt       AS DOUBLE               'Retirement W/H

    DAMT(1 TO 12) AS DOUBLE                 'Voluntary Deduction amounts / pcts

    TotDedAmt AS DOUBLE                     'Total Voluntary Deductions
    
    NETPAY    AS DOUBLE

    PeriodHistRec AS INTEGER                'YTD Totals?

    MatchRetAmt     AS DOUBLE               'Employer's Retirement Match
    MatchSocAmt     AS DOUBLE               'Employer's Social Secity Match
    MatchMedAmt     AS DOUBLE               'Employer's Medicare Match
'-----------------------------
'can sum these for a report from transaction history file

'    EICAmt     AS DOUBLE
END TYPE
'****

TYPE TransRecType2
    TActive         AS INTEGER              'Active Transaction Flag
    PrevTransRec    AS INTEGER              'Pointer to employee's prev trans
    EmpPin          AS INTEGER              'Pointer to employee rec

    PaySFlag        AS STRING * 1           'Pay Salary Flag in time trans

    CHECKNUM        AS LONG                 'Payroll Check Number
    PayPdStart      AS INTEGER              'Start of Pay Period
    PayPdEnd        AS INTEGER              'End of Pay Period
    CheckDate       AS INTEGER   'yea       'Date checks written
    PostDate        AS INTEGER              'Date Transaction are posted

    PayType         AS STRING * 1           'Salaried or Hourly
    BaseRate        AS DOUBLE               'Base Rate or Salary Amt
    OTRate          AS DOUBLE               'Overtime Rate

    RegHrsWork       AS DOUBLE              'Hours worked this period
    VacUsed          AS DOUBLE              'vacation used this period
    SickUsed         AS DOUBLE              'Sick hours used this period
    CompUsed         AS DOUBLE              'comp hours used this period
    HOLHOURS         AS DOUBLE              'holiday hours used this period
    RegHrsPaid       AS DOUBLE              'sum of reg hours paid

    OTHours          AS DOUBLE              'OT hours this period
    OTHrsPaid        AS DOUBLE              'OT hours paid this period
    OT2Comp          AS DOUBLE              'Hours to comp time

    TDist(1 TO 8)    AS TransWageDistType   'Above TransWageDistType                                    'wage distributions

    TotRegWage     AS DOUBLE                'Total Reg Wage distributions
    TotOTWage      AS DOUBLE                'Total OT Wage distributions

    GrossWage      AS DOUBLE                'Reg Wage + OT Wage

    EAmt(1 TO 3) AS DOUBLE                  'Add Earnings amounts

    EDist(1 TO 6)   AS TransEarnDistType    'Add Earnings distribitions (G/L) accts

    TotAdditEarn    AS DOUBLE               'Total Additional Earnings

    GROSSPAY        AS DOUBLE               'Add Earnings + GrossWage

    SocGrossPay     AS DOUBLE               'Social Security Gross
    MedGrossPay     AS DOUBLE               'Medicare Gross
    FedGrossPay     AS DOUBLE               'Federal Gross
    StaGrossPay     AS DOUBLE               'State Gross

    SocTaxAmt       AS DOUBLE               'Social Security Tax W/H
    MedTaxAmt       AS DOUBLE               'Medicare Tax W/H
    FedTaxAmt       AS DOUBLE               'Fed Tax W/H
    StaTaxAmt       AS DOUBLE               'State Tax W/H
    TotTaxAmt       AS DOUBLE               'Total Taxes W/H
    RetireAmt       AS DOUBLE               'Retirement W/H

    DAMT(1 TO 12) AS DOUBLE                 'Voluntary Deduction amounts / pcts

    TotDedAmt AS DOUBLE                     'Total Voluntary Deductions

    EICAmt     AS DOUBLE

    NETPAY    AS DOUBLE

    PeriodHistRec AS INTEGER                'YTD Totals?

    MatchRetAmt     AS DOUBLE               'Employer's Retirement Match
    MatchSocAmt     AS DOUBLE               'Employer's Social Secity Match
    MatchMedAmt     AS DOUBLE               'Employer's Medicare Match
'-----------------------------
'can sum these for a report from transaction history file

END TYPE

