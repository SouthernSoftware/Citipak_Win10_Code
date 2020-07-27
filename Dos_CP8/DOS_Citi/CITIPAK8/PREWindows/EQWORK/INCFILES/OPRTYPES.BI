TYPE nTransRecType
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


TYPE nEmpData2Type         'new emp 2 rec
'------ EMPMA1
    EMPNO    AS STRING * 10
    EMPSSN   AS STRING * 11
    EMPLNAME AS STRING * 24
    EMPFNAME AS STRING * 24
    EMPADDR1 AS STRING * 36
    EMPADDR2 AS STRING * 36
    EMPCITY  AS STRING * 24
    EMPSTATE AS STRING * 2
    EMPZIP   AS STRING * 10
    EMPBDAY  AS INTEGER
    EMPGENDR AS STRING * 6
    EMPRACE  AS STRING * 14
    EMPRETNO AS STRING * 16
    EMPRETTP AS STRING * 24
'new
    DRAFTCOD AS STRING * 1
    EMPDDACC AS STRING * 20
    PRENOTED AS STRING * 1
    BANKNAME AS STRING * 33
    BANKLOC  AS STRING * 30
    TRANSIT  AS STRING * 9
'new
    EMPJOB   AS STRING * 28
    EMPWCCLS AS STRING * 12
    EMPSTATS AS STRING * 10
    EMPBCODE AS DOUBLE
    EMPPTYPE AS STRING * 10
    EMPPFREQ AS STRING * 16
    EMPPRATE AS DOUBLE
    EMPORATE AS DOUBLE
    EMPHDATE AS INTEGER
    EMPRDATE AS INTEGER
    EMPTDATE AS INTEGER
'------ EMPMA2
    EMPFEDX  AS STRING * 1
    EMPFEDO2 AS STRING * 1
    EMPFEDO1 AS DOUBLE
    EMPFEDS  AS STRING * 1
    EMPFEDA  AS INTEGER       'num of allowance
    EMPFEDAA AS DOUBLE
    EMPSTAX  AS STRING * 1
    EMPSTAO2 AS STRING * 1
    EMPSTAO1 AS DOUBLE
    EMPSTAS  AS STRING * 1
    EMPSTAA  AS INTEGER       '
    EMPSTAAA AS DOUBLE
    EMPSOCX  AS STRING * 1
    EMPMEDX  AS STRING * 1
    EMPEIC   AS STRING * 1

    EmpDed(1 TO 12)  AS EmpDedType

'------ page 3
    EMPEACT1 AS STRING * 14
    EMPEAMT1 AS DOUBLE
    EMPEACT2 AS STRING * 14
    EMPEAMT2 AS DOUBLE
    EMPEACT3 AS STRING * 14
    EMPEAMT3 AS DOUBLE

    EMPHP    AS STRING * 1

    EDist(1 TO 8)   AS EmpWageDistType

    EMPVACE  AS DOUBLE
    EMPVUSED AS DOUBLE
    EMPVBAL  AS DOUBLE

    EMPSLE   AS DOUBLE
    EMPSLUSE AS DOUBLE
    EMPSLBAL AS DOUBLE

    EMPCTE   AS DOUBLE
    EMPCTUSE AS DOUBLE
    EMPCTBAL AS DOUBLE

    PERERN   AS DOUBLE
    PERUSED  AS DOUBLE
    PERBAL   AS DOUBLE

    HOLERN   AS DOUBLE
    HOLUSED  AS DOUBLE
    HOLBAL   AS DOUBLE

    LeaveTbl AS INTEGER

    ExcludeESC  AS STRING * 1
    UseLife     AS STRING * 1

'------ Misc
    LastTransRec AS INTEGER
    EmpPin       AS INTEGER
    Deleted      AS INTEGER
    'for new leave table stuff
' PreNoteFlag     AS INTEGER
    Unused       AS STRING * 51
END TYPE



TYPE oEmpData2Type1a           'old type 1
'------ EMPMA1
    EMPNO    AS STRING * 10
    EMPSSN   AS STRING * 11
    EMPLNAME AS STRING * 24
    EMPFNAME AS STRING * 24
    EMPADDR1 AS STRING * 36
    EMPADDR2 AS STRING * 36
    EMPCITY  AS STRING * 24
    EMPSTATE AS STRING * 2
    EMPZIP   AS STRING * 10
    EMPBDAY  AS INTEGER
    EMPGENDR AS STRING * 6
    EMPRACE  AS STRING * 14
    EMPRETNO AS STRING * 16
    EMPRETTP AS STRING * 24
    EMPDDACC AS STRING * 30
    EMPJOB   AS STRING * 28
    EMPWCCLS AS STRING * 12
    EMPSTATS AS STRING * 10
    EMPBCODE AS DOUBLE
    EMPPTYPE AS STRING * 10
    EMPPFREQ AS STRING * 16
    EMPPRATE AS DOUBLE
    EMPORATE AS DOUBLE
    EMPHDATE AS INTEGER
    EMPRDATE AS INTEGER
    EMPTDATE AS INTEGER
'------ EMPMA2
    EMPFEDX  AS STRING * 1
    EMPFEDO2 AS STRING * 1
    EMPFEDO1 AS DOUBLE
    EMPFEDS  AS STRING * 1
    EMPFEDA  AS INTEGER       'num of allowance
    EMPFEDAA AS DOUBLE
    EMPSTAX  AS STRING * 1
    EMPSTAO2 AS STRING * 1
    EMPSTAO1 AS DOUBLE
    EMPSTAS  AS STRING * 1
    EMPSTAA  AS INTEGER       '
    EMPSTAAA AS DOUBLE
    EMPSOCX  AS STRING * 1
    EMPMEDX  AS STRING * 1
    EMPEIC   AS STRING * 1

    EmpDed(1 TO 12)  AS EmpDedType

'------ page 3
    EMPEACT1 AS STRING * 14
    EMPEAMT1 AS DOUBLE
    EMPEACT2 AS STRING * 14
    EMPEAMT2 AS DOUBLE
    EMPEACT3 AS STRING * 14
    EMPEAMT3 AS DOUBLE

    EMPHP    AS STRING * 1

    EDist(1 TO 8)   AS EmpWageDistType

    EMPVACE  AS DOUBLE
    EMPVUSED AS DOUBLE
    EMPVBAL  AS DOUBLE

    EMPSLE   AS DOUBLE
    EMPSLUSE AS DOUBLE
    EMPSLBAL AS DOUBLE

    EMPCTE   AS DOUBLE
    EMPCTUSE AS DOUBLE
    EMPCTBAL AS DOUBLE
    LeaveTbl AS INTEGER
'------ EMPMA4
    EMPNOTES AS INTEGER
'------ Misc
    LastTransRec AS INTEGER
    EmpPin   AS INTEGER
    Deleted  AS INTEGER
    'for new leave table stuff
    Unused   AS STRING * 26

END TYPE

TYPE oEmpData2Type2     'mowasa
'------ EMPMA1
    EMPNO    AS STRING * 10
    EMPSSN   AS STRING * 11
    EMPLNAME AS STRING * 24
    EMPFNAME AS STRING * 24
    EMPADDR1 AS STRING * 36
    EMPADDR2 AS STRING * 36
    EMPCITY  AS STRING * 24
    EMPSTATE AS STRING * 2
    EMPZIP   AS STRING * 10
    EMPBDAY  AS INTEGER
    EMPGENDR AS STRING * 6
    EMPRACE  AS STRING * 14
    EMPRETNO AS STRING * 16
    EMPRETTP AS STRING * 24
'new
    DRAFTCOD AS STRING * 1
    EMPDDACC AS STRING * 20
    BANKNAME AS STRING * 34
    BANKLOC  AS STRING * 30
    TRANSIT  AS STRING * 9
'new
    EMPJOB   AS STRING * 28
    EMPWCCLS AS STRING * 12
    EMPSTATS AS STRING * 10
    EMPBCODE AS DOUBLE
    EMPPTYPE AS STRING * 10
    EMPPFREQ AS STRING * 16
    EMPPRATE AS DOUBLE
    EMPORATE AS DOUBLE
    EMPHDATE AS INTEGER
    EMPRDATE AS INTEGER
    EMPTDATE AS INTEGER
'------ EMPMA2
    EMPFEDX  AS STRING * 1
    EMPFEDO2 AS STRING * 1
    EMPFEDO1 AS DOUBLE
    EMPFEDS  AS STRING * 1
    EMPFEDA  AS INTEGER       'num of allowance
    EMPFEDAA AS DOUBLE
    EMPSTAX  AS STRING * 1
    EMPSTAO2 AS STRING * 1
    EMPSTAO1 AS DOUBLE
    EMPSTAS  AS STRING * 1
    EMPSTAA  AS INTEGER       '
    EMPSTAAA AS DOUBLE
    EMPSOCX  AS STRING * 1
    EMPMEDX  AS STRING * 1
    EMPEIC   AS STRING * 1

    EmpDed(1 TO 12)  AS EmpDedType

'------ page 3
    EMPEACT1 AS STRING * 14
    EMPEAMT1 AS DOUBLE
    EMPEACT2 AS STRING * 14
    EMPEAMT2 AS DOUBLE
    EMPEACT3 AS STRING * 14
    EMPEAMT3 AS DOUBLE

    EMPHP    AS STRING * 1

    EDist(1 TO 8)   AS EmpWageDistType

    EMPVACE  AS DOUBLE
    EMPVUSED AS DOUBLE
    EMPVBAL  AS DOUBLE

    EMPSLE   AS DOUBLE
    EMPSLUSE AS DOUBLE
    EMPSLBAL AS DOUBLE

    EMPCTE   AS DOUBLE
    EMPCTUSE AS DOUBLE
    EMPCTBAL AS DOUBLE
    LeaveTbl AS INTEGER

    UseHosp  AS STRING * 1
    UseLife  AS STRING * 1

    LastTransRec AS INTEGER
    EmpPin   AS INTEGER
    Deleted  AS INTEGER
    'for new leave table stuff
 PreNoteFlag AS INTEGER
    Unused   AS STRING * 24
END TYPE

