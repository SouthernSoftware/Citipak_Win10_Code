

TYPE PeriodDefaultRecType
    PACTIVE  AS INTEGER
    PERBEG   AS INTEGER
    PEREND   AS INTEGER
    USEDEF   AS STRING * 1

    PAYWK    AS STRING * 1
    PAYBIWK  AS STRING * 1
    PAYSEMIM AS STRING * 1
    PAYMO    AS STRING * 1
    PAYQTR   AS STRING * 1
    PAYSEMIA AS STRING * 1
    PAYANNL  AS STRING * 1

    UseDed(1 TO 50)   AS STRING * 1
    USEAE1   AS STRING * 1
    USEAE2   AS STRING * 1
    USEAE3   AS STRING * 1
    MACTIVE  AS INTEGER
END TYPE

TYPE UnitFileRecType
    UFEMPR   AS STRING * 35
    UFATTN   AS STRING * 35
    UFADDR1  AS STRING * 35
    UFADDR2  AS STRING * 35
    UFCITY   AS STRING * 25
    UFSTATE  AS STRING * 2
    UFZIP    AS STRING * 10
    UFFEDID  AS STRING * 14
    UFSTAID  AS STRING * 14
    UFRETID  AS STRING * 14
    UFRETIDL AS STRING * 14
    ESCRTYPE AS INTEGER
    TAXWBASE AS DOUBLE
    BBTCNTNO AS STRING * 12
    BBTBATCH AS STRING * 12
    USEACH   AS STRING * 1
    IMMDNUM  AS STRING * 9
    IMMONUM  AS STRING * 9
    DBANK    AS STRING * 23
    OBANK    AS STRING * 23
'    FileVer  As Double 'commented out 8/28/03 because of Wrightsville Beach
    FileVer AS STRING * 7
    LMT401YN AS STRING * 1
    BankDraft AS STRING * 1
    '********added 11/11/02
    ESCRemitNum AS STRING * 20
    ESCEmplrNum AS STRING * 20
    '********added 8/31/04
    GMatch401K  AS DOUBLE
    LMatch401K  AS DOUBLE
    SSNOnCheck  AS STRING * 1

END TYPE

TYPE TaxRetLiabType
   Acct AS STRING * 14
END TYPE

TYPE RegDSysFileRecType
    USEIMP   AS STRING * 1
    CashAcct AS STRING * 14
    IDRACCT  AS STRING * 14
    ICRACCT  AS STRING * 14
    Liab(1 TO 5) AS TaxRetLiabType
    CITIDIR  AS STRING * 48
    SplitFlag AS STRING * 1
    EXPMETHD AS STRING * 1
    FRNGRATE AS DOUBLE
    FRNGEXP  AS STRING * 7
    FRNGDR   AS STRING * 14
    FRNGCR   AS STRING * 14
    INDRATE  AS DOUBLE
    INDEXP   AS STRING * 7
    INDDR    AS STRING * 14
    INDCR    AS STRING * 14
    SOCEXP   AS STRING * 14
    SOCLIAB  AS STRING * 14
    MEDEXP   AS STRING * 14
    MEDLIAB  AS STRING * 14
    RETEXP   AS STRING * 14
    RETLIAB  AS STRING * 14
    AcctCnt  AS INTEGER
    GLActLen AS INTEGER
    CheckStyle AS INTEGER
    GLCheckYN AS STRING * 1
'    VAC2SICK As String * 1
END TYPE

'two records in file
TYPE EIC1RecType
    EIC1OVR0 AS DOUBLE
    EIC1NVR0 AS DOUBLE
    EIC1AMT0 AS DOUBLE
    EIC1OVR1 AS DOUBLE
    EIC1NVR1 AS DOUBLE
    EIC1AMT1 AS DOUBLE
    EIC1OVR2 AS DOUBLE
    EIC1NVR2 AS DOUBLE
    EIC1AMT2 AS DOUBLE
    EIC1LESS AS DOUBLE
    EIC1EXES AS DOUBLE
END TYPE

TYPE EICRecType
    EIC(1 TO 2) AS EIC1RecType
END TYPE
'---------------------------------

TYPE RetireRecType
    TYPEDES1 AS STRING * 20
    TYPEWH1  AS DOUBLE
    TYPEM1   AS DOUBLE
    TYPEOT1  AS STRING * 1
    TYPETD1  AS STRING * 1
END TYPE
'---------------------------------

TYPE PRPPDraftInfoType
    EmpRec     AS LONG
    DraftDate  AS INTEGER
    NetPay     AS DOUBLE
END TYPE

TYPE PRDraftRecType
    BANKDEST AS STRING * 9
    BANKORIG AS STRING * 9
    BankName AS STRING * 23
    BANKLOC  AS STRING * 23
    FEDPREFX AS STRING * 1
    FEDID AS STRING * 9
END TYPE

TYPE DraftInfoFileName
    BANKDEST AS STRING * 9
    BANKORIG AS STRING * 9
    BankName AS STRING * 23
    BANKLOC  AS STRING * 23
    FEDPREFX AS STRING * 1
    FEDID AS STRING * 9
END TYPE

TYPE UBDraftPayRecType
    CustAcctNum   AS LONG
    DraftAmt      AS DOUBLE
END TYPE

TYPE ErnCodeRecType
    ERNCODE1 AS STRING * 10
    ERNFWT1  AS STRING * 1
    ERNSWT1  AS STRING * 1
    ERNSOC1  AS STRING * 1
    ERNMED1  AS STRING * 1
    ERNRET1  AS STRING * 1
    EarnYN   AS STRING * 1
    Pad      AS STRING * 10
END TYPE

TYPE DedCodeRecType
    DCDESC1  AS STRING * 10
    DCACCT1  AS STRING * 14
    DCFWT1   AS STRING * 1
    DCSWT1   AS STRING * 1
    DCSOC1   AS STRING * 1
    DCMED1   AS STRING * 1
END TYPE

TYPE StateTaxRecType
'--- Single
    TAX101 AS DOUBLE
    TAX102 AS DOUBLE
    TAX103 AS DOUBLE
    TAX104 AS DOUBLE
    TAX105 AS DOUBLE
    TAX106 AS DOUBLE
    TAX107 AS DOUBLE
    STS(1 TO 3, 1 TO 12) AS DOUBLE

'--- Married

    TAX201 AS DOUBLE
    TAX202 AS DOUBLE
    TAX203 AS DOUBLE
    TAX204 AS DOUBLE
    TAX205 AS DOUBLE
    TAX206 AS DOUBLE
    TAX207 AS DOUBLE
    STM(1 TO 3, 1 TO 12) AS DOUBLE

'--- Head of House

    TAX301 AS DOUBLE
    TAX302 AS DOUBLE
    TAX303 AS DOUBLE
    TAX304 AS DOUBLE
    TAX305 AS DOUBLE
    TAX306 AS DOUBLE
    TAX307 AS DOUBLE
    STH(1 TO 3, 1 TO 12) AS DOUBLE

'--- Georgia Page 4

    TAX401 AS DOUBLE
    TAX402 AS DOUBLE
    TAX403 AS DOUBLE
    TAX404 AS DOUBLE
    TAX405 AS DOUBLE
    TAX406 AS DOUBLE
    TAX407 AS DOUBLE
    ST4(1 TO 3, 1 TO 12) AS DOUBLE

END TYPE

TYPE FederalTaxRecType
'--- single
    FTSEMPSS AS DOUBLE
    FTSEMRSS AS DOUBLE
    FTSSSMW  AS DOUBLE
    FTSEMPM  AS DOUBLE
    FTSEMRM  AS DOUBLE
    FTSMMW   AS DOUBLE
    FTSSDAA  AS DOUBLE

    FTS(1 TO 3, 1 TO 10) AS DOUBLE

'--- married
    FTMEMPSS AS DOUBLE
    FTMEMRSS AS DOUBLE
    FTMSSMW  AS DOUBLE
    FTMEMPM  AS DOUBLE
    FTMEMRM  AS DOUBLE
    FTMMMW   AS DOUBLE
    FTMSDAA  AS DOUBLE

    FTM(1 TO 3, 1 TO 10) AS DOUBLE
END TYPE

'npayroll
TYPE EmployeeIndexType
    DataRecNum    AS INTEGER
END TYPE

TYPE EmpNumType
    EmpNum        AS STRING * 10
END TYPE

TYPE EmpDedType
    DPct        AS STRING * 7
    DAmt        AS DOUBLE
    DOTI        AS STRING * 1
END TYPE

TYPE EmpWageDistType
    DAcct       AS STRING * 14
    DAmt        AS DOUBLE
END TYPE

TYPE EmpData1Type
    EmpNo       AS STRING * 10     '1-10
    EmpLName    AS STRING * 24     '11-34
    EmpFName    AS STRING * 24     '35-58
    Data1RecNum AS INTEGER         '59-60
    TransRecNum AS INTEGER         '61-62
    Deleted     AS INTEGER         '63-64
    'KillFlag    AS STRING * 1
    'FillerPad   AS STRING * 1
END TYPE

TYPE EmpData2Type         'new emp 2 rec
    EmpNo    AS STRING * 10
    EmpSSN   AS STRING * 11
    EmpLName AS STRING * 24
    EmpFName AS STRING * 24
    EmpAddr1 AS STRING * 36
    EMPADDR2 AS STRING * 36
    EmpCity  AS STRING * 24
    EmpState AS STRING * 2
    EmpZip   AS STRING * 10
    EMPBDAY  AS INTEGER
    EMPGENDR AS STRING * 6
    EMPRACE  AS STRING * 14
    EMPRETNO AS STRING * 16
    EMPRETTP AS STRING * 24
'new
    DRAFTCOD AS STRING * 1
    EMPDDACC AS STRING * 20
    PRENOTED AS STRING * 1
    BankName AS STRING * 33
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

    EmpDed(1 TO 50)  AS EmpDedType

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
    PerUsed  AS DOUBLE
    PERBAL   AS DOUBLE

    HOLERN   AS DOUBLE
    HolUsed  AS DOUBLE
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

'    Unused       AS STRING * 51

    LDTDate      AS INTEGER      'last test date
    CDTDate      AS INTEGER      'current test date
    InprocFlag   AS INTEGER      'in process flag

    Unused       AS STRING * 43
    CheckType    AS INTEGER
    '*******added 11/11/02
    YN401K   AS STRING * 1
    PrimeDept AS STRING * 6
    HomePhone AS STRING * 14
    EmrgncyCntctName AS STRING * 48
    EmrgncyCntctPhnNum AS STRING * 14
    EmrgncyCntctRelation AS STRING * 16
    '*******added 8/31/04
    Comment      AS STRING * 25
    
END TYPE

TYPE EmpData3Type
    Data1RecNum     AS INTEGER
    YTDGrossPay     AS DOUBLE
    YTDSocGrossPay  AS DOUBLE
    YTDMedGrossPay  AS DOUBLE
    YTDFedGrossPay  AS DOUBLE
    YTDStaGrossPay  AS DOUBLE
    YTDOTPay        AS DOUBLE
    YTDRegPay       AS DOUBLE
    YTDNet          AS DOUBLE
    YTDSocial       AS DOUBLE
    YTDMedicare     AS DOUBLE
    YTDFederal      AS DOUBLE
    YTDState        AS DOUBLE
    YTDRetire       AS DOUBLE
    YTDDAmt(1 TO 50) AS DOUBLE
    YTDDAmtT        AS DOUBLE
    YTDEarn1        AS DOUBLE        'e
    YTDEarn2        AS DOUBLE
    YTDEarn3        AS DOUBLE
    YTDEarnT        AS DOUBLE
    YTDEIC          AS DOUBLE
    YTDOther2       AS DOUBLE
END TYPE

TYPE NameSortIdxType
  EmpName AS STRING * 14
  DataRecNum AS INTEGER
END TYPE

TYPE NumbSortIdxType
  EmpNumb AS STRING * 14   '14
  DataRecNum AS INTEGER    '2
END TYPE

TYPE EmployeeIdxType
  DataRecNum AS INTEGER
END TYPE

'npayroll

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

''''''''''''''
TYPE TransRecType
    TActive         AS INTEGER             '1     'Active Transaction Flag
    PrevTransRec    AS INTEGER             '2     'Pointer to employee's prev trans
    EmpPin          AS INTEGER             '3     'Pointer to employee rec
    PaySFlag        AS STRING * 1          '4     'Pay Salary Flag in time trans
    CheckNum        AS LONG                '5     'Payroll Check Number
    PayPdStart      AS INTEGER             '6     'Start of Pay Period
    PayPdEnd        AS INTEGER             '7     'End of Pay Period
    CheckDate       AS INTEGER   'yea      '8     'Date checks written
    PostDate        AS INTEGER             '9     'Date Transaction are posted
    PayType         AS STRING * 1          '10    'Salaried or Hourly
    BaseRate        AS DOUBLE              '11    'Base Rate or Salary Amt
    OTRate          AS DOUBLE              '12    'Overtime Rate
    RegHrsWork      AS DOUBLE             '13    'Hours worked this period
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
    GrossPay        AS DOUBLE              '      'Add Earnings + GrossWage
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
    DAmt(1 TO 50) AS DOUBLE              '      'Voluntary Deduction amounts / pcts
    TotDedAmt AS DOUBLE                    '      'Total Voluntary Deductions
    EICAmt     AS DOUBLE                   '
    NetPay    AS DOUBLE                    '
    PeriodHistRec AS INTEGER    'not used        '      'YTD Totals?
    MatchRetAmt     AS DOUBLE              '      'Employer's Retirement Match
    MatchSocAmt     AS DOUBLE              '      'Employer's Social Secity Match
    MatchMedAmt     AS DOUBLE              '      'Employer's Medicare Match
    RetGrossPay     AS DOUBLE              '      'Retirement Gross
    TaxFring        AS DOUBLE              '      'Taxable Fringe
    Less401k(1 TO 3)  AS INTEGER
    Voided          AS STRING * 1 'added 12/17/04
    VoidRec         AS LONG
    Pad1            AS STRING * 3
'-----------------------------
'can sum these for a report from transaction history file
END TYPE

TYPE ManualTransRecType
    PDSTART  AS INTEGER           ' 1
    PDEND    AS INTEGER           '
    ChkDate  AS INTEGER           '
    CheckNum AS LONG              '
    RegHrs   AS DOUBLE            ' 5
    SICKHRS  AS DOUBLE            '  6
    VACHRS   AS DOUBLE             ' 7
    COMPHRS  AS DOUBLE             ' 8
    PERSHRS  AS DOUBLE
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

    GrossPay AS DOUBLE              '19
    RegWage  AS DOUBLE              '20
    OTWage   AS DOUBLE              '21
    FEDTAX   AS DOUBLE              '22
    STATAX   AS DOUBLE              '23
    SOCTAX   AS DOUBLE              '24
    MEDTAX   AS DOUBLE              '25
    RETAMT   AS DOUBLE              '26

'    DAmt(1 To 12)  As Double        '27-38
    DAmt(1 TO 50)  AS DOUBLE        '27-38

    TOTDED   AS DOUBLE              '39
    EIC      AS DOUBLE              '40
    NetPay   AS DOUBLE              '41
    FedGross AS DOUBLE              '42
    STAGROSS AS DOUBLE              '43
    SocGross AS DOUBLE              '44
    MedGross AS DOUBLE              '45
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
    GrossPay   AS DOUBLE
    MATSocAmt  AS DOUBLE
    MATMedAmt  AS DOUBLE
    MATRetAmt  AS DOUBLE
END TYPE

TYPE YTDFundMnthType
  RegHrs    AS DOUBLE
  OTHrs     AS DOUBLE
  RegWage   AS DOUBLE
  OTWage    AS DOUBLE
'  VACHRS    AS DOUBLE
'  SICKHRS   AS DOUBLE
'  HOLHRS    AS DOUBLE
'  CTHrs     AS DOUBLE
'  CWHrs     AS DOUBLE
END TYPE


TYPE YTDFundRptType
  FundNum AS LONG
  Mths(1 TO 12) AS YTDFundMnthType
END TYPE

TYPE QtrDateType
  LDate   AS INTEGER
  HDate   AS INTEGER
END TYPE

TYPE K401RptType
'062498 Added RetType & EmpNum
  RetType AS STRING * 1
  EmpNum  AS STRING * 10
  EmpName AS STRING * 32
  SSN     AS STRING * 11
  VAmt    AS STRING * 11
  LAmt    AS STRING * 11
  MAmt    AS STRING * 11
  Batch   AS STRING * 11
  HDate   AS STRING * 10
  Gross   AS STRING * 10
END TYPE

TYPE DedRptType
  EmpName AS STRING * 35
  SSN     AS STRING * 11
  DAmt    AS STRING * 23
  CrLf    AS STRING * 2
END TYPE

TYPE EmpTransHistRptType
  EmpName        AS STRING * 35
  SSN            AS STRING * 11
  CheckNum       AS LONG
  CheckDate      AS INTEGER   'yea
  GrossPay       AS DOUBLE    'this is add earnings and wages combined
  'GrossWage      AS DOUBLE   'this is also in the transaction history
  TotDedAmt      AS DOUBLE
  NetPay         AS DOUBLE
END TYPE

TYPE EmpHistFormType
    FirstEmp     AS LONG
    LastEmp      AS LONG
    StartDate    AS INTEGER
    EndDate      AS INTEGER
    SumOnly      AS STRING * 1
END TYPE

TYPE EmpHistoryRptType
    RegHrs       AS DOUBLE
    VACHRS       AS DOUBLE
    SICKHRS      AS DOUBLE
    HOLHRS       AS DOUBLE
    COMPHRS      AS DOUBLE
    TotalHrs     AS DOUBLE
'*******
    PHrs         AS DOUBLE
'*******
    TotHrs       AS DOUBLE

    TOTPaid      AS DOUBLE
    TotEIC       AS DOUBLE
    TRegWage     AS DOUBLE
    TOTWage      AS DOUBLE
    GPay         AS DOUBLE
    SSTax        AS DOUBLE
    MTax         AS DOUBLE
    FTax         AS DOUBLE
    STax         AS DOUBLE
    RETTOT       AS DOUBLE
    RETMATT      AS DOUBLE
    TNetPay      AS DOUBLE
END TYPE

TYPE GrossWageRptType
    EmpNo        AS STRING * 11
    EmpName      AS STRING * 21
    GrossPay     AS DOUBLE
    FedGross     AS DOUBLE
    FEDTAX       AS DOUBLE
    SocGross     AS DOUBLE
    SOCTAX       AS DOUBLE
    MedGross     AS DOUBLE
    MEDTAX       AS DOUBLE
    'StaGross     AS DOUBLE
    EIC          AS DOUBLE
    STATAX       AS DOUBLE
END TYPE

TYPE GrossFundsType
   FundNum AS INTEGER
   DAmt    AS DOUBLE
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
END TYPE


TYPE ESCGrossWageRptType
    EmpName      AS STRING * 31
    ESSN         AS STRING * 11
    GrossPay     AS DOUBLE
END TYPE

TYPE ESC2DiskRecType1
    Qtr          AS STRING * 9     '1 - 9
    Fill1        AS STRING * 13    '10-22
    ESSN         AS STRING * 16    '23-38
    EName        AS STRING * 24    '39-62
    Seasonal     AS STRING * 1     '63
    GPay         AS STRING * 15    '64-78
    CrLf         AS STRING * 2     '79- 80
END TYPE

TYPE EQWRptRecType
  ENumb    AS STRING * 11
  EName    AS STRING * 21
  GPay     AS STRING * 11
  FedGr    AS STRING * 11
  FedTx    AS STRING * 11
  SocGr    AS STRING * 11
  SocTx    AS STRING * 11
  MedGr    AS STRING * 11
  MedTx    AS STRING * 11
  'StaGr    AS STRING * 11
  EIC      AS STRING * 11
  StaTx    AS STRING * 11
  CrLf     AS STRING * 2
END TYPE

TYPE ESCMAG2DiskType
   Blank1       AS STRING * 1
   SSN          AS STRING * 9
   LastName     AS STRING * 12
   EmpInitials  AS STRING * 2
   EmpWages     AS STRING * 9
   SeasInd      AS STRING * 1   '"N"
   RemitNumb    AS STRING * 6
   EmplrAcct    AS STRING * 7   'Employer Account Number
   BranchAcct   AS STRING * 7
   RQuarter     AS STRING * 1
   RYear        AS STRING * 4
   EmplrName    AS STRING * 20
   Blank2       AS STRING * 1
   CrLf         AS STRING * 2
END TYPE

TYPE CompSortType
  CompCode AS STRING * 12
  RecNo    AS LONG
END TYPE

TYPE EmpSortType
  EmpNo AS STRING * 14
  RecNo AS INTEGER
END TYPE
 
TYPE LeaveEntryType
  YEARS   AS INTEGER
  EARN    AS DOUBLE
END TYPE

TYPE LeaveRecType
  VacMax   AS DOUBLE
  VEntry(1 TO 20)  AS LeaveEntryType
  SICKMAX  AS DOUBLE
  SEntry(1 TO 20)  AS LeaveEntryType
  PerMax   AS DOUBLE
  PEntry(1 TO 20)  AS LeaveEntryType
  HolMax   AS DOUBLE
  HEntry(1 TO 20)  AS LeaveEntryType
END TYPE

TYPE DosPRNSetupRecType
    RPT1     AS INTEGER
    RPT2     AS INTEGER
    RPT3     AS INTEGER
    RPT4     AS INTEGER
    RPT5     AS INTEGER
    RPT6     AS INTEGER
    RPT7     AS INTEGER
    RPT8     AS INTEGER
    RPT9     AS INTEGER
    RPT10    AS INTEGER
    RPT11    AS INTEGER
    RPT12    AS INTEGER
    RPT13    AS INTEGER
    RPT14    AS INTEGER
    RPT15    AS INTEGER
    RPT16    AS INTEGER
END TYPE

TYPE PRNSetupRecType
    Printer AS STRING * 20
    RPT(1 TO 18) AS INTEGER
'    RPT(1 To 19) As Integer '8/13 went from 18 to 19
    'when "Checks by Number" report was added
    CheckType AS INTEGER
END TYPE


TYPE PRDEDType
   DCode       AS STRING * 10
   DAmt        AS DOUBLE
   YTDDAmt     AS DOUBLE
END TYPE

TYPE PRCheckRecType
   CActive       AS INTEGER
   CheckNum      AS LONG
   CheckDate     AS INTEGER

   EmpName       AS STRING * 33
   EmpNo         AS STRING * 10
   EmpSSN        AS STRING * 11

'=-=-=-=-=-
   EmpAddr1 AS STRING * 36
   EmpCity  AS STRING * 24
   EmpState AS STRING * 2
   EmpZip   AS STRING * 10
'-=-=-=-=-=-=

   PayEndDate    AS INTEGER
   BaseRate      AS DOUBLE
   GrossPay      AS DOUBLE
   FedTaxAmt     AS DOUBLE
   StaTaxAmt     AS DOUBLE
   MedTaxAmt     AS DOUBLE
   SocTaxAmt     AS DOUBLE
   TotDedAmt     AS DOUBLE

'added
   RetireAmt     AS DOUBLE

   NetPay        AS DOUBLE
   YTDGrossPay   AS DOUBLE
   YTDFederal    AS DOUBLE
   YTDState      AS DOUBLE
   YTDSocial     AS DOUBLE
   YTDMedicare   AS DOUBLE
   YTDTotDed     AS DOUBLE
   YTDNetPay     AS DOUBLE

'added
   YTDRetire     AS DOUBLE

   VactBal       AS DOUBLE   '
   SickBal       AS DOUBLE   '
   CompBal       AS DOUBLE

'added-----
   CompEarn      AS DOUBLE
   RegHrsWork    AS DOUBLE
   OTHrsPaid     AS DOUBLE
   TotRegWage    AS DOUBLE
   VacUsed       AS DOUBLE
   SickUsed      AS DOUBLE
   CompUsed      AS DOUBLE
   HolUsed       AS DOUBLE
   PerUsed       AS DOUBLE

   RegHrsPaid    AS DOUBLE
   TotOTWage     AS DOUBLE

   AEarn(1 TO 3) AS PRDEDType

   TotAdditEarn  AS DOUBLE

   EICAmt        AS DOUBLE
   TaxFring      AS DOUBLE

'----------
'   CDED(1 To 12) As PRDEDType
   CDED(1 TO 50) AS PRDEDType
   DDFlag        AS INTEGER
END TYPE

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-

TYPE ScrnCalcType
   REGEARN AS DOUBLE
   OTEARN AS DOUBLE
   ALTEARN1 AS DOUBLE
   ALTEARN2 AS DOUBLE
   ALTEARN3 AS DOUBLE
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
   Ded(1 TO 50) AS DOUBLE
   SOCTAX AS DOUBLE
   MEDTAX AS DOUBLE
   FEDTAX AS DOUBLE
   STATAX AS DOUBLE
   RETIRE AS DOUBLE
   GrossPay AS DOUBLE
   TOTDED AS DOUBLE
   EIC AS DOUBLE
   NetPay AS DOUBLE
  
END TYPE

TYPE HourlyIDistType
    DAcct    AS STRING * 14    '  10 WAGEDST1 AS STRING * 14
    DRHrs    AS DOUBLE         '  11 D1REGHRS AS DOUBLE
    DOHrs    AS DOUBLE         '  12 D1OTHRS  AS DOUBLE
END TYPE

TYPE HourlyInputType
    WORKHRS  AS DOUBLE         '  1  WORKHRS  AS DOUBLE
    VACHRS   AS DOUBLE         '  2  VACHRS   AS DOUBLE
    SICKHRS  AS DOUBLE         '  3  SICKHRS  AS DOUBLE
    HOLHRS   AS DOUBLE         '  4  HOLHRS   AS DOUBLE
    COMPHRS  AS DOUBLE         '  5  COMPHRS  AS DOUBLE

'022498 New for Wrightsville beech
    PerHRS   AS DOUBLE         '  6
'*********************

    TOTHRSPD AS DOUBLE         '  6  TOTHRSPD AS DOUBLE
    OTWORKED AS DOUBLE         '  7  OTWORKED AS DOUBLE
    OTHRSPD  AS DOUBLE         '  8  OTHRSPD  AS DOUBLE
    OT2Comp  AS DOUBLE         '  9  OT2COMP  AS DOUBLE

    HDist(1 TO 8)  AS HourlyIDistType

    DTREGHRS AS DOUBLE         '  34 TOTREGHR AS DOUBLE
    DTOTHRS  AS DOUBLE         '  35 TOTOTHRS AS DOUBLE

    ALTEARN1 AS DOUBLE         '  36 ALTEARN1 AS DOUBLE
    ALTEARN2 AS DOUBLE         '  37 ALTEARN2 AS DOUBLE
    ALTEARN3 AS DOUBLE         '  38 ALTEARN3 AS DOUBLE

    TaxFring  AS DOUBLE

    AERNDST1 AS STRING * 14    '  39 AERNDST1 AS STRING * 14
    AERNAMT1 AS DOUBLE         '  40 AERNAMT1 AS DOUBLE
    AERNDST2 AS STRING * 14    '  41 AERNDST2 AS STRING * 14
    AERNAMT2 AS DOUBLE         '  42 AERNAMT2 AS DOUBLE
    AERNDST3 AS STRING * 14    '  43 AERNDST3 AS STRING * 14
    AERNAMT3 AS DOUBLE         '  44 AERNAMT3 AS DOUBLE
    AERNDST4 AS STRING * 14    '  45 AERNDST4 AS STRING * 14
    AERNAMT4 AS DOUBLE         '  46 AERNAMT4 AS DOUBLE
    AERNDST5 AS STRING * 14    '  47 AERNDST5 AS STRING * 14
    AERNAMT5 AS DOUBLE         '  48 AERNAMT5 AS DOUBLE
    AERNDST6 AS STRING * 14    '  49 AERNDST6 AS STRING * 14
    AERNAMT6 AS DOUBLE         '  50 AERNAMT6 AS DOUBLE
    TOTAERN  AS DOUBLE         '  51 TOTAERN  AS DOUBLE
    TOTEADST AS DOUBLE         '  52 TOTEADST AS DOUBLE
END TYPE

TYPE SalaryIDistType
    DAcct    AS STRING * 14    '  10 WAGEDST1 AS STRING * 14
    DPct     AS DOUBLE         '  11 D1REGHRS AS DOUBLE
END TYPE

TYPE SalaryInputType
    PAYSAL   AS STRING * 1     '  1
    VACHRS   AS DOUBLE         '  2
    SICKHRS  AS DOUBLE         '  3
    HOLHRS   AS DOUBLE         '  4
'022498 New for Wrightsville beech
    PerHRS   AS DOUBLE         '  5
    COMPHRS  AS DOUBLE 'added 9/1/04
    Hrs2Cmp  AS DOUBLE 'added 9/1/04
'*********************************
    SDist(1 TO 8)   AS SalaryIDistType

    TotOTHrs AS DOUBLE         '  22

    ALTEARN1 AS DOUBLE         '  23
    ALTEARN2 AS DOUBLE         '  24
    ALTEARN3 AS DOUBLE         '  25

    TaxFring AS DOUBLE         '  26

    AERNDST1 AS STRING * 14    '  27
    AERNAMT1 AS DOUBLE         '  28
    AERNDST2 AS STRING * 14    '  29
    AERNAMT2 AS DOUBLE         '  30
    AERNDST3 AS STRING * 14    '  31
    AERNAMT3 AS DOUBLE         '  32
    AERNDST4 AS STRING * 14    '  33
    AERNAMT4 AS DOUBLE         '  34
    AERNDST5 AS STRING * 14    '  35
    AERNAMT5 AS DOUBLE         '  36
    AERNDST6 AS STRING * 14    '  37
    AERNAMT6 AS DOUBLE         '  38

    TOTAERN  AS DOUBLE         '  39
    TOTEADST AS DOUBLE         '  40
END TYPE

TYPE GLIFDataType12              'Hummm.
   TranAcct   AS STRING * 12     ' 12 AS ARTRANACCT$
   TranDate   AS STRING * 8      ' 8 AS artrandate$
   TranDesc   AS STRING * 20     ' 19 AS artrandesc$
   CrAmt      AS DOUBLE          ' 8 AS ARCRAMT$
   DrAmt      AS DOUBLE          ' 8 AS ARDRAMT$
   Detail     AS DOUBLE          ' 8 AS ardetail$
   Source     AS STRING * 8      ' 8 AS arsource$
   rRNEXTTR   AS SINGLE          ' 4 AS ARNEXTTR$
   Fill       AS STRING * 8
   FromFlag   AS STRING * 1
END TYPE

TYPE SplitDedsType
   Acct       AS STRING * 14
   FEDTAX     AS DOUBLE
   STATAX     AS DOUBLE
   MEDTAX     AS DOUBLE
   SOCTAX     AS DOUBLE
   RetTax     AS DOUBLE
   EICPay     AS DOUBLE
'   DedAmts(1 To 12) As Double
   DedAmts(1 TO 50) AS DOUBLE

END TYPE

TYPE GLIFDataType14
   TranAcct   AS STRING * 14   'For New G/L
   TranDate   AS STRING * 8
   TranDesc   AS STRING * 20
   CrAmt      AS DOUBLE
   DrAmt      AS DOUBLE
   Detail     AS DOUBLE
   Source     AS STRING * 8
   rRNEXTTR   AS SINGLE
   Fill       AS STRING * 6    'Adj. from 8 to 6
   FromFlag   AS STRING * 1
END TYPE

TYPE FundType
   FundCode   AS STRING * 14
   Credit     AS DOUBLE
   Debit      AS DOUBLE
   Net        AS DOUBLE
END TYPE

TYPE AcctSumType
   FundCode   AS STRING * 14
   Credit     AS DOUBLE
   Debit      AS DOUBLE
   'Net        AS DOUBLE
END TYPE

TYPE JGLAcctIdxType
  AcctNum AS STRING * 14
  RecNo   AS INTEGER
END TYPE

TYPE GLAcctIdxType
  AcctNum AS DOUBLE
  RecNo   AS SINGLE
END TYPE

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

TYPE GLSetupRecType                 'still under const.
   UserName    AS STRING * 30
   TotAcctLen  AS INTEGER
   FundLen     AS INTEGER
   AcctLen     AS INTEGER
   DetLen      AS INTEGER
   CashAcct    AS STRING * 14
   APAcct      AS STRING * 14
   EncAcct     AS STRING * 14
   FBAcct      AS STRING * 14
   FYBeg       AS INTEGER
   FYEnd       AS INTEGER
   NYBeg       AS INTEGER
   NYEnd       AS INTEGER
   CDCash      AS STRING * 14  'new
   CDDue       AS STRING * 14
   CDActive    AS STRING * 1
   CRCashAcct  AS STRING * 14
   DeptCode    AS STRING * 1
   LPDate      AS INTEGER
   HPDate      AS INTEGER
   CDCashAcct  AS STRING * 14
   CDsbCash    AS STRING * 14
   APChkCode   AS INTEGER
   POStop      AS INTEGER    ' this is new 7/22/02 for tabstop on invoice entry
 'Fields added for V205
   PSLFlag     AS INTEGER   '1 for default to Yes, 0 for No
   DupInvFlag  AS INTEGER   '1 to allow duplicates, 0 for No
   CRBank      AS INTEGER   'banknum as default on entry
   CDBank      AS INTEGER   'banknum for default on entry
   ChkBank     AS INTEGER   'banknum for default on check printing
END TYPE

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

TYPE GLAcctRecType                  'Account Record Type: ? bytes
   Deleted     AS INTEGER           'Active Account Flag
   Num         AS STRING * 14       'Formatted Account Number
   Title       AS STRING * 30       'Account Description
   Typ         AS STRING * 1        'Account Type
   FrstTran    AS LONG              'Pointer to First Trans
   LastTran    AS LONG              'Pointer to Last Trans
   PYAct       AS DOUBLE            'Prior Year Actual
   BegBal      AS DOUBLE            'Beginning Balance
   Bgt         AS DOUBLE            'Budget Amount
   Bal         AS DOUBLE            'Running Balance
   Encumb      AS DOUBLE            'Encumbered Amount
   MTD         AS DOUBLE            'Month to Date Bal (calc as needed)
   YTD         AS DOUBLE            'Year to Date Bal (calc as needed)
   NYEst       AS DOUBLE            'Bgt Estimate
   NYReq       AS DOUBLE            'New Year Bgt Requested
   NYRec       AS DOUBLE            'New Year Bgt Recommended
   NYApp       AS DOUBLE            'New Year Bgt Approved
   FrstBTran   AS INTEGER           'Pointer to First Budget Trans
   LastBTran   AS INTEGER           'Pointer to Last Budget Trans
   FrstPTran   AS INTEGER           'Pointer to First Budget Trans
   LastPTran   AS INTEGER           'Pointer to Last Budget Trans
   'Res         AS STRING * 25       'Reserved for future needs
   Work        AS DOUBLE            'Temp added 08/17/96 for closeout
'edit the res added function rec pointer 6/11/04
   FNCTRec     AS LONG
   Res         AS STRING * 12
   ChkByte     AS STRING * 1    'this is updated at GASB34 conversion with chr$(1)
   Marked      AS INTEGER           '
END TYPE


TYPE OSChekRecType
   Chknum   AS LONG         '4 AS chknum$
   ChkDate  AS INTEGER       '8 AS chkdate$
   Desc     AS STRING * 30   '30 AS chkdesc$
   Amt      AS DOUBLE        '4 AS chkamt$
   Src      AS INTEGER       '2 AS CHKSOURCE$
   Cleared  AS INTEGER       'added by JB
   BankCode AS INTEGER
   Filler   AS STRING * 12   '16 AS nul$
END TYPE

'Type OSChkRecType
'   ChkNum   As Single        '4 AS chknum$
'   CHKDATE  As String * 8    '8 AS chkdate$
'   Desc     As String * 30   '30 AS chkdesc$
'   Amt      As Single        '4 AS chkamt$
'   Src      As Integer       '2 AS CHKSOURCE$
'   Cleared  As Integer       'added by JB
'   BankCode As Integer       '09/07/96 Gate City's
'   filler   As String * 12 '14   '16 AS nul$
'End Type

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

TYPE GLTransRecType                 'Transaction Record: 128 bytes
   AcctRec     AS INTEGER           'Pointer to Acct Record
   AcctNum     AS STRING * 14       'Formatted Acct Number string
   TRDATE      AS INTEGER           'Date2Num function
   Desc        AS STRING * 20       'Transaction Description
  'add extra desc for V205
   LDesc       AS STRING * 32       'Originally for invoice desc but use for other trans as well
   Ref         AS STRING * 8        'Document Reference
   DrAmt       AS DOUBLE            'Debit Amount
   CrAmt       AS DOUBLE            'Credit Amount
   Src         AS STRING * 8        'Module Source Code
   NextTran    AS LONG              'Pointer to Next Trans
   Res         AS STRING * 20       'Reserved for future needs
   Marked      AS INTEGER
END TYPE

TYPE TrEditRecType                  'Experimental GJ edit record:
   Deleted     AS INTEGER           'Deleted transaction flag
   Posted      AS INTEGER           'Posted flag
   AcctRec     AS INTEGER           'Pointer to acct rec
   AcctNum     AS STRING * 14       'Formatted Acct number string
   AcctName    AS STRING * 30       'Account Title
   TRDATE      AS INTEGER           'Date2Num
   DrAmt       AS DOUBLE            'Transaction Debit Amount
   CrAmt       AS DOUBLE            'Transaction Credit Amount
   EType       AS STRING * 1        'Entry Type (Debit/Credit)
   Desc        AS STRING * 20       'Transaction Description
   Ref         AS STRING * 8        'Document Reference #
   Src         AS STRING * 8        'Module Source Code
   'Res         AS STRING *          'Reserve for future needs
END TYPE

TYPE TrSortType                     'Used for sorting trans in history rpt
   TRDATE     AS INTEGER            'Transaction Date
   Record     AS LONG               'Pointer to transaction record
END TYPE

TYPE TrSortType1                    'Used for sorting trans in history rpt
   TRDATE     AS STRING * 12             'Transaction Date
   Record     AS LONG               'Pointer to transaction record
END TYPE

TYPE IFRecType
   AcctNum AS STRING * 9      '9 AS tranacct$
   TRDATE AS STRING * 8       '8 AS trandate$
   Desc AS STRING * 20        '20 AS trandesc$
   CrAmt AS DOUBLE            '8 AS cramt$
   DrAmt AS DOUBLE            '8 AS dramt$
   Ref AS STRING * 8          '8 AS detail$
   Src AS STRING * 8          '8 AS source$
   Filler AS STRING * 14      '4 AS nexttr$
   Posted AS INTEGER
END TYPE

TYPE GLFBAdjRecType
   AcctNum  AS STRING * 16
   AdjAmt AS DOUBLE
END TYPE

TYPE GLFundCloseRecType
   FundNum AS STRING * 3
END TYPE

TYPE BankRecType   '128 bytes
   Deleted AS INTEGER
   BankNum AS INTEGER
   BankName AS STRING * 25
   BankAcct AS STRING * 25
   GLAcct AS STRING * 25
   Pad AS STRING * 49
END TYPE

TYPE GLSumSetupRecType
   Beg1  AS STRING * 6
   End1  AS STRING * 6
   Desc1 AS STRING * 30
   Beg2  AS STRING * 6
   End2  AS STRING * 6
   Desc2 AS STRING * 30
   Beg3  AS STRING * 6
   End3  AS STRING * 6
   Desc3 AS STRING * 30
   Beg4  AS STRING * 6
   End4  AS STRING * 6
   Desc4 AS STRING * 30
   Beg5  AS STRING * 6
   End5  AS STRING * 6
   Desc5 AS STRING * 30
   Beg6  AS STRING * 6
   End6  AS STRING * 6
   Desc6 AS STRING * 30
   Beg7  AS STRING * 6
   End7  AS STRING * 6
   Desc7 AS STRING * 30
   Beg8  AS STRING * 6
   End8  AS STRING * 6
   Desc8 AS STRING * 30
   Beg9  AS STRING * 6
   End9  AS STRING * 6
   Desc9 AS STRING * 30
   Pad   AS STRING * 75
END TYPE

TYPE InvTaxAcct
    AcctNo  AS STRING * 16
    TaxAmt  AS DOUBLE
END TYPE

TYPE GLFundRecType                  'Fund Record Type: 64 bytes
   Deleted     AS INTEGER           'Deleted Flag
   FundNum     AS STRING * 4        'Fund Code
   Title       AS STRING * 30       'Fund Title
   Res         AS STRING * 28       'Reserve for future needs
END TYPE

TYPE GLAcctIndexType                'Account Index: 16 bytes
   AcctNum     AS STRING * 14       'Formatted account Number string
   RecNum      AS INTEGER           'Pointer to record
END TYPE

TYPE GLFundIndexType                'Fund Index : 16 bytes
   FundNum     AS STRING * 4        'Fund Number
   RecNum      AS INTEGER           'Pointer to record
END TYPE


TYPE PRDraftRecord1Type
    Field1 AS STRING * 1        ' Record Type Code Must = 1
    Field2 AS STRING * 2        ' Priority Code Must = 01
    Field3 AS STRING * 10       ' Immediate Destination Bank Transit Number (BB&T=b053101121 where b=blank space) Right Justified
    Field4 AS STRING * 10       ' Immediate Origin Bank Transit Number Must be Right Justified
    Field5 AS STRING * 6        ' Transmission File Creation Date (yymmdd)
    Field6 AS STRING * 4        ' Transmission File Creation Time (hhmm)
    Field7 AS STRING * 1        ' Field Modifier ID Must = A
    Field8 AS STRING * 3        ' Record Size  Must = 094
    Field9 AS STRING * 2        ' Blocking Factor  Must = 10
    Field10 AS STRING * 1       ' Format Code  Must = 1
    Field11 AS STRING * 23      ' Name of Destination Bank (Receiving Transmission)
    Field12 AS STRING * 23      ' Name of Originating Bank
    Field13 AS STRING * 8       ' Reserved Needs to be 8 blanks padded
END TYPE


TYPE PRDraftRecord5Type
    Field1 AS STRING * 1        ' Record Type Code Must = 5
    Field2 AS STRING * 3        ' Service Class Code Must = 200
    Field3 AS STRING * 16       ' Company Submitting Name in ALL CAPS
    Field4 AS STRING * 20       ' Discretionary Data
    Field5 AS STRING * 10       ' Prefix num & Company Federal Tax ID
    Field6 AS STRING * 3        ' Standard Entry Class (PPD for Direct Deposits and Drafts)
    Field7 AS STRING * 10       ' Company Entry Description SUCH AS UTIL BILL
    Field8 AS STRING * 6        ' Company Desc Date YYMMDD
    Field9 AS STRING * 6        ' Effective Entry Date YYMMDD
    Field10 AS STRING * 3       ' RESERVED LEAVE WITH 3 BLANKS
    Field11 AS STRING * 1       ' Must Equal 1 for Originator Status Code
    Field12 AS STRING * 8       ' Originating Fin. Inst. ID  05310112 for BB&T
    Field13 AS STRING * 7       ' Batch Number Beginning with 0000001
  END TYPE

TYPE PRDraftRecord6Type
    Field1 AS STRING * 1        ' Record Type Code  Must = 6
    Field2 AS STRING * 2        ' Transaction Code      22 Credit Checking
                                '                       27 Debit Checking
                                '                       32 Credit Savings
                                '                       37 Debit Savings
                                '                       28 PRENOTE DEBIT CHECKING  (Amt Must be all zeros
                                '                       Generally Add 1 to Get PreNote Transaction Code
    Field3 AS STRING * 8        ' Individual's Bank ID Transit Routing #
    Field4 AS STRING * 1        ' Transit Routing Check Digit
    Field5 AS STRING * 17       ' Individual's Bank Account Number
    Field6 AS STRING * 10       ' Amount (assume 2 decimal places)
    Field7 AS STRING * 15       ' Individual's ID Number to Customer (Usually Customer Utility Account Number)
    Field8 AS STRING * 22       ' Individual's Name
    Field9 AS STRING * 2        ' Set to 2 spaces (Not Used)
    Field10 AS STRING * 1       ' Set to '0' to signify no addenda records
    Field11 AS STRING * 15      ' Trace Number
                                ' Consists of Bank ID # 05310112 plus
                                ' Line Item Number starting w/ 0000001
                                ' and incrementing once for each line (Record6)
END TYPE


TYPE PRDraftRecord8Type
    Field1 AS STRING * 1        ' Record Code Must = 8
    Field2 AS STRING * 3        ' Service Class Code  Must = 200
    Field3 AS STRING * 6        ' Number of Detail (TYPE 6) Records
    Field4 AS STRING * 10       ' Hash Total
                                ' Hash#=Hash#+val(banktransit#) for Each Type 6 Record
    Field5 AS STRING * 12       ' Total Debit Amount
    Field6 AS STRING * 12       ' Total Credit Amount
    Field7 AS STRING * 10       ' Federal ID Tax Number
    Field8 AS STRING * 19       ' RESERVED KEEP BLANK
    Field9 AS STRING * 6        ' RESERVED BY FEDERAL RESERVE BANK
    Field10 AS STRING * 8       ' Originating Financial Inst. 05310112 for BB&T
    Field11 AS STRING * 7       ' Batch # Beginning with 0000001
END TYPE

TYPE PRDraftRecord9Type
    Field1 AS STRING * 1        ' Record Code Must=9
    Field2 AS STRING * 6        ' Batch Count (Sum of Batches) NORMALLY 000001
    Field3 AS STRING * 6        ' Block Count Number of Records
                                ' Found by taking Total Size of File and Dividing by 940
    Field4 AS STRING * 8        ' Total 06 Record Type Entries
    Field5 AS STRING * 10       ' Enter Hash  See Above
    Field6 AS STRING * 12       ' Total Debit Entry Dollar Amount (Assume 2 decimals)
    Field7 AS STRING * 12       ' Total Credit Entry Dollar Amount (Assume 2 decimals)
    Field8 AS STRING * 39       ' RESERVED FOR FUTURE USE
END TYPE

TYPE DetailRecType
  ID          AS STRING * 1
  Batch       AS STRING * 5
  PCN         AS STRING * 5
  ProcDate    AS STRING * 6
  SSN         AS STRING * 9
  EmpName     AS STRING * 30
  EmpVolDed   AS STRING * 8
  EmpLoanPay  AS STRING * 8
  EmpContAmt  AS STRING * 8
  CrLf        AS STRING * 2
END TYPE

TYPE TrailerRecType
  ID          AS STRING * 1
  TotVolDED   AS STRING * 11
  TotLoanAmt  AS STRING * 11
  TotContAmt  AS STRING * 11
  TotDRecs    AS STRING * 6
  Filler      AS STRING * 40
  CrLf        AS STRING * 2
END TYPE

TYPE W2DedType
    CHKDED  AS STRING * 20
    AMTBOX  AS STRING * 3
    DedCode AS STRING * 4
END TYPE

TYPE W2SetUpType
                           'only two options per check box  15c 15g
    ExtrYear AS INTEGER
'    CHKDED0  AS STRING * 20
'    AMTBOX0  AS STRING * 3
'    RETCODE  AS STRING * 4

    Deds(0 TO 12) AS W2DedType

END TYPE

TYPE TempAccrualType
    EMPVACE  AS DOUBLE
    EMPVBAL  AS DOUBLE
    EMPSLE   AS DOUBLE
    EMPSLBAL AS DOUBLE
    EMPHOLE  AS DOUBLE
    EMPHBAL  AS DOUBLE
    EMPPERE   AS DOUBLE
    EMPPBAL AS DOUBLE
END TYPE

TYPE AccrualDates
    PreviousDate AS LONG
    CurrentDate AS LONG
END TYPE

TYPE RetRecType
  Unit        AS STRING * 5   '1-5
  RetNum      AS STRING * 6   '6-11
  FirstN      AS STRING * 1   '12
  MidN        AS STRING * 1   '13
  LastN       AS STRING * 11  '14-24
  SSN         AS STRING * 9   '25-33
  Fill1       AS STRING * 10  '34-43
  EMonth      AS STRING * 2   '44-45
  EYear       AS STRING * 2   '46-47
  Fill2       AS STRING * 3   '48-50
  WageAmt     AS STRING * 7   '51-57
  NegWage     AS STRING * 1   '58
  RETAMT      AS STRING * 7   '59-65
  NegRet      AS STRING * 1   '66
  Fill3       AS STRING * 4   '67-70
  CrLf        AS STRING * 2
END TYPE

TYPE VoidDEDType
   DedGLNum       AS STRING * 14
   DAmt        AS DOUBLE
   DedDesc     AS STRING * 12
END TYPE


TYPE VoidCheckType
  EmpNum      AS STRING * 10  'employee number
  NumOfAccts  AS INTEGER      'number of wage accounts for this employee
  CheckNum    AS LONG         'used to match up the correct check
  CheckDate   AS INTEGER      'used to match up the correct check
  CheckAmt    AS DOUBLE       'used to match up the correct check
  TransRec   AS DOUBLE  'Pointer to transaction record for this check
  DedData(1 TO 50) AS VoidDEDType
  PRNet       AS DOUBLE
  PRNetGL     AS STRING * 14
  FEDWHAmt    AS DOUBLE
  FEDWHGL     AS STRING * 14
  STAWHAmt    AS DOUBLE
  STAWHGL     AS STRING * 14
  SOCWHAmt    AS DOUBLE
  SOCWHGL     AS STRING * 14
  SOCMATCRAmt AS DOUBLE
  SOCMATCRGL  AS STRING * 14
  SOCMATDBAmt AS DOUBLE
  SOCMATDBGL  AS STRING * 14
  MEDWHAmt    AS DOUBLE
  MEDWHGL     AS STRING * 14
  MEDMATCRAmt AS DOUBLE
  MEDMATCRGL  AS STRING * 14
  MEDMATDBAmt AS DOUBLE
  MEDMATDBGL  AS STRING * 14
  RETWHAmt    AS DOUBLE
  RETWHGL     AS STRING * 14
  RETMATCRAmt AS DOUBLE
  RETMATCRGL  AS STRING * 14
  RETMATDBAmt AS DOUBLE
  RETMATDBGL  AS STRING * 14
  WagesAmt    AS DOUBLE
  WagesGL     AS STRING * 14
  VoidFlag    AS INTEGER 'False = has not been voided
  PPEAmt      AS DOUBLE
  PPEGL       AS STRING * 14
  PPETotAmt   AS DOUBLE
  PPETotGL    AS STRING * 14
  Type        AS STRING * 1 '(C)entral Depository, (P)ool, (I)mprest
  Pad         AS STRING * 40
END TYPE
  
  
TYPE PayRateIndexType
    PayRateRec  AS INTEGER
    EmpName AS STRING * 48
END TYPE

TYPE PayRateIdxNumType
    PayRateRec  AS INTEGER
    EmpNum AS STRING * 10
END TYPE

TYPE TempDedAlertType
    DCDESC1  AS STRING * 10
    Number   AS INTEGER
END TYPE

TYPE TempEarnAlertType
    ERNCODE1 AS STRING * 10
    Number   AS INTEGER
END TYPE
  
TYPE K401DedType
    VolDed  AS STRING * 14
    LPDed   AS STRING * 14
END TYPE

'this type is temporary file kept on local drive for passing values between exe's
TYPE CitiPassTempType
  usernum   AS INTEGER
  UserName  AS STRING * 15
  frommdl   AS INTEGER   'this is to indicate to citipak ok to have file
END TYPE

TYPE MessLineType
  Msg AS STRING * 69
END TYPE

DIM a AS TransRecType
PRINT LEN(a)

