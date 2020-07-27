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
  EmpName AS STRING * 32
  SSN     AS STRING * 11
  VAmt    AS STRING * 11
  LAmt    AS STRING * 11
  MAmt    AS STRING * 11
  Batch   AS STRING * 11
  HDate   AS STRING * 10
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
    TOTHrs       AS DOUBLE

    TOTPaid      AS DOUBLE
    TOTEIC       AS DOUBLE
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
    FedTax       AS DOUBLE
    SocGross     AS DOUBLE
    SocTax       AS DOUBLE
    MedGross     AS DOUBLE
    MedTax       AS DOUBLE
    'StaGross     AS DOUBLE
    EIC          AS DOUBLE
    StaTax       AS DOUBLE
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


