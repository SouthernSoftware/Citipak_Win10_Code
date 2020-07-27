TYPE HamletCustType
  District    AS STRING * 1
  Book        AS STRING * 2
  Seqnumb     AS STRING * 5
  AcctCtrl    AS STRING * 2      'order of people who live there
  CustName    AS STRING * 22
  ServAddr    AS STRING * 22
  Addr1       AS STRING * 22
  Addr2       AS STRING * 22
  Addr3       AS STRING * 22     'City & State???
  ZipCode     AS STRING * 9
  Class       AS STRING * 1
  LocCode     AS STRING * 1
  BankCode    AS STRING * 2
  BankAcct    AS STRING * 15
  BadChkCnt   AS STRING * 2
  LastBadDate AS STRING * 6
  Deposit     AS STRING * 7  'Fmt   5-2   00000.00
  DepNum      AS STRING * 8
  DepDate     AS STRING * 6
  RefundDate  AS STRING * 6
  RefundFlag  AS STRING * 1
  FirstNDate  AS STRING * 6
  SecondNDate AS STRING * 6
  Extensions  AS STRING * 2
  ExtensDate  AS STRING * 6
  StartDate   AS STRING * 6
  EndDate     AS STRING * 6
  ContService AS STRING * 1  '(y/n) ????
  SSN         AS STRING * 9  '3-2-4   Three fields
  Phone1      AS STRING * 3
  Phone2      AS STRING * 4
  AcctStatus  AS STRING * 1
  LstMnthBill AS STRING * 2  'Last month billed
  DueDate     AS STRING * 6
  AdjDate     AS STRING * 6
  LstPayDate  AS STRING * 6
  LastPayAmt  AS STRING * 7  'Fmt  5-2   00000.00
  AcctBal     AS STRING * 8  'Fmt  6-2  000000.00
  Sepp1       AS STRING * 1  '"+" in data file    '270
  WPrevDate   AS STRING * 6
  WCurrDate   AS STRING * 6
  WaterRate   AS STRING * 1  'coded 1-9
  WatUnits    AS STRING * 3  '
  WatStatCode AS STRING * 1  'codes are M,B and space  ????
  WatBegMtr   AS STRING * 8
  WatCurMtr   AS STRING * 8
  WatPreMtr   AS STRING * 8
  WatUsage    AS STRING * 8
  NoClue1     AS STRING * 2  '?????????
  MtrSerial   AS STRING * 10
  MtrOwner    AS STRING * 22  'wrong
  LastOccp    AS STRING * 2   'fixed
  
  WatMtrTCode AS STRING * 1
  WatUpDtFlag AS STRING * 1
  
  MtrMaker    AS STRING * 3
  MtrLoc      AS STRING * 1  'O outside????
  MtrInstDate AS STRING * 6
  WMtrCRDate  AS STRING * 6  'Water current read date??
  WMtrPRDate  AS STRING * 6  'Water previous read date??
  WatEst      AS STRING * 1  'byte 380
  WNoShutOff  AS STRING * 2  'Count number times water shut off???
  LastShutOff AS STRING * 6  'last date
  WatHyds     AS STRING * 3
  WatSpks     AS STRING * 4
  WatChrgs    AS STRING * 7  'Fmt 00000-00
  WatHydChrg  AS STRING * 7
  WatSpkChrg  AS STRING * 7
  WatOnOffChg AS STRING * 7
  WatAdjs     AS STRING * 7
  Sepp2       AS STRING * 1  '"+" in data file  Byte 431
  WatCredit   AS STRING * 7
  WatBalance  AS STRING * 7  'byte 445
  Sepp3       AS STRING * 1  '"+" in data file
'ok to here   byte 446
  WatPrevChrg AS STRING * 7   '453
  WatMtrSize  AS STRING * 2   '44 455
  WatFlatChrg AS STRING * 7
  WatBillCnt  AS STRING * 1  'Number of bills to print???
  WatAMtrPrev AS STRING * 8  '
  WatAMtrCurr AS STRING * 8
  WatBMtrPrev AS STRING * 8
  WatBMtrCurr AS STRING * 8
  WatCMtrPrev AS STRING * 8
  WatCMtrCurr AS STRING * 8
'ok to here   byte 511
  WatMessage  AS STRING * 20
  WatAcctBill AS STRING * 1
  Zip4        AS STRING * 10
'ok to here   byte 542
  Fill1       AS STRING * 3
'ok to here   byte 545
  SewRate     AS STRING * 1
  SewEst      AS STRING * 1
'ok to here   byte 547
  SewChrg     AS STRING * 7
  SewAdjs     AS STRING * 7
  Sepp4       AS STRING * 1  '"+" in data file
'ok to here   byte 562
  SewCredit   AS STRING * 7
  SewBalance  AS STRING * 7
  Sepp5       AS STRING * 1  '"+" in data file
'ok to here   byte 577
  SewPrevChrg AS STRING * 7
  SewFlatRate AS STRING * 7
  CRNum       AS STRING * 6
  Fill2       AS STRING * 12
'ok to here   byte 609
  GarCode     AS STRING * 1
  GarStatus   AS STRING * 1
  GarFDate    AS STRING * 6
  GarTDate    AS STRING * 6   'to date   623
  GarNumUnits AS STRING * 3   'Garbage units 626
  GarChrg     AS STRING * 7   '  633
  GarAdjs     AS STRING * 7   '
  Sepp6       AS STRING * 1   '"+" in data file
'ok to here   byte 641
  GarCredit   AS STRING * 7   '
  GarBalance  AS STRING * 7   '
  Sepp7       AS STRING * 1   '"+" in data file
'ok to here   byte 656
  GarFlatRate AS STRING * 7   '
  Fill3       AS STRING * 6
'ok to here   byte 669
  SeqTime     AS STRING * 6
  SeqDate     AS STRING * 6
'ok to here   byte 681
  Fill4       AS STRING * 94
  CrLf        AS STRING * 2
END TYPE

