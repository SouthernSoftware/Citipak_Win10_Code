TYPE ServicesType
  RATECODE AS STRING * 4
  RMtrType AS STRING * 1
END TYPE

TYPE FlatRateType
  FRDESC   AS STRING * 18
  FRAMT    AS DOUBLE
  FRFREQ   AS STRING * 1
  REVSRC   AS INTEGER
  NumMin   AS INTEGER
END TYPE

TYPE RevDataType
  RevName    AS STRING * 20
  RATECODE   AS STRING * 4
  RevMtrType AS STRING * 1
END TYPE

TYPE LocMeterType
  MtrNum    AS STRING * 12
  MTRMulti  AS INTEGER
  MtrType   AS STRING * 1
  MTRUnit   AS STRING * 1
  NumUser   AS INTEGER
  InsDate   AS INTEGER
  CurRead   AS LONG
  PrevRead  AS LONG
  CurDate   AS INTEGER
  PastDate  AS INTEGER       'hidden & protected
  ReadFlag  AS STRING * 1    'hidden & protected
  AvgUse    AS LONG          'hidden & protected
  UseCnt    AS INTEGER       'hidden & protected
  MtrIDNO   AS STRING * 11
  MtrLat    AS DOUBLE
  MtrLng    AS DOUBLE
END TYPE

TYPE MonthlyPayType
  AMTOWED      AS DOUBLE
  TotAmtPD     AS DOUBLE
  PayAmt       AS DOUBLE
  RevSource    AS INTEGER
END TYPE

TYPE UBTransRecType
   TransDate              AS INTEGER      '
   TransType              AS INTEGER      '
   TransDesc              AS STRING * 21  'may change
   Transamt               AS DOUBLE       'total revenue amount
   RevAmt(1 TO 15)        AS DOUBLE       'Revenue amounts
   TaxAmt(1 TO 15)        AS SINGLE       'Tax Amounts
'01-20-97 Added meter types field to hold meter type at time of transaction
   MtrTypes(1 TO 7)       AS INTEGER
'*******************
   CurRead(1 TO 7)        AS LONG         'Last/Current meter readings
   PrevRead(1 TO 7)       AS LONG         'Previous readings
   ESTREAD(1 TO 7)        AS STRING * 1   'Y/N Flags for meter est's
   BillNumber             AS LONG         'Number on the bill that Printed
   ReadDate               AS INTEGER
   BillDate               AS INTEGER
   PastDueDate            AS INTEGER
   DraftDate              AS INTEGER      '
'111398
   ProRatePCT             AS INTEGER
   ChkByte                AS STRING * 1   'Added check byte
   EPPFlag                AS STRING * 1   'Equal Payment Flag
   CustStatus             AS STRING * 1   'Customer Status at Time of Transaction
'020199
   EPPTrans               AS LONG         'Pointer to Equal Pay trans
   PenAtBill              AS SINGLE       'Used to flag IRR Meter (Sunset)
'****************
   PayTypeCode            AS INTEGER      'Payment Type:  1=Cash, 2=Check, 3=Cash/Check, 4=Charge
   OperatorNumber         AS INTEGER      '
   CustAcctNo             AS LONG         'Pointer to RecNo in ubcust.dat
   PrevTrans              AS LONG
   VoidFlag               AS INTEGER       'Changed for wadesboro
   FromCMFlag             AS INTEGER
   ActiveFlag             AS INTEGER      'Valid transaction flag
   RunBalance             AS DOUBLE
   CheckAmount            AS DOUBLE
   CashAmount             AS DOUBLE
   BillMsg                AS STRING * 20
   ApplyDepFlag           AS STRING * 1
   Posted2GL              AS STRING * 1
   PrevDate               AS INTEGER
   PenalFlag              AS STRING * 1
   TaxExempt              AS STRING * 1
   NONProfit              AS STRING * 1
END TYPE

TYPE NewUBCustRecType
  Book          AS STRING * 2
  SEQNUMB       AS STRING * 6
  Status        AS STRING * 1
  OPENDATE      AS INTEGER
  SEARCH        AS STRING * 10
  CustName      AS STRING * 35
  ADDR1         AS STRING * 35
  ADDR2         AS STRING * 35
  ServAddr      AS STRING * 35
  CITY          AS STRING * 18
  STATE         AS STRING * 2
  ZIPCODE       AS STRING * 10
  HPHONE        AS STRING * 14
  WPHONE        AS STRING * 14
  SOSEC         AS STRING * 11
  DRVLIC        AS STRING * 16
  CUSTTYPE      AS STRING * 3
  Addr911       AS STRING * 14
  '051498 added bill to field. Removed 1 byte from 911 addr
  BillTo        AS STRING * 1
  '********************************************************
  BILLCOPY      AS INTEGER
  POSTRTE       AS STRING * 4
  BILLCYCL      AS INTEGER
  ZONE          AS STRING * 3
  Seq           AS LONG
  'Page 2
  CASHONLY      AS STRING * 1
  LATEFEE       AS STRING * 1
  CUTOFFYN      AS STRING * 1
  TAXEXPT       AS STRING * 1
  SRCIT         AS STRING * 1
  EPPFlag       AS STRING * 1
  '032299 Modified for Bank draft account type
'    EPPAMT        AS DOUBLE
  Filler1       AS STRING * 7
  USEDRAFT      AS STRING * 1
  AcctType      AS STRING * 1
  '032299 Inserted account type
  BankName      AS STRING * 34
  BANKLOC       AS STRING * 30
  TRANSIT       AS STRING * 9
  BankAcct      AS STRING * 20
  BILLCMNT      AS STRING * 25
  PAYCMNT       AS STRING * 25
  PumpCode      AS STRING * 4
  USERCODE1     AS STRING * 4
  USERCODE2     AS STRING * 2
  ProRatePCT    AS INTEGER
  HHMSG1        AS STRING * 20
  HHMSG2        AS STRING * 20
  HHMSG3        AS STRING * 20
  'Page 3
  Serv(1 TO 15)      AS ServicesType
  FlatRates(1 TO 4)  AS FlatRateType
  'Page 4
  Monthly(1 TO 2)    AS MonthlyPayType
  MFEE1         AS DOUBLE
  MFEE2         AS DOUBLE
  LocMeters(1 TO 7)  AS LocMeterType
  'END OF Quick Screen Form
  CustPIN       AS LONG
  LastTrans     AS LONG
  CurrBalance   AS DOUBLE
  PrevBalance   AS DOUBLE
  CurrRevAmts(1 TO 15) AS DOUBLE
  PrevRevAmts(1 TO 15) AS DOUBLE
  DepositAmt    AS DOUBLE
  DelFlag       AS INTEGER
  PreNoteFlag   AS INTEGER
  WOLastTrans   AS LONG            'work order last trans pointer
  EstFlag       AS STRING * 1
  MessageRec    AS LONG             ' Points to Message Record
  OldRec        AS LONG
  EPPLastTran   AS LONG
  NewNotes      AS INTEGER
  DPCode        AS STRING * 2
  FillPad       AS STRING * 112
  ChkByte       AS STRING * 1
END TYPE


TYPE UBGilSensusGetReadRecType      ' File Layout For Retreiving Records
  CustLastName   AS STRING * 25
  CustFirstName  AS STRING * 25
  MeterID        AS STRING * 8
  Account        AS STRING * 8
  LowRead        AS STRING * 8
  HighRead       AS STRING * 8
  SensusType     AS STRING * 1        ' B=Touch Read : M=Manual
  PastRead       AS STRING * 8
  CurRead        AS STRING * 8
  ServAddress    AS STRING * 20
  LocationNumber AS STRING * 10
  fil1           AS STRING * 1
  ReadDate       AS STRING * 6
  Message        AS STRING * 19   'was 26
  DateRead       AS STRING * 4
  MtrLat         AS STRING * 11
  MtrLng         AS STRING * 11
  CrLf           AS STRING * 2
END TYPE


