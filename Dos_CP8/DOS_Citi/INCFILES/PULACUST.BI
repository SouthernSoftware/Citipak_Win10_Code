TYPE PrintBillInfoType
    FrstBill    AS LONG
    LastBill    AS LONG
    BillDate    AS INTEGER
    PastDate    AS INTEGER
    PRDate      AS INTEGER
    CRDate      AS INTEGER
    DrftDate    AS INTEGER
    PastDate2   AS INTEGER
    PrnOrder    AS STRING * 25
    MsgLine1    AS STRING * 75
    MsgLine2    AS STRING * 75
    MsgLine3    AS STRING * 75
    MsgLine4    AS STRING * 75
END TYPE

TYPE BillMtrType
  CurrRead      AS LONG
  PrevRead      AS LONG
  MUsage        AS LONG
END TYPE

TYPE MetersOnBillType
  Mtr(1 TO 7) AS BillMtrType
END TYPE

 TYPE UBLateLetterType
  Head1    AS STRING * 40
  Head2    AS STRING * 40
  Head3    AS STRING * 40
  Head4    AS STRING * 40
  Head5    AS STRING * 40
  Body(1 TO 20) AS STRING * 75
END TYPE

TYPE NoticeInfoType
  FromBC        AS INTEGER
  ThruBC        AS INTEGER
  NoticeDate    AS INTEGER         '1
  PayByDate     AS INTEGER         '2
 'FromDate      AS INTEGER         '3
 'TODate        AS INTEGER         '4
  MinBalance    AS DOUBLE          '5
  BalanceType   AS INTEGER         '6
  PrnOrder      AS INTEGER         '7
  UseAFlag      AS INTEGER         '8
  MsgLine1      AS STRING * 25
  MsgLine2      AS STRING * 25
  MsgLine3      AS STRING * 25
  MsgLine4      AS STRING * 25
  PrnCnt        AS INTEGER
END TYPE

TYPE PenaltyInfoType
  PenDate    AS INTEGER
  PenDesc    AS STRING * 21
  RevSource  AS INTEGER
  ChargeOn   AS STRING * 17
'032299 Changed to a double variable
  PctCharge  AS DOUBLE
  'PctCharge  AS INTEGER
  AmtCharge  AS DOUBLE
  GreatLess  AS STRING * 1
  MinBalance AS DOUBLE
  CycFirst   AS INTEGER
  CycLast    AS INTEGER
  BookFirst  AS INTEGER
  BookLast   AS INTEGER
  PenCnt     AS INTEGER
END TYPE

TYPE PSAZipIndexType
  ZIPCODE  AS STRING * 10
  SName    AS STRING * 10
  RecNum   AS INTEGER
  Pad      AS STRING * 10
END TYPE

TYPE MOWZipIndexType
  ZIPCODE  AS STRING * 10
  RecNum   AS INTEGER
  FillPad AS STRING * 4
END TYPE

TYPE UBPostalIndexType
  ZIPCODE  AS STRING * 10
  Route    AS STRING * 4
  RecNum   AS INTEGER
END TYPE

TYPE UBServiceAddressIndexType
  ServiceAddress  AS STRING * 14
  RecNum   AS INTEGER
END TYPE

TYPE UBSequenceIndexType
  SeqNumber AS LONG
  RecNum    AS INTEGER
  Fill      AS STRING * 10      'This is to fill this to a 16 byte boundary
END TYPE

TYPE UBCustIndexRecType
  RecNum AS LONG
END TYPE

TYPE oUBCustReIndexRecType
  SearchName AS STRING * 10
  DelFlag    AS INTEGER
  RecNum     AS LONG
END TYPE

TYPE nUBCustReIndexRecType
  SearchName AS STRING * 10
  DelFlag    AS STRING * 1
  Status     AS STRING * 1
  RecNum     AS LONG
END TYPE

TYPE oUBCustReIndexRecType1
  SearchName AS STRING * 10
  First      AS STRING * 2
'  DelFlag    AS STRING * 1
'  Status     AS STRING * 1
  RecNum     AS LONG
END TYPE

TYPE UBLocaReIndexRecType
  Book       AS STRING * 2
  SEQNUMB    AS STRING * 6
  RecNum     AS LONG
  Pad        AS STRING * 4
END TYPE

TYPE UBLocaReIndexRecTypeVB
  BookSEQNUMB    AS STRING * 8
  RecNum         AS LONG
END TYPE

TYPE UBPINType
  PIN AS LONG
END TYPE

TYPE BookSeqRecType
  BookSeq  AS LONG
END TYPE

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

'Type OLocMeterType
'    MtrNum    As String * 12
'    MTRMulti  As Integer
'    MtrType   As String * 1
'    MTRUnit   As String * 1
'    NumUser   As Integer
'    InsDate   As Integer
'    CurRead   As Long
'    PrevRead  As Long
'    CurDate   As Integer
'    PastDate  As Integer       'hidden & protected
'    ReadFlag  As String * 1    'hidden & protected
'    AvgUse    As Long          'hidden & protected
'    UseCnt    As Integer       'hidden & protected
'    'MtrIDNO   as string * 11
'End Type

TYPE LocMeterType
    MtrNum    AS STRING * 12
    MTRMulti  AS INTEGER
    MtrType   AS STRING * 1
    MtrUnit   AS STRING * 1
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

'Type ONewUBCustRecType
'    Book          As String * 2
'    SEQNUMB       As String * 6
'    Status        As String * 1
'    OPENDATE      As Integer
'    SEARCH        As String * 10
'    CustName      As String * 35
'    ADDR1         As String * 35
'    ADDR2         As String * 35
'    ServAddr      As String * 35
'    CITY          As String * 18
'    STATE         As String * 2
'    ZIPCODE       As String * 10
'    HPHONE        As String * 14
'    WPHONE        As String * 14
'    SOSEC         As String * 11
'    DRVLIC        As String * 16
'    CUSTTYPE      As String * 3
'    Addr911       As String * 14
''051498 added bill to field. Removed 1 byte from 911 addr
'    BillTo        As String * 1
''********************************************************
'    BILLCOPY      As Integer
'    POSTRTE       As String * 4
'    BILLCYCL      As Integer
'    ZONE          As String * 3
'    Seq           As Long
''Page 2
'    CASHONLY      As String * 1
'    LATEFEE       As String * 1
'    CUTOFFYN      As String * 1
'    TAXEXPT       As String * 1
'    SRCIT         As String * 1
'    EPPFlag       As String * 1
''032299 Modified for Bank draft account type
''    EPPAMT        AS DOUBLE
'    Filler1       As String * 7
'    USEDRAFT      As String * 1
'    AcctType      As String * 1
''032299 Inserted account type
'    BankName      As String * 34
'    BANKLOC       As String * 30
'    TRANSIT       As String * 9
'    BankAcct      As String * 20
'    BILLCMNT      As String * 25
'    PAYCMNT       As String * 25
'    PumpCode      As String * 4
'    USERCODE1     As String * 4
'    USERCODE2     As String * 2
'    ProRatePCT    As Integer
'    HHMSG1        As String * 20
'    HHMSG2        As String * 20
'    HHMSG3        As String * 20
''Page 3
'    Serv(1 To 15)      As ServicesType
'    FlatRates(1 To 4)  As FlatRateType
''Page 4
'    Monthly(1 To 2)    As MonthlyPayType
'    MFEE1         As Double
'    MFEE2         As Double
'    LocMeters(1 To 7)  As OLocMeterType
''END OF Quick Screen Form
'    CustPIN       As Long
'    LastTrans     As Long
'    CurrBalance   As Double
'    PrevBalance   As Double
'    CurrRevAmts(1 To 15) As Double
'    PrevRevAmts(1 To 15) As Double
'    DepositAmt    As Double
'    DelFlag       As Integer
'    PreNoteFlag   As Integer
'    WOLastTrans   As Long            'work order last trans pointer
'    EstFlag       As String * 1
'    MessageRec    As Long            ' Points to Message Record
'    OldRec        As Long
'    EPPLastTran   As Long
'    NewNotes      As Integer
'    FillPad       As String * 4
'    ChkByte       As String * 1
'End Type

TYPE NewUBCustRecType
    Book          AS STRING * 2
    SEQNUMB       AS STRING * 6
    Status        AS STRING * 1
    OPENDATE      AS INTEGER
    SEARCH        AS STRING * 10
    CustName      AS STRING * 35
    ADDR1         AS STRING * 35
    ADDR2         AS STRING * 35
    SERVADDR      AS STRING * 35
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
    MessageRec    AS LONG            ' Points to Message Record
    OldRec        AS LONG
    EPPLastTran   AS LONG
    NewNotes      AS INTEGER
    DPCode        AS STRING * 2
    FillPad       AS STRING * 112
    ChkByte       AS STRING * 1
END TYPE

TYPE WrkOrdTextType
  Text(1 TO 6)  AS STRING * 67
END TYPE

TYPE WorkOrderRecType
    CustRec           AS LONG
    ENTRYDATE         AS INTEGER
    OrdersText        AS WrkOrdTextType
    RepliesText       AS WrkOrdTextType
    CompleteByDate    AS INTEGER
    CompletedDate     AS INTEGER
    PrevTransRec      AS LONG
END TYPE

'TYPE WorkOrderDefType
'    Deleted           AS Boolean
'    WOType            AS STRING * 20
'    OrdersText        AS WrkOrdTextType
'    RepliesText       AS WrkOrdTextType
'    Xtra              AS STRING * 20 'just in case
'END TYPE

TYPE Newport
    Acct AS STRING * 6
    Name AS STRING * 31
    Address AS STRING * 25
    CITY AS STRING * 15
    ST AS STRING * 2
    Zip AS STRING * 5
    ServAddress AS STRING * 20
    Source AS STRING * 1
    Ctype AS STRING * 1
    CLoc AS STRING * 1
    Blk AS STRING * 2
    Garb AS STRING * 1
END TYPE

TYPE SetUpAcctType
   RevName    AS STRING * 15
   DebitAcct  AS STRING * 14
   CreditAcct AS STRING * 14
END TYPE

TYPE RevSetUpType
    RevName AS STRING * 15
    UseDep   AS STRING * 1
    USERATE  AS STRING * 1
    TAXRATE  AS SINGLE
    UseMtr   AS STRING * 1
    DistOr   AS INTEGER
    ProRate  AS STRING * 1
END TYPE

TYPE UBSetupRecType
    UTILNAME        AS STRING * 35
    DEFCITY         AS STRING * 18
    DEFSTATE        AS STRING * 2
    ZIPCODE         AS STRING * 10
    PreByBook       AS STRING * 1
    RecpPort        AS STRING * 1
    RECPDEFT        AS STRING * 1
    ESTREAD         AS STRING * 1
    BANKDFT         AS STRING * 1
    UseSeq          AS STRING * 1
    BILLCYCL        AS STRING * 1
    DefLook         AS STRING * 1
    MethAcct        AS STRING * 1      'new 02-14-97
    SkipInactive    AS STRING * 1
    SkipSeparator   AS STRING * 1
    Make99File      AS STRING * 1
    LowRead         AS INTEGER
    HighRead        AS INTEGER
    HHDEVICE        AS STRING * 1    'P=PC3000 S=Sensus C=Syscom R=Radix N=None
    Revenues(1 TO 15) AS RevSetUpType
    BillAcct(1 TO 15) AS SetUpAcctType
    PayAcct(1 TO 15)  AS SetUpAcctType
    DepAcct(1 TO 15)  AS SetUpAcctType
END TYPE

'Added billsetup 9/29/03
TYPE UBBillSetupType
    Bill          AS INTEGER
    LateNotice    AS INTEGER
    PostBar       AS STRING * 1
    AcctBar       AS STRING * 1
    Permit        AS STRING * 1
    ChargeOn      AS STRING * 17
    PctCharge     AS DOUBLE
    AmtCharge     AS DOUBLE
    GreatLess     AS STRING * 1
    MinBalance    AS DOUBLE
    BL1Head1     AS STRING * 30
    BL1Head2     AS STRING * 40
    BL1Head3     AS STRING * 40
    BL1Permit1   AS STRING * 16
    BL1Permit2   AS STRING * 16
    BL1Permit3   AS STRING * 16
    BL1Permit4   AS STRING * 16
    BL1Permit5   AS STRING * 16
    BL1Opt1      AS STRING * 20
    BL1Opt2      AS STRING * 20
    BL1Opt3      AS STRING * 22
    BL1Opt4      AS STRING * 22
    BL1Opt5      AS STRING * 22
    BL1Opt6      AS STRING * 22
    BL1Opt7      AS STRING * 22
    BL1Opt8      AS STRING * 22
    BL1Opt9      AS STRING * 22
    BL1Opt10     AS STRING * 22
    AmtChge2     AS DOUBLE
    RtePrint     AS INTEGER
    reserved     AS STRING * 20
END TYPE
'Added 6/29/04
TYPE UBBillLetterType
    IncLogoFlag  AS INTEGER  ' 0 for not to print, 1 to print
    MtrNumFlag   AS INTEGER   '1 for Meter Serial, 2 for Meter ID
    BL1Head1     AS STRING * 30
    BL1Head2     AS STRING * 40
    BL1Head3     AS STRING * 40
    MsgOpt1      AS STRING * 40  'set to print on top section
    MsgOpt2      AS STRING * 40  'prints top section
    MsgPgph1     AS STRING * 125 'prints under service charges
    MsgPgph2     AS STRING * 125
    MsgPgph3     AS STRING * 125
    MsgPgph4     AS STRING * 125
    MsgPgph5     AS STRING * 125
    MsgOpt3      AS STRING * 50  'prints top of stub
    MsgOpt4      AS STRING * 30  'prints mid of stub
    MsgOpt5      AS STRING * 30  'prints mid right lower side of stub
    Padding      AS STRING * 95
END TYPE
'added 7/30/04 for Tray Labels for Postal Zip sorting Utility Bills
'Type PostTrayLabelType
'    '
'End Type
TYPE UBHHPathRecType
    PathWay AS STRING * 48
END TYPE

'Type UBSensusRecType
'    PathWay As String * 20
'End Type

'Type UBLogiconRecType
'    PathWay As String * 20
'End Type

'Type UBPC3000RecType
'    PathWay As String * 20
'End Type

TYPE DistArrayType
   DistOrder AS INTEGER
   DistCnt   AS INTEGER
END TYPE

' This Sensus Layout Files are Spec'd Out Exactly to Long View NC
'Type LUBSensusReadRecType        ' File Layout for Sending Out Records
'    ServAddress  As String * 20
'    MeterID      As String * 8
'    LowRead      As String * 8
'    HighRead     As String * 8
'    Account      As String * 8
'    SensusType   As String * 1        ' B=Touch Read : M=Manual
''    CustName     AS STRING * 25
''    SerialNumb   AS STRING * 8        'Added Per Mickey on 6-23-97
'End Type
'
'Type LUBSensusGetReadRecType     ' LONGVIEW File Layout For Retreiving Records
'    Account As String * 12
'    MeterID As String * 8
'    Reading As String * 8
'    DateRead As String * 4
'    NotUse2 As String * 2    'CRLF
'End Type
''*******************************************************************

TYPE UBSensusReadRecType        ' File Layout for Sending Out Records
    ServAddress  AS STRING * 20
    MeterID      AS STRING * 8
    LowRead      AS STRING * 8
    HighRead     AS STRING * 8
    Account      AS STRING * 8
    SensusType   AS STRING * 1        ' B=Touch Read : M=Manual
    CustName     AS STRING * 25
    SerialNumb   AS STRING * 8        'Added Per Mickey on 6-23-97
END TYPE

TYPE UBSensusGetReadRecType     ' File Layout For Retreiving Records
    Account AS STRING * 8
    NotUsed AS STRING * 5
    MeterID AS STRING * 8
    Filler AS STRING * 1
    Reading AS STRING * 8
    NotUse1 AS STRING * 1
    DateRead AS STRING * 6
    NotUse2 AS STRING * 4
END TYPE

TYPE UBLogiconReadRecType
    RecType   AS STRING * 1
    RouteNo   AS STRING * 2
    AcctNo    AS STRING * 6
    RecName   AS STRING * 30
    ServAddress AS STRING * 25
    ReadDate    AS STRING * 6
    ReadTime    AS STRING * 6
    Consumption AS STRING * 8
    PrevRead    AS STRING * 8
    CurRead     AS STRING * 8
    LowRead     AS STRING * 8
    HighRead   AS STRING * 8
    MtrNumb    AS STRING * 12
    CountChg   AS STRING * 1
    ForceFlag  AS STRING * 1
    ReportCode AS STRING * 2
    Remark     AS STRING * 40
    Label      AS STRING * 19
    PrintFlag  AS STRING * 1
    MessageOut AS STRING * 30
    Book       AS STRING * 2
'mod for Wrongsville beech!
    MtrSize    AS STRING * 2
    AvgUse     AS STRING * 8
    Future     AS STRING * 19
'    Future     As String * 29
    Recend     AS STRING * 1               'Must be 'X'
    CrLf       AS STRING * 2
END TYPE

TYPE UBLogiconGetReadRecType
    RecType   AS STRING * 1
    RouteNo   AS STRING * 2
    AcctNo    AS STRING * 6
    RecName   AS STRING * 30
    ServAddress AS STRING * 25
    ReadDate    AS STRING * 6
    ReadTime    AS STRING * 6
    Consumption AS STRING * 8
    PrevRead    AS STRING * 8
    CurRead     AS STRING * 8
    LowRead     AS STRING * 8
    HighRead   AS STRING * 8
    MtrNumb    AS STRING * 12
    CountChg   AS STRING * 1
    ForceFlag  AS STRING * 1
    ReportCode AS STRING * 2
    Remark     AS STRING * 40
    Label      AS STRING * 19
    PrintFlag  AS STRING * 1
    MessageOut AS STRING * 30
    Book       AS STRING * 2
    Future     AS STRING * 29
    Recend     AS STRING * 1               'Must be 'X'
    CrLf       AS STRING * 2
END TYPE

TYPE UBPC3000ReadRecType           'File Layout for Sending Out Records
    CustName      AS STRING * 20
    ServAddress   AS STRING * 20
    MeterID       AS STRING * 8
    LowRead       AS DOUBLE
    HighRead      AS DOUBLE
    Account       AS STRING * 8
    ReadFlag      AS STRING * 1        'Y/N
    MeterType     AS STRING * 1
    Book          AS INTEGER
    CurRead       AS DOUBLE
    PastRead      AS DOUBLE
    ReadDate      AS INTEGER
    ReadTime      AS STRING * 5
    Note1         AS STRING * 20
    Note2         AS STRING * 20
    Note3         AS STRING * 20
    NoteStatus    AS STRING * 1    'T=Temp Note  P=Perm Not'
'    MtrIDMST      As String * 0
'    MtrIDNO       As String * 1
END TYPE

TYPE UBDGRecType
    PathWay AS STRING * 20
END TYPE

TYPE UBDGHHRecType        ' File Layout for Sending Out Records
    RouteID AS STRING * 20
    SvcTyp AS STRING * 1
    CustName AS STRING * 25
    SvcLoc AS STRING * 21
    MeterSN AS STRING * 20
    MeterType AS STRING * 1       ' C for reg mtr   D for demand elec
    High AS STRING * 10
    Low AS STRING * 10
    Msg AS STRING * 110
    Account AS STRING * 10
    NewRdng AS STRING * 10
    NewDmnd AS STRING * 10
    Date AS STRING * 6
    Time AS STRING * 6
    NewAcctRte AS STRING * 20
END TYPE

'Itron Layouts
TYPE UBItronRecType
    PathWay AS STRING * 20
END TYPE

TYPE ItronFType                         'Header Record Type
    RecordCode AS STRING * 1
    Route AS STRING * 8
    Message AS STRING * 64
    Filler AS STRING * 5
    CrLf AS STRING * 2
END TYPE

TYPE ItronAType                         'Customer Record One
    RecordCode AS STRING * 1
    Route AS STRING * 8
    AcctNumb AS STRING * 10
    Geo AS STRING * 12
    SEQNUMB AS STRING * 5
    Message AS STRING * 1
    AType AS STRING * 1
    Filler AS STRING * 40
    CrLf AS STRING * 2
END TYPE

TYPE ItronBType                         'Customer Record Two
    RecordCode AS STRING * 1
    CustName AS STRING * 32
    CUSTADDR AS STRING * 32
    Filler AS STRING * 13
    CrLf AS STRING * 2
END TYPE

TYPE ItronHType
    RecordCode AS STRING * 1
    MeterNumb AS STRING * 12
    Dials AS STRING * 1
    LowRead AS STRING * 8
    HighRead AS STRING * 8
    LastRead AS STRING * 8
    MeterType AS STRING * 1
    Multiplier AS STRING * 6
    NoMths AS STRING * 1
    MtrMake AS STRING * 2
    DispCode AS STRING * 2
    NumbDec AS STRING * 1
    MustRead AS STRING * 1
    Status AS STRING * 1
    Filler AS STRING * 25
    CrLf AS STRING * 2
END TYPE

TYPE ItronZType
    RecordCode AS STRING * 1
    Route AS STRING * 8
    NumberAccts AS STRING * 4
    NumberMeters AS STRING * 4
    Filler AS STRING * 61
    CrLf AS STRING * 2
END TYPE

'Itron Read Layouts
TYPE ItronCType
    RecordCode AS STRING * 1           'Must be a C
    Route      AS STRING * 8
    Acct       AS STRING * 10
    SpecInst   AS STRING * 2
    Survey     AS STRING * 1
    ReadStatus AS STRING * 1
    ReSeqFlag  AS STRING * 1
    ReadDate   AS STRING * 4           'mmdd
    AcctChg    AS STRING * 1
    Filler     AS STRING * 51
END TYPE

TYPE ItronDType
     RecordCode AS STRING * 1           'Must be a D
     MeterNum   AS STRING * 12
     LocCode    AS STRING * 2
     MtrCon1    AS STRING * 2
     MtrCon2    AS STRING * 2
     Mult       AS STRING * 6
     ChgeFlag   AS STRING * 1
     Filler     AS STRING * 54
END TYPE

TYPE ItronEType
    RecordCode AS STRING * 1           'Must be a E
    NbrDials   AS STRING * 1
    CurReading AS STRING * 8           'Zero's if No Read
    DecPlaces  AS STRING * 1
    ReadTime   AS STRING * 6
    ReadChg    AS STRING * 1
    DispCode   AS STRING * 2
    ReadCount  AS STRING * 1
    ReadVerify AS STRING * 1
    NoReadCode AS STRING * 2
    Filler     AS STRING * 56
END TYPE

TYPE HHCodeRecType
  HHCRec  AS INTEGER
  HHCode  AS STRING * 20
END TYPE

TYPE RePrintIndexType
  BillNum  AS LONG
  BillRec  AS LONG
END TYPE

TYPE UBXferInfoType
  DAcctNo   AS STRING * 14
  DebitAmt  AS DOUBLE
  DRecNo    AS INTEGER
  DTitle    AS STRING * 30
  CAcctNo   AS STRING * 14
  CreditAmt AS DOUBLE
  CRecNo    AS INTEGER
  CTitle    AS STRING * 30
END TYPE

TYPE GJXferRecType
  RevText    AS STRING * 15
  BAcctInfo  AS UBXferInfoType     'Billing Accounts
  PAcctInfo  AS UBXferInfoType     'Payment Accounts
  DAcctInfo  AS UBXferInfoType     'Deposit Accounts
END TYPE

'Note:  if transaction is an adjustment then
'       CurRead field will contain the adjust amount
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

TYPE OldMessLineType
  Line AS STRING * 59
  LineDate AS STRING * 10
END TYPE

TYPE MessLineType
  Msg AS STRING * 69
END TYPE

TYPE UBMessRecType
  MessLine(1 TO 15) AS MessLineType
  CustRec AS LONG
END TYPE

TYPE TblBreakRecType
    UNITS      AS LONG
    UNITAMT    AS DOUBLE
END TYPE

TYPE UBRateTblRecType
    RATECODE AS STRING * 4
    RATEDESC AS STRING * 29
    ChkByte  AS STRING * 1
    MINAMT   AS DOUBLE
    MINUNITS AS LONG
    MaxAmt   AS DOUBLE
    TblBreaks(1 TO 10) AS TblBreakRecType
END TYPE

TYPE oUBRateTblRecType
    RATECODE AS STRING * 4
    RATEDESC AS STRING * 30
    MINAMT   AS DOUBLE
    MINUNITS AS LONG
    DiscPct  AS INTEGER
    TblBreaks(1 TO 10) AS TblBreakRecType
END TYPE

TYPE BookConsumpType
    Book             AS INTEGER
    CustCnt          AS LONG
    Consump(1 TO 15) AS DOUBLE
    RevAmt(1 TO 15)  AS DOUBLE
    TaxAmt(1 TO 15)  AS DOUBLE
END TYPE

TYPE PumpConsumpType
    PumpCode         AS STRING * 4
    CustCnt          AS LONG
    Consump          AS DOUBLE
END TYPE

TYPE UBOwnerRecType
    OwnLName  AS STRING * 20
    OwnFName  AS STRING * 15
    ADDR1     AS STRING * 35
    ADDR2     AS STRING * 35
    CITY      AS STRING * 18
    STATE     AS STRING * 2
    ZIPCODE   AS STRING * 10
    HPHONE    AS STRING * 14
    WPHONE    AS STRING * 14
    ChkByte   AS STRING * 1
END TYPE

TYPE MtrNotesType
  Note    AS STRING * 20
END TYPE

TYPE UBHuskyHHRecType                'File Layout for Sending Out Records
  CustName      AS STRING * 20
  ServAddress   AS STRING * 20
  UCode1        AS STRING * 2
  UCode2        AS STRING * 2
  MeterID       AS STRING * 8
  LowRead       AS LONG
  HighRead      AS LONG
  Account       AS STRING * 8
  ReadFlag      AS STRING * 1        'Y/N
  MeterType     AS STRING * 1
  Book          AS INTEGER
  CurRead       AS LONG
  PastRead      AS LONG
  ReadDate      AS INTEGER
  'ReadTime      AS STRING * 5
  NoteStatus    AS STRING * 1    'T=Temp Note  P=Perm Note
  Notes(1 TO 3)  AS MtrNotesType
END TYPE
'Added the following types when did the cycle count sum report, PS 5-7-03
TYPE MtrDateSortType
  MtrDate AS INTEGER
  RecNum   AS INTEGER
END TYPE

TYPE CycleType
   CustCnt AS LONG
   PendCnt AS LONG
END TYPE
'062603 Added draft info
TYPE oUBDraftRecType
    BANKDEST AS STRING * 9
    BANKORIG AS STRING * 9
    BankName AS STRING * 23
    BANKLOC  AS STRING * 23
    FEDID    AS STRING * 9
    FEDPREFX AS STRING * 1
END TYPE

TYPE oUBDraftRecType2
    BANKDEST AS STRING * 9
    BANKORIG AS STRING * 9
    BankName AS STRING * 23
    BANKLOC  AS STRING * 23
    COMPACCT AS STRING * 20
    FEDID    AS STRING * 9
    FEDPREFX AS STRING * 1
END TYPE

TYPE UBDraftRecType
    BANKDEST AS STRING * 9
    BANKORIG AS STRING * 9
    BankName AS STRING * 23
    BANKLOC  AS STRING * 23
    COMPACCT AS STRING * 20
    FEDID    AS STRING * 9
    FEDPREFX AS STRING * 1
    FileName AS STRING * 12
END TYPE

TYPE UBDraftPayRecType
    CustAcctNum   AS LONG
    DraftAmt      AS DOUBLE
END TYPE

TYPE UBDraftRecord1Type
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

TYPE UBDraftRecord5Type
    Field1 AS STRING * 1        ' Record Type Code Must = 5
    Field2 AS STRING * 3        ' Service Class Code Must = 200
    Field3 AS STRING * 16       ' Company Submitting Name in ALL CAPS
    Field4 AS STRING * 20       ' Discretionary Data
    Field5 AS STRING * 10       ' Company ID (Federal Tax ID Number)
    Field6 AS STRING * 3        ' Standard Entry Class (PPD for Direct Deposits and Drafts)
    Field7 AS STRING * 10       ' Company Entry Description SUCH AS UTIL BILL
    Field8 AS STRING * 6        ' Company Desc Date YYMMDD
    Field9 AS STRING * 6        ' Effective Entry Date YYMMDD
    Field10 AS STRING * 3       ' RESERVED LEAVE WITH 3 BLANKS
    Field11 AS STRING * 1       ' Must Equal 1 for Originator Status Code
    Field12 AS STRING * 8       ' Originating Fin. Inst. ID  05310112 for BB&T
    Field13 AS STRING * 7       ' Batch Number Beginning with 0000001
  END TYPE

TYPE UBDraftRecord6Type
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


TYPE UBDraftRecord8Type
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
    Field10 AS STRING * 8      ' Originating Financial Inst. 05310112 for BB&T
    Field11 AS STRING * 7       ' Batch # Beginning with 0000001
END TYPE

TYPE UBDraftRecord9Type
    Field1 AS STRING * 1        ' Record Code Must=9
    Field2 AS STRING * 6        ' Batch Count (Sum of Batches) NORMALLY 000001
    Field3 AS STRING * 6        ' Block Count Number of Records
                                ' Found by taking Total Size of File and Dividing ty 940
    Field4 AS STRING * 8        ' Total 06 Record Type Entries
    Field5 AS STRING * 10       ' Enter Hash  See Above
    Field6 AS STRING * 12       ' Total Debit Entry Dollar Amount (Assume 2 decimal)
    Field7 AS STRING * 12       ' Total Credit Entry Dollar Amount (Assume 2 decimal)
    Field8 AS STRING * 39       ' RESERVED FOR FUTURE USE
END TYPE

  TYPE DraftRptType
    TRANSIT  AS STRING * 9
    BankName AS STRING * 14
    'BANKNAME AS STRING * 34
    CustAcct AS STRING * 5
    CustName AS STRING * 23
    AcctType AS STRING * 1
    BillAmt  AS STRING * 9
    BankAcct AS STRING * 20
  END TYPE

  TYPE BDRptType
    BankName  AS STRING * 14
    CustRec   AS INTEGER
    TransRec  AS LONG
  END TYPE

  TYPE BankTotalsType
    BankName  AS STRING * 14
    Amount    AS DOUBLE
  END TYPE
  
TYPE PaidOwedType
    AMTOWE1  AS DOUBLE
    AMTPD1   AS DOUBLE
END TYPE
  
TYPE UBPaymentRecType
    OPERNUM         AS INTEGER
    PAYDATE         AS INTEGER
    CustAcct        AS LONG
    CustName        AS STRING * 24
    CUSTADDR        AS STRING * 24
    CUSTCMNT        AS STRING * 32
'052598 Added tax exempt flag changed cust comment len to 32
    TaxExempt       AS STRING * 1
    AMTOWED         AS DOUBLE
    TENDERTY        AS STRING * 12
    CASHAMT         AS DOUBLE
    CHKAMT          AS DOUBLE
    AMTRECD         AS DOUBLE
    Change          AS DOUBLE
    Desc            AS STRING * 19
    PaidOwed(1 TO 15)   AS PaidOwedType
    TOTOWED         AS DOUBLE
    AMTPAID         AS DOUBLE
'112801 Added cust status changed desc to 19
    Status          AS STRING * 1
END TYPE

TYPE PayListType
  CustRec   AS LONG
  Listrec   AS LONG
END TYPE

TYPE BookTotalType
  Count   AS LONG
  Cash    AS DOUBLE
  Check   AS DOUBLE
  Charge  AS DOUBLE
  Change  AS DOUBLE
END TYPE

TYPE CMOperRecType
    OperatorNumber AS INTEGER
    OperatorName AS STRING * 30
    OperatorPassword AS STRING * 8
    NotUsed AS STRING * 24
END TYPE

TYPE MiscCodeRecType
    MiscCode AS STRING * 7
    Description AS STRING * 25
    GlAcctNumb AS STRING * 14
    NotUsed AS STRING * 18
END TYPE

TYPE RMReceiptRecType
    RecName AS STRING * 25
    RecAddress AS STRING * 25
    RecDesc AS STRING * 25
    RecAmtOwed AS DOUBLE
    RecPayType AS INTEGER
    RecCashAmt AS DOUBLE
    RecCheckAmt AS DOUBLE
    RecChangeDue AS DOUBLE
    RecDate AS STRING * 10
    RecOperator AS STRING * 2
    RecptNumber AS SINGLE
    RecBlank AS STRING * 1
    RecBalance AS SINGLE
END TYPE

TYPE CMTransRecType
    TransDate    AS INTEGER
    TransAmount  AS DOUBLE
    TransCash    AS DOUBLE
    TransCheck   AS DOUBLE
    TransAmtOwed AS DOUBLE
    TransDesc    AS STRING * 25
    TransSource  AS INTEGER            '1-Misc 2-Util 3-Tax 4-License 5-decal
    TransName    AS STRING * 25
    TransAcctNum AS LONG               'Holds Master Acct Record Number in Mod
    TransDetNum  AS LONG               'Holds Record Number of Transaction Det
    TransRevAmt(1 TO 15) AS DOUBLE
    TransOperNum AS LONG
    Trans2GL      AS STRING * 1
    TransPad     AS STRING * 25
END TYPE

TYPE CMConfigType
    TOWNNAME AS STRING * 30
    CASHACCT AS STRING * 14
    LPTPORT  AS STRING * 1
    PrnDefYN AS STRING * 1
    ENDMSG   AS STRING * 30
END TYPE
  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
'This is for new local receipt setup file stored on each computer on
'drive c:\
'added ctlDefYN on 7/27/04
TYPE ReceiptPRNType
  RcpPort  AS STRING * 40
  PrnDefYN AS INTEGER
  CtlDefYN AS INTEGER
  ZExtra   AS STRING * 20
END TYPE

TYPE UBIntermecHHRecType         '
' First two fields are required by CE File I/O dll.
  CEVariant       AS STRING * 2  'MUST BE  CHR$(8) + CHR$(0)
  CEStrLen        AS STRING * 2  'MUST BE CHR$(165) + CHR$(0)
'**************************************************************
  CustName      AS STRING * 20
  ServAddress   AS STRING * 20
  MeterID       AS STRING * 8
  LowRead       AS STRING * 9
  HighRead      AS STRING * 9
  Account       AS STRING * 8
  ReadFlag      AS STRING * 1         'Y/N
  MeterType     AS STRING * 1
  Book          AS STRING * 2
  CurRead       AS STRING * 9
  PastRead      AS STRING * 9
  ReadDate      AS STRING * 8   ' fmt mmddyyyy
  Note1         AS STRING * 20
  Note2         AS STRING * 20
  Note3         AS STRING * 20
  NoteStatus    AS STRING * 1  'T=Temp  P=Perm
END TYPE

TYPE SReadType
   Book        AS STRING * 2
   Seq         AS STRING * 7
   CustName    AS STRING * 30
   SERVADDR    AS STRING * 20
   CurrRead    AS STRING * 10
   ReadDate    AS STRING * 10
   CrLf        AS STRING * 2
END TYPE

TYPE SchlumHHType
  'ne=No Equivalent, r=Required, o=Optional
  'r=right justified l=left justified z=zero filled s=space filled
  Route        AS STRING * 10  'r       uBook-Sequence   LS
  WalkSeq      AS STRING * 4   'r
  PageNum      AS STRING * 4   'r  18
  ReadSeq      AS STRING * 2   'r
  HHID         AS STRING * 6   'o  26
  ReadDir      AS STRING * 1   'o ne
  NumDial      AS STRING * 1   'o ne
  IDExpected   AS STRING * 13  'r meter serial#    43
  IDCaptured   AS STRING * 13  'o
  IDOverride   AS STRING * 13  'o                  67
  Decimals     AS STRING * 1   'o ne               68
  MtrRead      AS STRING * 10  'o                  78
  ReadOVRide   AS STRING * 10  'o ne               88
  HighLimit    AS STRING * 10  'o                  98
  LowLimit     AS STRING * 10  'o                  108
  Date2Read    AS STRING * 6   'o
  Date2Exp     AS STRING * 6   'o
  NoteCodes    AS STRING * 8   'o ne
  LocatCode    AS STRING * 2   'o
  MtrRCode     AS STRING * 2   'o
  RecType      AS STRING * 2   'o   'reading status
  RecStatus    AS STRING * 1   'o ne internal debugging
  ReadDate     AS STRING * 6   'o must have  Actual read date
  ReadTime     AS STRING * 6   'o ne
  ReadType     AS STRING * 1   'o Reading class
  NetNumb      AS STRING * 2   'o ne
  ReadAtmpt    AS STRING * 1   'o ne
  UserChar     AS STRING * 7   '????? no code in schlum manual
  HHManufac    AS STRING * 1   'o ne
  ActStatus    AS STRING * 1   'o Account Status  A/I
  MtrType      AS STRING * 1   'o Type for automatic read
  ReadFailCode AS STRING * 1   'o Reading Failure code
  PrevRead     AS STRING * 10  'o Previous Mtr reading
  PrevDate     AS STRING * 6   'o Previous read date
  HHDisp1      AS STRING * 24  'r Service Addr
  HHDisp2      AS STRING * 24  'r Remote Location (Book-Sequence)
  HHDisp3      AS STRING * 24  'r More Location info
  HHDisp4      AS STRING * 24  'r Other Info (Maybe name)
  Notes1       AS STRING * 24  'o Notes Fields
  Notes2       AS STRING * 24  'o Notes Fields
  Notes3       AS STRING * 24  'o Notes Fields
  Notes4       AS STRING * 24  'o Notes Fields
  Notes5       AS STRING * 24  'o Notes Fields
  Notes6       AS STRING * 24  'o Notes Fields
  Notes7       AS STRING * 24  'o Notes Fields
  Notes8       AS STRING * 24  'o Notes Fields
  OpCode       AS STRING * 1   'r Reserved
  UBAcctNo     AS STRING * 6   'Sosoft Reserved Account Number
  MtrSlot      AS STRING * 1   'Sosoft Reserved Meter Slot
  UtilFld      AS STRING * 33  'o For us to use
  CrLf         AS STRING * 2   'r CrLf terminator
END TYPE

' These Sensus Layout Files are Spec'd Out Exactly to Gilbert SC

TYPE UBGilSensusReadRecType         ' File Layout for Sending Out Records
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
    Message        AS STRING * 30
'040604
    MtrIDMST       AS STRING * 10
    MtrIDNO        AS STRING * 1
    MtrLat         AS STRING * 11
    MtrLng         AS STRING * 11
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
END TYPE

'10/18/04 Old ESensus type structure.
TYPE UBOESensusReadRecType         ' File Layout for Sending Out Records
    CustLastName AS STRING * 25
    CustFirstName AS STRING * 25
    MeterID AS STRING * 8
    Account AS STRING * 8
    LowRead AS STRING * 8
    HighRead AS STRING * 8
    SensusType AS STRING * 1        ' B=Touch Read : M=Manual
    PastRead AS STRING * 8
    CurRead AS STRING * 8
    ServAddress AS STRING * 20
    LocationNumber AS STRING * 10
    Message AS STRING * 30
END TYPE

TYPE UBOESensusGetReadRecType      ' File Layout For Retreiving Records
    CustLastName AS STRING * 25
    CustFirstName AS STRING * 25
    MeterID AS STRING * 8
    Account AS STRING * 8
    LowRead AS STRING * 8
    HighRead AS STRING * 8
    SensusType AS STRING * 1        ' B=Touch Read : M=Manual
    PastRead AS STRING * 8
    CurRead AS STRING * 8
    ServAddress AS STRING * 20
    LocationNumber AS STRING * 10
    fil1           AS STRING * 1
    ReadDate       AS STRING * 6
    Message AS STRING * 19   'was 26
    DateRead AS STRING * 4
END TYPE

'This is for temporary Tax bill for Spruce Pine
TYPE TxBillDefaultsType
    TxtHead1 AS STRING * 50
    TxtHead2 AS STRING * 50
    txtOpt1 AS STRING * 40
    TxtOpt2 AS STRING * 40
    TxtOpt3 AS STRING * 40
    TxtOpt4 AS STRING * 40
    txtPgph0 AS STRING * 125
    txtPgph1 AS STRING * 125
    txtPgph2 AS STRING * 125
    txtPgph3 AS STRING * 125
    txtPgph4 AS STRING * 125
    txtPgph5 AS STRING * 125
    txtPgph6 AS STRING * 125
    txtPgph7 AS STRING * 125
    TxtOpt5 AS STRING * 75
    txtHead4 AS STRING * 40
    txtHead5 AS STRING * 40
    txtHead6 AS STRING * 40
    TxtOpt6 AS STRING * 45
    TxtOpt7 AS STRING * 75
    dologo AS INTEGER  '0 for no 1 for yes
END TYPE

TYPE BillOutServType
   ServText               AS STRING * 20
   ServAmt                AS STRING * 10
END TYPE

TYPE BillOutRecType   '
   AcctNo                 AS STRING * 8
   LocationNum            AS STRING * 9        'format ##-######
   CustName               AS STRING * 35
   ADDR1                  AS STRING * 35
   ADDR2                  AS STRING * 35
   SERVADDR               AS STRING * 35
   CITY                   AS STRING * 18
   STATE                  AS STRING * 2
   ZIPCODE                AS STRING * 10
   BillType               AS STRING * 1        'N=Normal F=Final
   DepAppAmt              AS STRING * 10       'if final applied deposit amt.
   PrevDue                AS STRING * 15       'total revenue amount
   CurrDue                AS STRING * 15       'total revenue amount
   TotalDue               AS STRING * 15       'total revenue amount
   MtrType                AS STRING * 1        'w=water s=sewer c=combined
   MtrUnit                AS STRING * 1
   CurrDate               AS STRING * 10
   PrevDate               AS STRING * 10
   ServDays               AS STRING * 4
   CurrRead               AS STRING * 10       'Last/Current meter readings
   PrevRead               AS STRING * 10       'Previous readings
   Consump                AS STRING * 10
   IServDays              AS STRING * 4
   ICurrRead              AS STRING * 10       'Last/Current meter readings
   IPrevRead              AS STRING * 10       'Previous readings
   IConsump               AS STRING * 10
   ServInfo(1 TO 15)      AS BillOutServType
   BillDate               AS STRING * 10
   PastDueDate            AS STRING * 10
   DraftDate              AS STRING * 10
'mod for Johnston CO.
   MsgLine1               AS STRING * 22     'was 25
   MsgLine2               AS STRING * 22     'was 25
   MsgLine3               AS STRING * 22     'was 25
   MsgLine4               AS STRING * 22     'was 25
   MtrNumb                AS STRING * 12     'Added Meter Number
   CrLf                   AS STRING * 2
END TYPE



'Type UBPaymentRecType
'  OPERNUM         As Integer
'  PAYDATE         As Integer
'  CUSTACCT        As Long
'  CUSTNAME        As String * 24
'  CUSTADDR        As String * 24
'  CUSTCMNT        As String * 32
''052598 Added tax exempt flag changed cust comment len to 32
'  TaxExempt       As String * 1
'  AMTOWED         As Double
'  TENDERTY        As String * 12
'  CASHAMT         As Double
'  CHKAMT          As Double
'  AMTRECD         As Double
'  CHANGE          As Double
'  DESC            As String * 19
'  PaidOwed(1 To 15)   As PaidOwedType
'  TOTOWED         As Double
'  AMTPAID         As Double
''112801 Added cust status changed desc to 19
'  Status          As String * 1
'End Type
'
'
'
'Type PayListType
'  CustRec   As Long
'  ListRec   As Long
'End Type
'
'Type BookTotalType
'  Count   As Long
'  Cash    As Double
'  Check   As Double
'End Type

TYPE LockBoxRecType
  AcctNum     AS STRING * 6     '1-6
  Amount      AS STRING * 10    '7-16
  TenderType  AS STRING * 2     '17/18 CH for Check, CA for cash
  ChkNum      AS STRING * 8     '19/26 chk num
  PaymentDate AS STRING * 10    '27-36 01/01/1999
  Fill1       AS STRING * 44    '37/80
  CrLf        AS STRING * 2     '81/82 CrLf
END TYPE

'First Citizens
TYPE LockBoxRecTypeFC
  AcctNum     AS STRING * 8     '1-6   Utility Billing Account Number
  Amount      AS STRING * 10    '7-16  Decimal assumed (2.00 = 200) in record
  TenderType  AS STRING * 2     '17-18 CH for Check, CA for cash
  ChkNum      AS STRING * 8     '19-26 Check Number
  PaymentDate AS STRING * 10    '27-36 Payment Date. Formate: MM-DD-YYYY
  Blank       AS STRING * 42    '37-80 N/A
  CrLf        AS STRING * 2     '81-82 CR+LF
END TYPE

