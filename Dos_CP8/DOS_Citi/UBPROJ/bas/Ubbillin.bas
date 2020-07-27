Attribute VB_Name = "Module2"
Option Explicit

DefInt A-Z
  
Const MaxRevsCnt = 15

Type oRevSetUpType
    REVNAME As String * 15
    UseDep   As String * 1
    USERATE  As String * 1
    TAXRATE  As Single
    UseMtr   As String * 1
    DistOr   As Integer
End Type

Type SetUpAcctType
   REVNAME    As String * 15
   DebitAcct  As String * 14
   CreditAcct As String * 14
End Type

Type RevSetUpType
    REVNAME As String * 15
    UseDep   As String * 1
    USERATE  As String * 1
    TAXRATE  As Single
    UseMtr   As String * 1
    DistOr   As Integer
    ProRate  As String * 1
End Type

Type oUBSetupRecType
    UTILNAME        As String * 35
    DEFCITY         As String * 18
    DEFSTATE        As String * 2
    ZIPCODE         As String * 10
    PreByBook       As String * 1
    RecpPort        As String * 1
    RECPDEFT        As String * 1
    EstRead         As String * 1
    BANKDFT         As String * 1
    UseSeq          As String * 1
    BILLCYCL        As String * 1
    DefLook         As String * 1
    MethAcct        As String * 1      'new 02-14-97
    SkipInactive    As String * 1
    SkipSeparator   As String * 1
    Make99File      As String * 1
    LowRead         As Integer
    HighRead        As Integer
    HHDEVICE        As String * 1    'P=PC3000 S=Sensus C=Syscom R=Radix N=None
    Revenues(1 To 15) As oRevSetUpType
    BillAcct(1 To 15) As SetUpAcctType
    PayAcct(1 To 15)  As SetUpAcctType
    DepAcct(1 To 15)  As SetUpAcctType
End Type

Type UBSetupRecType
    UTILNAME        As String * 35
    DEFCITY         As String * 18
    DEFSTATE        As String * 2
    ZIPCODE         As String * 10
    PreByBook       As String * 1
    RecpPort        As String * 1
    RECPDEFT        As String * 1
    EstRead         As String * 1
    BANKDFT         As String * 1
    UseSeq          As String * 1
    BILLCYCL        As String * 1
    DefLook         As String * 1
    MethAcct        As String * 1      'new 02-14-97
    SkipInactive    As String * 1
    SkipSeparator   As String * 1
    Make99File      As String * 1
    LowRead         As Integer
    HighRead        As Integer
    HHDEVICE        As String * 1    'P=PC3000 S=Sensus C=Syscom R=Radix N=None
    Revenues(1 To 15) As RevSetUpType
    BillAcct(1 To 15) As SetUpAcctType
    PayAcct(1 To 15)  As SetUpAcctType
    DepAcct(1 To 15)  As SetUpAcctType
End Type

Type UBSensusRecType
    PathWay As String * 20
End Type

Type UBLogiconRecType
    PathWay As String * 20
End Type

Type UBPC3000RecType
    PathWay As String * 20
End Type

Type DistArrayType
   DistOrder As Integer
   DistCnt   As Integer
End Type

' These Sensus Layout Files are Spec'd Out Exactly to Long View NC

Type LUBSensusReadRecType        ' File Layout for Sending Out Records
    ServAddress  As String * 20
    MeterID      As String * 8
    LowRead      As String * 8
    HighRead     As String * 8
    Account      As String * 8
    SensusType   As String * 1        ' B=Touch Read : M=Manual
'    CustName     AS STRING * 25
'    SerialNumb   AS STRING * 8        'Added Per Mickey on 6-23-97
End Type

Type UBSensusReadRecType        ' File Layout for Sending Out Records
    ServAddress  As String * 20
    MeterID      As String * 8
    LowRead      As String * 8
    HighRead     As String * 8
    Account      As String * 8
    SensusType   As String * 1        ' B=Touch Read : M=Manual
    CUSTNAME     As String * 25
    SerialNumb   As String * 8        'Added Per Mickey on 6-23-97
End Type

Type UBLogiconReadRecType
 RecType   As String * 1
 RouteNo   As String * 2
 AcctNo    As String * 6
 RecName   As String * 30
 ServAddress As String * 25
 ReadDate    As String * 6
 ReadTime    As String * 6
 Consumption As String * 8
 PrevRead    As String * 8
 CurRead     As String * 8
 LowRead     As String * 8
 HighRead   As String * 8
 MtrNumb    As String * 12
 CountChg   As String * 1
 ForceFlag  As String * 1
 ReportCode As String * 2
 Remark     As String * 40
 Label      As String * 19
 PrintFlag  As String * 1
 MessageOut As String * 30
 Book       As String * 2
 Future     As String * 29
 Recend     As String * 1               'Must be 'X'
 CRLF       As String * 2
End Type


Type UBLogiconGetReadRecType
 RecType   As String * 1
 RouteNo   As String * 2
 AcctNo    As String * 6
 RecName   As String * 30
 ServAddress As String * 25
 ReadDate    As String * 6
 ReadTime    As String * 6
 Consumption As String * 8
 PrevRead    As String * 8
 CurRead     As String * 8
 LowRead     As String * 8
 HighRead   As String * 8
 MtrNumb    As String * 12
 CountChg   As String * 1
 ForceFlag  As String * 1
 ReportCode As String * 2
 Remark     As String * 40
 Label      As String * 19
 PrintFlag  As String * 1
 MessageOut As String * 30
 Book       As String * 2
 Future     As String * 29
 Recend     As String * 1               'Must be 'X'
 CRLF       As String * 2
End Type

Type UBSensusGetReadRecType     ' File Layout For Retreiving Records
    Account As String * 8
    NotUsed As String * 5
    MeterID As String * 8
    Filler As String * 1
    Reading As String * 8
    NotUse1 As String * 1
    DateRead As String * 6
    NotUse2 As String * 4
End Type



Type LUBSensusGetReadRecType     ' LONGVIEW File Layout For Retreiving Records
    Account As String * 12
    MeterID As String * 8
    Reading As String * 8
    DateRead As String * 4
    NotUse2 As String * 2    'CRLF
End Type

Type UBPC3000ReadRecType           'File Layout for Sending Out Records
  CUSTNAME      As String * 20
  ServAddress   As String * 20
  MeterID       As String * 8
  LowRead       As Double
  HighRead      As Double
  Account       As String * 8
  ReadFlag      As String * 1        'Y/N
  MeterType     As String * 1
  Book          As Integer
  CurRead       As Double
  PastRead      As Double
  ReadDate      As Integer
  ReadTime      As String * 5
  Note1         As String * 20
  Note2         As String * 20
  Note3         As String * 20
  NoteStatus    As String * 1    'T=Temp Note  P=Perm Note
End Type

Type UBDGRecType
    PathWay As String * 20
End Type

Type UBDGProcRecType        ' File Layout for Sending Out Records
    RouteID As String * 20
    SvcTyp As String * 1
    CUSTNAME As String * 25
    SvcLoc As String * 21
    MeterSN As String * 20
    MeterType As String * 1       ' C for reg mtr   D for demand elec
    High As String * 10
    Low As String * 10
    Msg As String * 110
    Account As String * 10
    NewRdng As String * 10
    NewDmnd As String * 10
    Date As String * 6
    Time As String * 6
    NewAcctRte As String * 20
End Type

'Itron Layouts

Type UBItronRecType
    PathWay As String * 20
End Type

Type ItronFType                         'Header Record Type
    RecordCode As String * 1
    Route As String * 8
    Message As String * 64
    Filler As String * 5
    CRLF As String * 2
End Type

Type ItronAType                         'Customer Record One
    RecordCode As String * 1
    Route As String * 8
    AcctNumb As String * 10
    Geo As String * 12
    SEQNUMB As String * 5
    Message As String * 1
    AType As String * 1
    Filler As String * 40
    CRLF As String * 2
End Type

Type ItronBType                         'Customer Record Two
    RecordCode As String * 1
    CUSTNAME As String * 32
    CustAddr As String * 32
    Filler As String * 13
    CRLF As String * 2
End Type

Type ItronHType
    RecordCode As String * 1
    MeterNumb As String * 12
    Dials As String * 1
    LowRead As String * 8
    HighRead As String * 8
    LastRead As String * 8
    MeterType As String * 1
    Multiplier As String * 6
    NoMths As String * 1
    MtrMake As String * 2
    DispCode As String * 2
    NumbDec As String * 1
    MustRead As String * 1
    Status As String * 1
    Filler As String * 25
    CRLF As String * 2
End Type

Type ItronZType
    RecordCode As String * 1
    Route As String * 8
    NumberAccts As String * 4
    NumberMeters As String * 4
    Filler As String * 61
    CRLF As String * 2
End Type

'Itron Read Layouts
Type ItronCType
     RecordCode As String * 1           'Must be a C
     Route      As String * 8
     Acct       As String * 10
     SpecInst   As String * 2
     Survey     As String * 1
     ReadStatus As String * 1
     ReSeqFlag  As String * 1
     ReadDate   As String * 4           'mmdd
     AcctChg    As String * 1
     Filler     As String * 51
End Type

Type ItronDType
     RecordCode As String * 1           'Must be a D
     MeterNum   As String * 12
     LocCode    As String * 2
     MtrCon1    As String * 2
     MtrCon2    As String * 2
     Mult       As String * 6
     ChgeFlag   As String * 1
     Filler     As String * 54
End Type

Type ItronEType
     RecordCode As String * 1           'Must be a E
     NbrDials   As String * 1
     CurReading As String * 8           'Zero's if No Read
     DecPlaces  As String * 1
     ReadTime   As String * 6
     ReadChg    As String * 1
     DispCode   As String * 2
     ReadCount  As String * 1
     ReadVerify As String * 1
     NoReadCode As String * 2
     Filler     As String * 56
End Type

Type HHCodeRecType
  HHCRec  As Integer
  HHCode  As String * 20
End Type
'**********************


Type RePrintIndexType
  BillNum  As Long
  BillRec  As Long
End Type

Type UBXferInfoType
  DAcctNo   As String * 14
  DebitAmt  As Double
  DRecNo    As Integer
  DTitle    As String * 30
  CAcctNo   As String * 14
  CreditAmt As Double
  CRecNo    As Integer
  CTitle    As String * 30
End Type

Type GJXferRecType
  RevText    As String * 15
  BAcctInfo  As UBXferInfoType     'Billing Accounts
  PAcctInfo  As UBXferInfoType     'Payment Accounts
  DAcctInfo  As UBXferInfoType     'Deposit Accounts
End Type

'Note:  if transaction is an adjustment then
'       CurRead field will contain the adjust amount
Type UBTransRecType
   TransDate              As Integer      '
   TransType              As Integer      '
   TransDesc              As String * 21  'may change
   TransAmt               As Double       'total revenue amount
   RevAmt(1 To 15)        As Double       'Revenue amounts
   TaxAmt(1 To 15)        As Single       'Tax Amounts
'01-20-97 Added meter types field to hold meter type at time of transaction
   MtrTypes(1 To 7)       As Integer
'*******************
   CurRead(1 To 7)        As Long         'Last/Current meter readings
   PrevRead(1 To 7)       As Long         'Previous readings
   EstRead(1 To 7)        As String * 1   'Y/N Flags for meter est's
   BillNumber             As Long         'Number on the bill that Printed
   ReadDate               As Integer
   BillDate               As Integer
   PastDueDate            As Integer
   DraftDate              As Integer      'mowasa & plymouths bills. Can be changed
'111398
   ProRatePCT             As Integer
   ChkByte                As String * 1   'Added check byte
   EPPFlag                As String * 1   'Equal Payment Flag
   CustStatus             As String * 1   'Customer Status at Time of Transaction
'020199
   EPPTrans               As Long         'Pointer to Equal Pay trans
   PenAtBill              As Single
   'Filler2                AS STRING * 4
'****************
   PayTypeCode            As Integer      'Payment Type:  1=Cash, 2=Check
   OperatorNumber         As Integer      '
   CustAcctNo             As Long         'Pointer to RecNo in ubcust.dat
   PrevTrans              As Long
   ActUsage               As Long         'Changed for wadesboro
   ActiveFlag             As Integer      'Valid transaction flag
   RunBalance             As Double
   CheckAmount            As Double
   CashAmount             As Double
   BillMsg                As String * 20
   ApplyDepFlag           As String * 1
   Posted2GL              As String * 1
   PrevDate               As Integer
   PenalFlag              As String * 1
   TaxExempt              As String * 1
   NONProfit              As String * 1
End Type

'Trans Types
Const TranUtilityBill = 1          '   1=Utility bill
Const TranLateCharge = 2           '   2=late charge      'NOT USED
Const TranReconnectFee = 3         '   3=reconnect fee    'NOT USED
Const TranBillPayment = 4          '   4=Bill Payment
Const TranAppliedDeposit = 5       '   5=Applied Deposit
Const TranPenaltyCharge = 6        '   6=Penalty Charge
Const TranDepositPayment = 7       '   7=Deposit Payment
Const TranDraftPayment = 8         '   8=Draft Payment
Const TranRefundDeposit = 9        '   9=Refund Deposit
Const TranBeginBalance = 10        '  10=Beginning Balance
Const TranUpwardAdjustment = 11    '  11=Bill Adjustments
Const TranDownwardAdjustment = 12  '  12=Bill Adjustments
Const TranMiscPayment = 99         '  99=Misc Payment

Const MtrWaterOnly = 1
Const MtrSewerOnly = 2
Const MtrCombined = 3
Const MtrElectric = 4
Const MtrDemand = 5
Const MtrGas = 6
Const MtrTouchRead = 7
Const MtrLightsService = 8

Type MessLineType
  Line As String * 59
  LineDate As String * 10
End Type

Type UBMessRecType
  MessLine(1 To 15) As MessLineType
  CustRec As Long
End Type
'**********************

Type oUBDraftRecType
    BANKDEST As String * 9
    BANKORIG As String * 9
    BANKNAME As String * 23
    BANKLOC  As String * 23
    FEDID    As String * 9
    FEDPREFX As String * 1
End Type

Type oUBDraftRecType2
    BANKDEST As String * 9
    BANKORIG As String * 9
    BANKNAME As String * 23
    BANKLOC  As String * 23
    COMPACCT As String * 20
    FEDID    As String * 9
    FEDPREFX As String * 1
End Type

Type UBDraftRecType
    BANKDEST As String * 9
    BANKORIG As String * 9
    BANKNAME As String * 23
    BANKLOC  As String * 23
    COMPACCT As String * 20
    FEDID    As String * 9
    FEDPREFX As String * 1
    FileName As String * 12
End Type

Type UBDraftPayRecType
    CustAcctNum   As Long
    DraftAmt      As Double
End Type

Type UBDraftRecord1Type
    Field1 As String * 1        ' Record Type Code Must = 1
    Field2 As String * 2        ' Priority Code Must = 01
    Field3 As String * 10       ' Immediate Destination Bank Transit Number (BB&T=b053101121 where b=blank space) Right Justified
    Field4 As String * 10       ' Immediate Origin Bank Transit Number Must be Right Justified
    Field5 As String * 6        ' Transmission File Creation Date (yymmdd)
    Field6 As String * 4        ' Transmission File Creation Time (hhmm)
    Field7 As String * 1        ' Field Modifier ID Must = A
    Field8 As String * 3        ' Record Size  Must = 094
    Field9 As String * 2        ' Blocking Factor  Must = 10
    Field10 As String * 1       ' Format Code  Must = 1
    Field11 As String * 23      ' Name of Destination Bank (Receiving Transmission)
    Field12 As String * 23      ' Name of Originating Bank
    Field13 As String * 8       ' Reserved Needs to be 8 blanks padded
End Type

Type UBDraftRecord5Type
    Field1 As String * 1        ' Record Type Code Must = 5
    Field2 As String * 3        ' Service Class Code Must = 200
    Field3 As String * 16       ' Company Submitting Name in ALL CAPS
    Field4 As String * 20       ' Discretionary Data
    Field5 As String * 10       ' Company ID (Federal Tax ID Number)
    Field6 As String * 3        ' Standard Entry Class (PPD for Direct Deposits and Drafts)
    Field7 As String * 10       ' Company Entry Description SUCH AS UTIL BILL
    Field8 As String * 6        ' Company Desc Date YYMMDD
    Field9 As String * 6        ' Effective Entry Date YYMMDD
    Field10 As String * 3       ' RESERVED LEAVE WITH 3 BLANKS
    Field11 As String * 1       ' Must Equal 1 for Originator Status Code
    Field12 As String * 8       ' Originating Fin. Inst. ID  05310112 for BB&T
    Field13 As String * 7       ' Batch Number Beginning with 0000001
  End Type

Type UBDraftRecord6Type
    Field1 As String * 1        ' Record Type Code  Must = 6
    Field2 As String * 2        ' Transaction Code      22 Credit Checking
                                '                       27 Debit Checking
                                '                       32 Credit Savings
                                '                       37 Debit Savings
                                '                       28 PRENOTE DEBIT CHECKING  (Amt Must be all zeros
                                '                       Generally Add 1 to Get PreNote Transaction Code
    Field3 As String * 8        ' Individual's Bank ID Transit Routing #
    Field4 As String * 1        ' Transit Routing Check Digit
    Field5 As String * 17       ' Individual's Bank Account Number
    Field6 As String * 10       ' Amount (assume 2 decimal places)
    Field7 As String * 15       ' Individual's ID Number to Customer (Usually Customer Utility Account Number)
    Field8 As String * 22       ' Individual's Name
    Field9 As String * 2        ' Set to 2 spaces (Not Used)
    Field10 As String * 1       ' Set to '0' to signify no addenda records
    Field11 As String * 15      ' Trace Number
                                ' Consists of Bank ID # 05310112 plus
                                ' Line Item Number starting w/ 0000001
                                ' and incrementing once for each line (Record6)
End Type


Type UBDraftRecord8Type
    Field1 As String * 1        ' Record Code Must = 8
    Field2 As String * 3        ' Service Class Code  Must = 200
    Field3 As String * 6        ' Number of Detail (TYPE 6) Records
    Field4 As String * 10       ' Hash Total
                                ' Hash#=Hash#+val(banktransit#) for Each Type 6 Record
    Field5 As String * 12       ' Total Debit Amount
    Field6 As String * 12       ' Total Credit Amount
    Field7 As String * 10       ' Federal ID Tax Number
    Field8 As String * 19       ' RESERVED KEEP BLANK
    Field9 As String * 6        ' RESERVED BY FEDERAL RESERVE BANK
    Field10 As String * 8      ' Originating Financial Inst. 05310112 for BB&T
    Field11 As String * 7       ' Batch # Beginning with 0000001
End Type

Type UBDraftRecord9Type
    Field1 As String * 1        ' Record Code Must=9
    Field2 As String * 6        ' Batch Count (Sum of Batches) NORMALLY 000001
    Field3 As String * 6        ' Block Count Number of Records
                                ' Found by taking Total Size of File and Dividing ty 940
    Field4 As String * 8        ' Total 06 Record Type Entries
    Field5 As String * 10       ' Enter Hash  See Above
    Field6 As String * 12       ' Total Debit Entry Dollar Amount (Assume 2 decimal)
    Field7 As String * 12       ' Total Credit Entry Dollar Amount (Assume 2 decimal)
    Field8 As String * 39       ' RESERVED FOR FUTURE USE
End Type


'020299
Const NameIndexFile = "UBCUSTNM.IDX"
Const BookIndexFile = "UBCUSTBK.IDX"
Const TempIndexName = "UBTEMP.IDX"
Const ServiceAddressIndexFile = "UBSVCADD.IDX"

Const UBBillsFile = "UBBILLS.DAT"
Const UBIBillFile = "UBIBILL.DAT"
Const UBFinPreRptFile = "UBPREFIN.RPT"
Const UBFinBillsFile = "UBFBILLS.DAT"
Const RePrintIdxFile = "UBREPRNT.IDX"

Type PSAZipIndexType
  ZIPCODE  As String * 10
  SName    As String * 10
  RecNum   As Integer
  pad      As String * 10
End Type

Type MOWZipIndexType
  ZIPCODE  As String * 10
  RecNum   As Integer
  FillPad As String * 4
End Type

Type UBPostalIndexType
  ZIPCODE  As String * 10
  Route    As String * 4
  RecNum   As Integer
End Type

Type UBServiceAddressIndexType
  ServiceAddress  As String * 14
  RecNum   As Integer
End Type

Type UBSequenceIndexType
  SeqNumber As Long
  RecNum    As Integer
  Fill     As String * 10      'This is to fill this to a 16 byte boundary
End Type

Type UBCustIndexRecType
  RecNum As Long
End Type

Type oUBCustReIndexRecType
  SearchName As String * 10
  DelFlag    As Integer
  RecNum     As Long
End Type

Type nUBCustReIndexRecType
  SearchName As String * 10
  DelFlag    As String * 1
  Status     As String * 1
  RecNum     As Long
End Type

Type oUBCustReIndexRecType1
  SearchName As String * 10
  First      As String * 2
'  DelFlag    AS STRING * 1
'  Status     AS STRING * 1
  RecNum     As Long
End Type

Type UBLocaReIndexRecType
  Book       As String * 2
  SEQNUMB    As String * 6
  RecNum     As Long
  pad        As String * 4
End Type

Type UBPINType
  PIN As Long
End Type

Type BookSeqRecType
  BookSeq  As Long
End Type

Type ServicesType
    RATECODE As String * 4
    RMtrType As String * 1
End Type

Type FlatRateType
    FRDESC   As String * 18
    FRAMT    As Double
    FRFREQ   As String * 1
    REVSRC   As Integer
    NumMin   As Integer
End Type

Type RevDataType
    REVNAME    As String * 20
    RATECODE   As String * 4
    RevMtrType As String * 1
End Type

Type LocMeterType
    MTRNUM    As String * 12
    MTRMulti  As Integer
    MTRType   As String * 1
    MTRUnit   As String * 1
    NumUser   As Integer
    InsDate   As Integer
    CurRead   As Long
    PrevRead  As Long
    CurDate   As Integer
    PastDate  As Integer       'hidden & protected
    ReadFlag  As String * 1    'hidden & protected
    AvgUse    As Long          'hidden & protected
    UseCnt    As Integer       'hidden & protected
End Type

Type MonthlyPayType
    AMTOWED      As Double
    TotAmtPD     As Double
    PayAmt       As Double
    RevSource    As Integer
End Type

Type NewUBCustRecType
    Book          As String * 2
    SEQNUMB       As String * 6
    Status        As String * 1
    OPENDATE      As Integer
    SEARCH        As String * 10
    CUSTNAME      As String * 35
    ADDR1         As String * 35
    ADDR2         As String * 35
    SERVADDR      As String * 35
    CITY          As String * 18
    STATE         As String * 2
    ZIPCODE       As String * 10
    HPHONE        As String * 14
    WPHONE        As String * 14
    SOSEC         As String * 11
    DRVLIC        As String * 16
    CUSTTYPE      As String * 3
    Addr911       As String * 14
'051498 added bill to field. Removed 1 byte from 911 addr
    BillTo        As String * 1
'********************************************************
    BILLCOPY      As Integer
    POSTRTE       As String * 4
    BILLCYCL      As Integer
    ZONE          As String * 3
    SEQ           As Long
'Page 2
    CASHONLY      As String * 1
    LATEFEE       As String * 1
    CUTOFFYN      As String * 1
    TAXEXPT       As String * 1
    SRCIT         As String * 1
    EPPFlag       As String * 1
'032299 Modified for Bank draft account type
'    EPPAMT        AS DOUBLE
    Filler1       As String * 7
    USEDRAFT      As String * 1
    AcctType      As String * 1
'032299 Inserted account type
    BANKNAME      As String * 34
    BANKLOC       As String * 30
    TRANSIT       As String * 9
    BANKACCT      As String * 20
    BILLCMNT      As String * 25
    PAYCMNT       As String * 25
    PumpCode      As String * 4
    USERCODE1     As String * 4
    USERCODE2     As String * 2
    ProRatePCT    As Integer
    HHMSG1        As String * 20
    HHMSG2        As String * 20
    HHMSG3        As String * 20
'Page 3
    Serv(1 To 15)      As ServicesType
    FlatRates(1 To 4)  As FlatRateType
'Page 4
    Monthly(1 To 2)    As MonthlyPayType
    MFEE1         As Double
    MFEE2         As Double
    LocMeters(1 To 7)  As LocMeterType
'END OF Quick Screen Form
    CustPIN       As Long
    LastTrans     As Long
    CurrBalance   As Double
    PrevBalance   As Double
    CurrRevAmts(1 To 15) As Double
    PrevRevAmts(1 To 15) As Double
      'Filler2       AS STRING * 120
    DepositAmt    As Double
    DelFlag       As Integer
    PreNoteFlag   As Integer
    WOLastTrans   As Long            'work order last trans pointer
    EstFlag       As String * 1
    MessageRec    As Long            ' Points to Message Record
    OldRec        As Long
    EPPLastTran   As Long
    NewNotes      As Integer
    FillPad       As String * 4
    ChkByte       As String * 1
End Type


Type WrkOrdTextType
  Text(1 To 6)  As String * 67
End Type

Type WorkOrderRecType
    CustRec           As Long
    ENTRYDATE         As Integer
    OrdersText        As WrkOrdTextType
    RepliesText       As WrkOrdTextType
    CompleteByDate    As Integer
    CompletedDate     As Integer
    PrevTransRec      As Long
End Type

Type Newport
    Acct As String * 6
    Name As String * 31
    Address As String * 25
    CITY As String * 15
    ST As String * 2
    Zip As String * 5
    ServAddress As String * 20
    Source As String * 1
    Ctype As String * 1
    CLoc As String * 1
    Blk As String * 2
    Garb As String * 1
End Type

Type TblBreakRecType
    UNITS      As Long
    UNITAMT    As Double
End Type

Type UBRateTblRecType
    RATECODE As String * 4
    RATEDESC As String * 29
    ChkByte  As String * 1
    MINAMT   As Double
    MINUNITS As Long
    MaxAmt   As Double
    TblBreaks(1 To 10) As TblBreakRecType
End Type

Type oUBRateTblRecType
    RATECODE As String * 4
    RATEDESC As String * 30
    MINAMT   As Double
    MINUNITS As Long
    DiscPct  As Integer
    TblBreaks(1 To 10) As TblBreakRecType
End Type

Type BookConsumpType
    Book             As Integer
    CustCnt          As Long
    Consump(1 To 15) As Double
    RevAmt(1 To 15)  As Double
    TaxAmt(1 To 15)  As Double
End Type

Type PumpConsumpType
    PumpCode         As String * 4
    CustCnt          As Long
    Consump          As Double
End Type

Public Function ErrorScrn(WhatError%, Acct&)
  
'  GoTo OutOfHere
'  'ErrorScrn = True
'
'  AcCol = 22
'  ReDim TempArray(0) As Integer
'  'SaveScrn TempArray()
'
'  'HideCursor
'  'BlockClear
'  'DisplayUBScrn "ERRSCRN1"
'  'ErrorScrn = False
'' ' Exit Function
'
'  Select Case WhatError
'  Case 1
'    'QPrintRC "Has Invalid Reading!", 10, 37, -1
'  Case 2
'    'QPrintRC "Invalid Book Number!", 10, 37, -1
'  Case 3
'    AcCol = 21
'    'QPrintRC "Has an INVALID RATE CODE!!", 10, 35, -1
'  Case 4
'    'QPrintRC "Has Mismatched Meters!", 10, 37, -1
'  Case 5
'    'QPrintRC "Has an INVALID Reading!", 10, 37, -1
'  Case 6
'    'QPrintRC "INVALID Flat Rate Info!", 10, 37, -1
'  Case 7
'    'QPrintRC "INVALID Monthly Billed Code!", 10, 35, -1
'  Case 8
'    'QPrintRC "Meters with NO RATE Code!", 10, 36, -1
'  Case 9
'    'QPrintRC "Invalid Customer Type!", 10, 36, -1
'  End Select
'  'QPrintRC "ACCOUNT:" + Str$(Acct&), 10, AcCol, -1
'  'QPrintRC "Correct and Print Again.", 13, 28, -1
'
'  'ShowCursor
'  'Get.Moose.OR.Key Ky$, MooseButton%, MRow%, MCol%
'
'  If Len(Ky$) = 2 Then
'    If Right$(Ky$, 1) = "g" Then
'      ErrorScrn = False
'      'LPRINT Acct&
'    End If
'  End If
'  'RestScrn TempArray()
'  Erase TempArray
'
OutOfHere:
End Function
  
Function GetNumOfRevs%()
  Dim NumOfRevs As Integer, UBSetUpLen As Integer
  Dim RevCnt As Integer
  Dim TempRev As String
  
  NumOfRevs = 15
  ReDim UBSetUpRec(1) As UBSetupRecType
  UBSetUpLen = Len(UBSetUpRec(1))
  
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
'  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
  
  For RevCnt = 1 To 15
    TempRev$ = Trim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME)
    If Len(TempRev$) = 0 Then
      NumOfRevs = RevCnt - 1
      Exit For
    End If
  Next
  GetNumOfRevs = NumOfRevs
  Erase UBSetUpRec
End Function
  
Sub GetPreBillOrder(Choice, ExitFlag, SeqFlag$)
  
  Dim MaxLen As Integer
    
'  Choice = 2
'  Exit Sub
  
  ReDim MChoice$(1 To 6)
  
  MChoice$(1) = "Customer Name Order"
  MChoice$(2) = "Account Number Order"
  MChoice$(3) = "Location Number Order"
  MChoice$(4) = "Postal Carrier Route Order"
  MChoice$(5) = "ZipCode Order"
  
  If SeqFlag$ = "Y" Then
    MChoice$(6) = "Sequence Number Order"
  End If
  
  
  'Restart:
  '--Center Menu within Screen
  
'  Do            '--Set upper left corner of menu, turn off the cursor
'    'LOCATE Row, Col, 0
'    BlockClear
'    TitleBox 2, Col, MaxLen + 3, "Pre-Billing Report ", Cnf
'    TitleBox 21, Col, MaxLen + 3, "Use " + Chr$(24) + "-" + Chr$(25) + " to select", Cnf
'    ShowCursor
'
'    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
'
'    If Ky$ = Chr$(27) Then
'      ExitFlag = True
'      Choice = 0
'    End If
'
'    Exit Do
'
'  Loop
'
  
End Sub
  
Static Function GetRevCharge#(RateTbl As UBRateTblRecType, TMeterConsp&, MeterMulti&)
  
  Dim MinBillAmt As Double, TAmt As Double
  Dim LastTblCnt As Integer, BCnt As Integer
  Dim MeterConsump As Long, UNITS  As Long
  
  'STOP
  
  MinBillAmt# = RateTbl.MINAMT
  
  If MinBillAmt# < -1000000 Then
    MinBillAmt# = 0
    TAmt# = -1
    GoTo GotTAmt
  End If
  
  LastTblCnt = 10
  For BCnt = 1 To 10
    If RateTbl.TblBreaks(BCnt).UNITAMT <= 0 Then
      LastTblCnt = BCnt - 1
      Exit For
    End If
  Next
  
  MeterConsump& = TMeterConsp&
  
  TAmt# = 0
  
  If LastTblCnt >= 2 Then
    If MeterConsump& >= RateTbl.TblBreaks(1).UNITS And MeterConsump& <= RateTbl.TblBreaks(2).UNITS Then
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(1).UNITS)
      'special patch for cave junction
      If UNITS& = 0 Then
        UNITS& = 1
      End If
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
      GoTo GotTAmt
    Else
      UNITS& = (RateTbl.TblBreaks(2).UNITS - RateTbl.TblBreaks(1).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
    End If
  Else          'no other rate breaks
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(1).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(1).UNITAMT))
    GoTo GotTAmt
  End If
  
  'Break 2
  If LastTblCnt >= 3 Then
    If MeterConsump& > RateTbl.TblBreaks(2).UNITS And MeterConsump& <= RateTbl.TblBreaks(3).UNITS Then
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(2).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
      GoTo GotTAmt
    Else
      UNITS& = (RateTbl.TblBreaks(3).UNITS - RateTbl.TblBreaks(2).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
    End If
  Else
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(2).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(2).UNITAMT))
    GoTo GotTAmt
  End If
  
  'Break 3
  If LastTblCnt >= 4 Then
    If MeterConsump& >= RateTbl.TblBreaks(3).UNITS And MeterConsump& <= RateTbl.TblBreaks(4).UNITS Then
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(3).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
      GoTo GotTAmt
    Else
      UNITS& = (RateTbl.TblBreaks(4).UNITS - RateTbl.TblBreaks(3).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
    End If
  Else
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(3).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(3).UNITAMT))
    GoTo GotTAmt
  End If
  
  'Break 4
  If LastTblCnt >= 5 Then
    If MeterConsump& >= RateTbl.TblBreaks(4).UNITS And MeterConsump& <= RateTbl.TblBreaks(5).UNITS Then
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(4).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
      GoTo GotTAmt
    Else
      UNITS& = (RateTbl.TblBreaks(5).UNITS - RateTbl.TblBreaks(4).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
    End If
  Else
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(4).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(4).UNITAMT))
    GoTo GotTAmt
  End If
  
  'break 5
  If LastTblCnt >= 6 Then
    If MeterConsump& >= RateTbl.TblBreaks(5).UNITS And MeterConsump& <= RateTbl.TblBreaks(6).UNITS Then
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(5).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
      GoTo GotTAmt
    Else
      UNITS& = (RateTbl.TblBreaks(6).UNITS - RateTbl.TblBreaks(5).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
    End If
  Else
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(5).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(5).UNITAMT))
    GoTo GotTAmt
  End If
  
  'break 6
  If LastTblCnt >= 7 Then
    If MeterConsump& >= RateTbl.TblBreaks(6).UNITS And MeterConsump& <= RateTbl.TblBreaks(7).UNITS Then
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(6).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
      GoTo GotTAmt
    Else
      UNITS& = (RateTbl.TblBreaks(7).UNITS - RateTbl.TblBreaks(6).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
    End If
  Else
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(6).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(6).UNITAMT))
    GoTo GotTAmt
  End If
  
  'break 7
  If LastTblCnt >= 8 Then
    If MeterConsump& >= RateTbl.TblBreaks(7).UNITS And MeterConsump& <= RateTbl.TblBreaks(8).UNITS Then
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(7).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
      GoTo GotTAmt
    Else
      UNITS& = (RateTbl.TblBreaks(8).UNITS - RateTbl.TblBreaks(7).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
    End If
  Else
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(7).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(7).UNITAMT))
    GoTo GotTAmt
  End If
  
  'break 8
  If LastTblCnt >= 9 Then
    If MeterConsump& >= RateTbl.TblBreaks(8).UNITS And MeterConsump& <= RateTbl.TblBreaks(9).UNITS Then
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(8).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
      GoTo GotTAmt
    Else
      UNITS& = (RateTbl.TblBreaks(9).UNITS - RateTbl.TblBreaks(8).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
    End If
  Else
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(8).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(8).UNITAMT))
    GoTo GotTAmt
  End If
  
  'break 9
  If LastTblCnt >= 10 Then
    If MeterConsump& >= RateTbl.TblBreaks(9).UNITS And MeterConsump& <= RateTbl.TblBreaks(10).UNITS Then
      UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
      GoTo GotTAmt
    Else
      UNITS& = (RateTbl.TblBreaks(10).UNITS - RateTbl.TblBreaks(9).UNITS)
      TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
    End If
  Else
    UNITS& = (MeterConsump& - RateTbl.TblBreaks(9).UNITS)
    TAmt# = Round#(TAmt# + (UNITS& * RateTbl.TblBreaks(9).UNITAMT))
    GoTo GotTAmt
  End If
  
GotTAmt:
  GetRevCharge# = Round#(MinBillAmt# + TAmt#)
  
End Function
  
Sub MakeBillFile(AbortFlag As Boolean, FuelAdjAmt#, ThisCycle%, ThisBook%)
  
  Dim MeterType As String, Ctype As String
  Dim RATECODE As String
  
  Dim UBSetUpLen As Integer, ThisRevCnt As Integer
  Dim ElecRev As Integer, UBBillRecLen As Integer
  Dim UBCustRecLen As Integer, NumOfRates As Integer
  Dim UBRateTblRecLen As Integer, RateFile As Integer
  Dim Cnt As Integer, BillFile As Integer
  Dim CustFile As Integer, BillCnt As Integer
  Dim NumOfRevs As Integer, RCnt As Integer
  Dim zz As Integer, FRCnt As Integer
  Dim WhatService As Integer, Multi As Integer
  Dim MRCnt As Integer, WhatTbl As Integer
  Dim MeterLocNum As Integer, MCnt As Integer
  Dim TZCnt As Integer, TCnt As Integer
  
  Dim ThisMeterConsp As Double, AddRevAmt As Double
  Dim ProPct As Double, FlatAmt As Double
  Dim TaxAmt As Double, TestAmt As Double
  Dim HowMuch As Double, NonMAmt As Double
  Dim TMaxAmt As Double, ProRevAmt As Double
  Dim RevAmt As Double, FuelAddAmt As Double
  Dim Diff As Double
  
  Dim NumCustRec As Long, LCnt As Long
  Dim MeterConsp As Long, TMeterConsp As Long
  Dim MaxMeterAmt As Long, NumUser As Long
  Dim MinimumConsp As Long, MeterMulti As Long
  
  Dim YadkinFlag As Boolean, WadeFlag As Boolean
  Dim ElkFlag As Boolean, ScottFlag As Boolean
  Dim DaleFlag As Boolean, SkipInactive As Boolean
  Dim PrinceFlag As Boolean, BookFlag As Boolean
  Dim CycleFlag As Boolean, ProRateFlag As Boolean
  Dim ConwayFlag As Boolean
  
  'BlockClear
  'ShowProcessingScrn "Calculating Utility Charges."
  
  UBLog "IN: MakeBillFile."
  UBLog "MBF: Calculating charges."
  
  ReDim ProrateServ(1 To 15) As Integer
  
  ReDim UBSetUp(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetUpLen
  
  If InStr(UBSetUp(1).UTILNAME, "PRINCETON") > 0 Then
    PrinceFlag = True
  End If
  If InStr(UBSetUp(1).UTILNAME, "YADKIN") > 0 Then
    YadkinFlag = True
  End If
  If InStr(UBSetUp(1).UTILNAME, "WADE") > 0 Then                'OR INSTR(UBSetUp(1).UTILNAME, "WADE") THEN
    WadeFlag = True
  End If
  If InStr(UBSetUp(1).UTILNAME, "ELKTON") > 0 Then
    ElkFlag = True
  End If
  If InStr(UBSetUp(1).UTILNAME, "SCOTTSBURG") > 0 Then
    ScottFlag = True
  End If
  If InStr(UBSetUp(1).UTILNAME, "SUMMERDALE") > 0 Then
    DaleFlag = True
  End If
  
  If UBSetUp(1).SkipInactive = "Y" Then
    SkipInactive = True
  End If
  
  If UBSetUp(1).PreByBook = "Y" And ThisBook > 0 Then
    BookFlag = True
  ElseIf UBSetUp(1).BILLCYCL = "Y" Then
    CycleFlag = True
  End If
  
  'find the electric revenue position
  For ThisRevCnt = 1 To 15
    If InStr(UBSetUp(1).Revenues(ThisRevCnt).REVNAME, "ELECTRIC") Then
      ElecRev = ThisRevCnt
      Exit For
    End If
  Next
  
  For ThisRevCnt = 1 To 15
    If UBSetUp(1).Revenues(ThisRevCnt).ProRate = "Y" Then
      ProrateServ(ThisRevCnt) = True
    End If
  Next
  
  ReDim UBBillRec(1) As UBTransRecType
  ReDim UBCustRec(1) As NewUBCustRecType
  
  UBBillRecLen = Len(UBBillRec(1))
  UBCustRecLen = Len(UBCustRec(1))
  
  NumOfRates = GetNumRateRecs%
  
  ReDim UBRateTbls(1 To NumOfRates) As UBRateTblRecType
  UBRateTblRecLen = Len(UBRateTbls(1))
  
  RateFile = FreeFile
  Open "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
  For Cnt = 1 To NumOfRates
    Get RateFile, Cnt, UBRateTbls(Cnt)
  Next
  Close RateFile
  
  NumCustRec& = FileSize&("UBCUST.DAT") \ UBCustRecLen
  
  If Exist(UBBillsFile) Then
    Kill UBBillsFile
  End If
  
  BillFile = FreeFile
  Open UBBillsFile For Random Shared As BillFile Len = UBBillRecLen
  
  CustFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As CustFile Len = UBCustRecLen
  
  BillCnt = 0
  NumOfRevs = GetNumOfRevs%
  
  For LCnt& = 1 To NumCustRec&
    ReDim UBBillRec(1) As UBTransRecType        'clear bill rec for this customer
    Get CustFile, LCnt&, UBCustRec(1)
    'IF LCnt& = 2 THEN STOP
    If UBCustRec(1).DelFlag <> 0 Then
      UBBillRec(1).TransAmt = 0
      UBBillRec(1).ActiveFlag = False
      UBBillRec(1).CustAcctNo = LCnt&
      GoTo MSkipEm
    End If
    
    If SkipInactive And UBCustRec(1).Status = "I" Then
      UBBillRec(1).TransAmt = 0
      UBBillRec(1).ActiveFlag = False
      UBBillRec(1).CustAcctNo = LCnt&
      GoTo MSkipEm
    End If
    
    If BookFlag Then
      If Val(UBCustRec(1).Book) <> ThisBook Then
        UBBillRec(1).TransAmt = 0
        For RCnt = 1 To NumOfRevs
          UBBillRec(1).RevAmt(RCnt) = 0
          UBBillRec(1).TaxAmt(RCnt) = 0
        Next
        For zz = 1 To 7
          UBBillRec(1).CurRead(zz) = 0
          UBBillRec(1).PrevRead(zz) = 0
        Next
        UBBillRec(1).ActiveFlag = False
        GoTo MSkipEm
      End If
    End If
    
    If CycleFlag Then
      If UBCustRec(1).BILLCYCL <> ThisCycle Then
        UBBillRec(1).TransAmt = 0
        For RCnt = 1 To NumOfRevs
          UBBillRec(1).RevAmt(RCnt) = 0
          UBBillRec(1).TaxAmt(RCnt) = 0
        Next
        For zz = 1 To 7
          UBBillRec(1).CurRead(zz) = UBCustRec(1).LocMeters(zz).CurRead
          UBBillRec(1).PrevRead(zz) = UBCustRec(1).LocMeters(zz).PrevRead
          UBBillRec(1).MtrTypes(zz) = GetCustMeterType(UBCustRec(), zz)
        Next
        UBBillRec(1).ActiveFlag = False
        GoTo MSkipEm
      End If
    End If
    
    If UBCustRec(1).Status <> "A" Then
      UBBillRec(1).TransAmt = 0
      For RCnt = 1 To NumOfRevs
        UBBillRec(1).RevAmt(RCnt) = 0
        UBBillRec(1).TaxAmt(RCnt) = 0
      Next
      For zz = 1 To 7
        UBBillRec(1).CurRead(zz) = UBCustRec(1).LocMeters(zz).CurRead
        UBBillRec(1).PrevRead(zz) = UBCustRec(1).LocMeters(zz).PrevRead
        UBBillRec(1).MtrTypes(zz) = GetCustMeterType(UBCustRec(), zz)
      Next
      UBBillRec(1).ActiveFlag = False
      GoTo MSkipEm
    End If
    '052698 Added tax exempt flag to bill rec
    UBBillRec(1).TaxExempt = UBCustRec(1).TAXEXPT
    
    '111398 Prorate
    ProRateFlag = False
    ProPct# = 100
    If UBCustRec(1).ProRatePCT < 100 And UBCustRec(1).ProRatePCT > 0 Then
      UBBillRec(1).ProRatePCT = UBCustRec(1).ProRatePCT
      UBLog "MBF: Prorated Account No:" + Str$(LCnt&) + "  @" + Trim$(Str$(UBBillRec(1).ProRatePCT)) + "%"
      ProPct# = Round#(UBBillRec(1).ProRatePCT * 0.01)
      ProRateFlag = True
    Else
      UBBillRec(1).ProRatePCT = 100
    End If
    
    MeterConsp& = 0
    TMeterConsp& = 0
    
    'look at flat rates
    For FRCnt = 1 To 4
      WhatService = UBCustRec(1).FlatRates(FRCnt).REVSRC
      If WhatService > NumOfRevs Then
        If ErrorScrn(6, LCnt&) Then
          AbortFlag = True
          GoTo AbortExit
        End If
      End If
      If UBCustRec(1).FlatRates(FRCnt).FRAMT <> 0 And WhatService > 0 Then
        '11/19/96 Fixed Rev. amt. to add to current rev amt
        If UBCustRec(1).FlatRates(FRCnt).FRAMT < -1000000 Then
          If ErrorScrn(6, LCnt&) Then
            AbortFlag = True
            GoTo AbortExit
          End If
        End If
        '01-09-97 Fixed Multiplier bug in flat rates
        Multi = UBCustRec(1).FlatRates(FRCnt).NumMin
        If Multi < 1 Then Multi = 1
        FlatAmt# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * Multi)
        '111398 Prorate
        If ProRateFlag And ProrateServ(WhatService) Then
          FlatAmt# = Round#(FlatAmt# * ProPct#)
        End If
        UBBillRec(1).RevAmt(WhatService) = Round#(UBBillRec(1).RevAmt(WhatService) + FlatAmt#)
        UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + FlatAmt#)
        If UBSetUp(1).Revenues(WhatService).TAXRATE > 0 And UBCustRec(1).TAXEXPT <> "Y" Then
          TaxAmt# = Round#(UBBillRec(1).RevAmt(WhatService) * UBSetUp(1).Revenues(WhatService).TAXRATE)
          UBBillRec(1).TaxAmt(WhatService) = TaxAmt#
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + UBBillRec(1).TaxAmt(WhatService))
        End If
      End If
    Next
    'end of flat rates
    '12-6-96  Monthly Billed amounts
    For MRCnt = 1 To 2
      WhatService = UBCustRec(1).Monthly(MRCnt).RevSource
      If WhatService > NumOfRevs Or WhatService < 0 Then
        'IF ErrorScrn(7, LCnt&) THEN
        '  AbortFlag = True
        '  GOTO AbortExit
        'END IF
      End If
      
      If UBCustRec(1).Monthly(MRCnt).PayAmt > 0 And WhatService > 0 Then
        TestAmt# = Round#(UBCustRec(1).Monthly(MRCnt).TotAmtPD + UBCustRec(1).Monthly(MRCnt).PayAmt)
        If TestAmt# > UBCustRec(1).Monthly(MRCnt).AMTOWED Then
          HowMuch# = Round#(UBCustRec(1).Monthly(MRCnt).AMTOWED - UBCustRec(1).Monthly(MRCnt).TotAmtPD)
        Else
          HowMuch# = UBCustRec(1).Monthly(MRCnt).PayAmt
        End If
        UBBillRec(1).RevAmt(WhatService) = Round#(UBBillRec(1).RevAmt(WhatService) + HowMuch#)
        UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + HowMuch#)
        If UBSetUp(1).Revenues(WhatService).TAXRATE > 0 And UBCustRec(1).TAXEXPT <> "Y" Then
          TaxAmt# = Round#(HowMuch# * UBSetUp(1).Revenues(WhatService).TAXRATE)
          UBBillRec(1).TaxAmt(WhatService) = Round#(UBBillRec(1).TaxAmt(WhatService) + TaxAmt#)
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + TaxAmt#)
        End If
      End If
    Next
    
    For RCnt = 1 To NumOfRevs   'look at each rev line
      MeterConsp& = 0
      TMeterConsp& = 0
      GoSub GetWhatRateTable
      If WhatTbl Then
        If UBSetUp(1).Revenues(RCnt).UseMtr = "N" Then
          'if this is a non-metered service
          '02-05-97 added fix add to current rev amt
          If UBRateTbls(WhatTbl).MINAMT > -1000000 Then
            NonMAmt# = UBRateTbls(WhatTbl).MINAMT
            If ProRateFlag And ProrateServ(RCnt) Then
              NonMAmt# = Round#(NonMAmt# * ProPct#)
            End If
            UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + NonMAmt#)
            UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + NonMAmt#)
          Else
            RateCodeErrScrn UBRateTbls(WhatTbl).RATECODE
            AbortFlag = True
            GoTo AbortExit
          End If
          GoTo GotAmt
        End If
        'it's metered
        MeterType$ = UBCustRec(1).Serv(RCnt).RMtrType
        MeterLocNum = 0
        For MCnt = 1 To 7
          If MeterType$ = UBCustRec(1).LocMeters(MCnt).MTRType Then
            MeterLocNum = MCnt
            UBBillRec(1).CurRead(MCnt) = UBCustRec(1).LocMeters(MCnt).CurRead
            UBBillRec(1).PrevRead(MCnt) = UBCustRec(1).LocMeters(MCnt).PrevRead
            UBBillRec(1).MtrTypes(MCnt) = GetCustMeterType(UBCustRec(), MCnt)
            'Found correct meter
            '052797 Added to stop overflow error.
            If (UBCustRec(1).LocMeters(MCnt).CurRead < 0) Or (UBCustRec(1).LocMeters(MCnt).PrevRead < 0) Then
              If ErrorScrn(1, LCnt&) Then
                AbortFlag = True
                GoTo AbortExit
              End If
              MeterConsp& = 0
            Else
              MeterConsp& = UBCustRec(1).LocMeters(MCnt).CurRead - UBCustRec(1).LocMeters(MCnt).PrevRead
            End If
            If MeterConsp& < 0 Then
              MaxMeterAmt& = 10& ^ (Len(Str$(UBCustRec(1).LocMeters(MCnt).PrevRead)) - 1)
              MeterConsp& = (MaxMeterAmt& - UBCustRec(1).LocMeters(MCnt).PrevRead) + UBCustRec(1).LocMeters(MCnt).CurRead
            End If
            If UBCustRec(1).LocMeters(MCnt).MTRMulti > 0 Then
              ThisMeterConsp# = (0# + MeterConsp&) * UBCustRec(1).LocMeters(MCnt).MTRMulti
              '                  ^This forces basic to convert to a Double
              '                   before calculation, traps overflow errors
              If ThisMeterConsp# > 2147483647 Then
                '                  ^Max long integer value
                If ErrorScrn(1, LCnt&) Then
                  AbortFlag = True
                  GoTo AbortExit
                End If
              End If
              MeterConsp& = ThisMeterConsp#
            End If
            If (UBBillRec(1).MtrTypes(MCnt) = 1 Or UBBillRec(1).MtrTypes(MCnt) = 2 Or UBBillRec(1).MtrTypes(MCnt) = 3) And UBCustRec(1).LocMeters(MCnt).MTRUnit = "C" Then
              MeterConsp& = MeterConsp& * 7.481
              'convert units from cubic feet to gallons here
            End If
            TMeterConsp& = TMeterConsp& + MeterConsp&
          End If
        Next
        If MeterLocNum = 0 Then
          If ErrorScrn(4, LCnt&) Then
            AbortFlag = True
            GoTo AbortExit
          End If
        End If
        AddRevAmt# = 0
        TMaxAmt# = 0
        If UBRateTbls(WhatTbl).MaxAmt > 0 Then
          TMaxAmt# = UBRateTbls(WhatTbl).MaxAmt
        End If
        
        If UBCustRec(1).LocMeters(MeterLocNum).NumUser > 1 Then
          TMaxAmt# = Round#(UBRateTbls(WhatTbl).MaxAmt * UBCustRec(1).LocMeters(MeterLocNum).NumUser)
          'adjust min consumption for calc below
          NumUser& = UBCustRec(1).LocMeters(MeterLocNum).NumUser - 1
          AddRevAmt# = NumUser& * UBRateTbls(WhatTbl).MINAMT
          MinimumConsp& = NumUser& * UBRateTbls(WhatTbl).MINUNITS
          TMeterConsp& = TMeterConsp& - MinimumConsp&
          If (TMeterConsp& - UBRateTbls(WhatTbl).MINUNITS) <= 0 Then
            '062697 fix for min consump test to actual (NumUsers * MINUNITS)
            '071201 Added fix for prorating
            ProRevAmt# = 0
            If ProRateFlag And ProrateServ(RCnt) Then
              ProRevAmt# = Round#((AddRevAmt# + UBRateTbls(WhatTbl).MINAMT) * ProPct#)
              UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + ProRevAmt#)
              UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + ProRevAmt#)
            Else
              UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + (AddRevAmt# + UBRateTbls(WhatTbl).MINAMT))
              UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + (AddRevAmt# + UBRateTbls(WhatTbl).MINAMT))
            End If
            GoTo GotAmt
          End If
        Else
          NumUser& = 1
          If TMaxAmt# > 0 Then
            TMaxAmt# = UBRateTbls(WhatTbl).MaxAmt
          End If
          '033198 Added code to Calc correctly for Conway...
          If ConwayFlag Then
            If TMeterConsp& Mod 1000 Then
              TMeterConsp& = (Int(TMeterConsp& / 1000) + 1)
            Else
              TMeterConsp& = Int(TMeterConsp& / 1000)
            End If
          End If
          '033198 Conway *********
          '052998 Added code for calc method Princeton
          'summerdale
          If PrinceFlag Or WadeFlag Or ScottFlag Or DaleFlag Then
            If TMeterConsp& Mod 1000 Then
              TMeterConsp& = (Int(TMeterConsp& / 1000) + 1)
            Else
              TMeterConsp& = Int(TMeterConsp& / 1000)
            End If
            TMeterConsp& = TMeterConsp& * 1000
          ElseIf YadkinFlag Then
            If TMeterConsp& Mod 1000 Then
              TMeterConsp& = TMeterConsp& / 1000
              TMeterConsp& = TMeterConsp& * 1000
            End If
          End If
          'Princeton*****
          If TMeterConsp& <= UBRateTbls(WhatTbl).MINUNITS Then
            'if we bill the minium
            If UBRateTbls(WhatTbl).MINAMT > -1000000 Then
              RevAmt# = NumUser& * UBRateTbls(WhatTbl).MINAMT
              If ProRateFlag And ProrateServ(RCnt) Then
                RevAmt# = Round#(RevAmt# * ProPct#)
              End If
              UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + RevAmt#)
              UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + RevAmt#)
              GoTo GotAmt
            End If
          End If
        End If
        '01-20-97 Added Fix for minium units test for equal to also.
        '04-23-1997 'Fixed to ADD TO TOTAL
        RevAmt# = GetRevCharge#(UBRateTbls(WhatTbl), TMeterConsp&, MeterMulti&)
        RevAmt# = Round#(RevAmt# + AddRevAmt#)
        '111398 Prorate
        If ProRateFlag And ProrateServ(RCnt) Then
          RevAmt# = Round#(RevAmt# * ProPct#)
        End If
        If TMaxAmt# > 0 Then
          If RevAmt# > TMaxAmt# Then
            RevAmt# = TMaxAmt#
          End If
        End If
        UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + RevAmt#)
        If RCnt = ElecRev Then
          FuelAddAmt# = Round#(FuelAdjAmt# * TMeterConsp&)
          UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + FuelAddAmt#)
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + FuelAddAmt#)
        End If
        UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + RevAmt#)
GotAmt:
        
        If UBSetUp(1).Revenues(RCnt).TAXRATE > 0 And UBCustRec(1).TAXEXPT <> "Y" Then
          TaxAmt# = Round#(UBBillRec(1).RevAmt(RCnt) * UBSetUp(1).Revenues(RCnt).TAXRATE)
          UBBillRec(1).TaxAmt(RCnt) = Round#(UBBillRec(1).TaxAmt(RCnt) + TaxAmt#)
          UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + TaxAmt#)
        End If
      Else
        If Len(Trim$(UBCustRec(1).Serv(RCnt).RMtrType)) > 0 Then
          If ErrorScrn(3, LCnt&) Then
            AbortFlag = True
            GoTo AbortExit
          End If
        End If
      End If
    Next        'loop through all revenue sources
    If UBCustRec(1).Status = "I" And UBBillRec(1).TransAmt > 0 Then
      UBBillRec(1).TransAmt = 0
      For RCnt = 1 To NumOfRevs
        UBBillRec(1).RevAmt(RCnt) = 0
      Next
      UBBillRec(1).ActiveFlag = False
      UBBillRec(1).CustAcctNo = LCnt&
    Else
      UBBillRec(1).ActiveFlag = False
    End If
    
    'Mod for cleveland***
    If UBCustRec(1).CUSTTYPE = "NON" Then
      UBBillRec(1).CustAcctNo = LCnt&
      UBBillRec(1).NONProfit = "Y"
    End If
    '********************
    
    If UBBillRec(1).TransAmt > 0 Then
      BillCnt = BillCnt + 1
      UBBillRec(1).ActiveFlag = True
      UBBillRec(1).CustAcctNo = LCnt&
    End If
    
    '0727 Added NEW trap for a meter defined with no rate code.
    '    FOR MTstCnt = 1 TO 7
    '      IF LEN(Trim$(UBCustRec(1).LocMeters(MTstCnt).MTRType)) > 0 THEN
    '        FOR MTCnt = 1 TO 7
    '          IF UBBillRec(1).MtrTypes(MTCnt) > 0 THEN
    '            GOTO ThereOK
    '          END IF
    '          IF ErrorScrn(8, LCnt&) THEN
    '            AbortFlag = True
    '            GOTO AbortExit
    '          END IF
    '        NEXT
    '      END IF
    '    NEXT
    
    '04-07-99 Added special tax calc for elkton
    If ElkFlag Then
      If UBBillRec(1).ActiveFlag Then
        For TZCnt = 1 To 15
          If UBBillRec(1).TaxAmt(TZCnt) > 0 Then
            Ctype$ = Trim$(UBCustRec(1).CUSTTYPE)
            Select Case Ctype$
            Case "R"
              If UBBillRec(1).TaxAmt(TZCnt) > 2 Then
                Diff# = Round#(UBBillRec(1).TaxAmt(TZCnt) - 2)
                UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt - Diff#)
                UBBillRec(1).TaxAmt(TZCnt) = 2
              End If
            Case "C"
              If UBBillRec(1).TaxAmt(TZCnt) > 20 Then
                Diff# = Round#(UBBillRec(1).TaxAmt(TZCnt) - 20)
                UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt - Diff#)
                UBBillRec(1).TaxAmt(TZCnt) = 20
              End If
            Case Else
              If ErrorScrn(9, LCnt&) Then
                AbortFlag = True
                GoTo AbortExit
              End If
            End Select
          End If
        Next
      End If
    End If
    
    '11-09-00           'water
    '      SpecialAmt# = 0
    '      TWAmt# = 0
    '      TSAmt# = 0
    '      TEAmt# = 0
    '      IF UBBillRec(1).RevAmt(1) > 0 THEN
    '        IF INSTR(UBCustRec(1).Serv(1).RATECODE, "WRIN") > 0 THEN
    '          TWAmt# = Round#((UBBillRec(1).RevAmt(1) * .005))
    '        END IF
    '      END IF
    '      IF UBBillRec(1).RevAmt(2) > 0 THEN
    '        IF INSTR(UBCustRec(1).Serv(2).RATECODE, "SRIN") > 0 THEN
    '          TSAmt# = Round#((UBBillRec(1).RevAmt(2) * .005))
    '        END IF
    '      END IF
    '      IF UBBillRec(1).RevAmt(3) > 0 THEN
    '        IF INSTR(UBCustRec(1).Serv(3).RATECODE, "ERIN") > 0 THEN
    '          TEAmt# = Round#((UBBillRec(1).RevAmt(3) * .005))
    '        END IF
    '      END IF
    '      SpecialAmt# = Round#(TWAmt# + TSAmt# + TEAmt#)
    '      IF SpecialAmt# > 0 THEN
    '        UBBillRec(1).RevAmt(8) = SpecialAmt#
    '        UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + SpecialAmt#)
    '      END IF
    '    END IF
    'end of elkton section
    
MSkipEm:
    Put BillFile, LCnt&, UBBillRec(1)
    'If AskAbandonPrint% Then
    '  AbortFlag = True
    '  Exit For
    'End If
    ShowPctComp LCnt&, NumCustRec&
  Next
  
AbortExit:
  
  Close BillFile, CustFile
  
  If AbortFlag Then
    UBLog "MBF: ABORTED!"
  Else
    UBLog "MBF: Finished calculations."
  End If
  
  UBLog "OUT: MakeBillFile."
  Erase UBBillRec, UBCustRec, UBSetUp, UBRateTbls
  Exit Sub
  '*******************************
  
GetWhatRateTable:
  WhatTbl = 0
  RATECODE$ = Trim$(UBCustRec(1).Serv(RCnt).RATECODE)
  If Len(RATECODE$) Then        'if this rev has a rate code
    For TCnt = 1 To NumOfRates  'find the right one
      If RATECODE$ = Trim$(UBRateTbls(TCnt).RATECODE) Then
        WhatTbl = TCnt
        Exit For
      End If
    Next
  End If
  
  Return
  
End Sub
  
'Sub PostBillTrans()
'
'  UBLog "IN: Bill Posting."
'
'  If Not Exist(UBBillsFile) Then
'    UBLog "ERROR: UBBILLS.DAT Calculation file NOT FOUND!"
'    CursorOff
'    BlockClear
'    DisplayUBScrn "NON2POST"
'    WaitForAction
'    GoTo ExitBillPost:
'  End If
'
'  If Not Exist("UBBILLS.PRN") Then
'    UBLog "ERROR: UBBILLS.PRN Print File NOT FOUND!"
'    CursorOff
'    BlockClear
'    DisplayUBScrn "NOTPRNTD"
'    WaitForAction
'    GoTo ExitBillPost:
'  End If
'
'  LibName$ = "UB"
'  ScrnName$ = "PSTBILLS"
'
'  ReDim UBSetUpRec(1) As UBSetupRecType
'  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen      'load setup file
'
'  TownName$ = UBSetUpRec(1).UTILNAME
'
'  'Section to check for customer modifications
'  'Town of Lilesville Special Discount Situation
'
'  If InStr(TownName$, "INDIAN TRAIL") Then
'    IndianFlag = True
'  End If
'
'  If InStr(TownName$, "SEDGEFIELD") Then
'    SedgeFlag = True
'  End If
'
'  '--define the multi-choice fields
'  '--Initialize the form name array
'  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
'
'  '--define Quick Screen form editing arrays
'  ReDim frm(1) As FormInfo
'  ReDim Form$(NumFlds, 2)
'  ReDim Fld(NumFlds) As FieldInfo
'
'  '--for each screen, get first and last fields
'  StartEl = 0
'
'  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
'  Action = 1
'  frm(1).StayOnField = True
'
'  '--Set screen number to one and display screen
'
'  BlockClear
'
'  DisplayUBScrn ScrnName$
'
'  If FileSize&("UBSNDEM.DAT") > 0 Then
'    For Cnt = 1 To 3
'      QPSound 1750, 2
'      QPSound 1650, 2
'    Next
'  End If
'
'  Do
'
'    EditForm Form$(), Fld(), frm(1), Cnf, Action
'
'    '--Check for Key presses
'    Select Case frm(1).KeyCode
'    Case F10Key
'      OKFlag = True
'    Case EscKey
'      ExitFlag = True
'    End Select
'
'  Loop Until OKFlag Or ExitFlag
'
'  If ExitFlag Then
'    BlockClear
'    CursorOff
'    DisplayUBScrn "POSTCAN"
'    WaitForAction
'    UBLog "ABORTED:  Bill Posting"
'    GoTo ExitBillPost
'  End If
'
'  UBLog "START: Posting Transactions."
'
'  ReDim UBBillRec(1) As UBTransRecType
'  ReDim UBCustRec(1) As NewUBCustRecType
'
'  UBBillRecLen = Len(UBBillRec(1))
'  UBCustRecLen = Len(UBCustRec(1))
'
'  UBCust = FreeFile
'  Open "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
'  UBBill = FreeFile
'  Open UBBillsFile For Random Shared As UBBill Len = UBBillRecLen
'  UBTran = FreeFile
'  Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBBillRecLen
'
'  NumOfTranRecs& = LOF(UBTran) \ UBBillRecLen
'  NumOfBillRecs = LOF(UBBill) \ UBBillRecLen
'
'  ShowProcessingScrn "Posting Billing Transactions"
'  For BillCnt = 1 To NumOfBillRecs
'    Get UBBill, BillCnt, UBBillRec(1)
'    If (UBBillRec(1).ActiveFlag And UBBillRec(1).TransAmt > 0) Or (UBBillRec(1).NONProfit = "Y") Then
'      PostedCnt& = PostedCnt& + 1
'      NumOfTranRecs& = NumOfTranRecs& + 1       'point to next trans to write
'      Get UBCust, BillCnt, UBCustRec(1)
'      EstFlag$ = Trim$(UBCustRec(1).EstFlag)
'      For MRCnt = 1 To 2
'        WhatService = UBCustRec(1).Monthly(MRCnt).RevSource
'        If UBCustRec(1).Monthly(MRCnt).PayAmt > 0 And WhatService > 0 Then
'          TestAmt# = Round#(UBCustRec(1).Monthly(MRCnt).TotAmtPD + UBCustRec(1).Monthly(MRCnt).PayAmt)
'          If TestAmt# > UBCustRec(1).Monthly(MRCnt).AMTOWED Then
'            HowMuch# = Round#(UBCustRec(1).Monthly(MRCnt).AMTOWED - UBCustRec(1).Monthly(MRCnt).TotAmtPD)
'          Else
'            HowMuch# = UBCustRec(1).Monthly(MRCnt).PayAmt
'          End If
'          UBCustRec(1).Monthly(MRCnt).TotAmtPD = Round#(UBCustRec(1).Monthly(MRCnt).TotAmtPD + HowMuch#)
'        End If
'      Next
'      '062597 added removal of nonrecurring flat rates
'      FRFlag = False
'      For FRCnt = 1 To 4        'Remove non-recurring flat rates
'        If UBCustRec(1).FlatRates(FRCnt).FRFREQ = "N" Then
'          UBCustRec(1).FlatRates(FRCnt).FRDESC = ""
'          UBCustRec(1).FlatRates(FRCnt).FRAMT = 0
'          UBCustRec(1).FlatRates(FRCnt).FRFREQ = ""
'          UBCustRec(1).FlatRates(FRCnt).REVSRC = 0
'          UBCustRec(1).FlatRates(FRCnt).NumMin = 0
'          FRFlag = True
'        End If
'      Next
'      If FRFlag Then
'        UBLog "BILL POST: Removed Flat Rate. Acct:" + Str$(BillCnt)
'      End If
'      '111698 Prorate
'      If UBBillRec(1).ProRatePCT < 100 Then
'        UBLog "BILL POST: Reset Prorate Acct:" + Str$(BillCnt) + " PCT:" + Str$(UBBillRec(1).ProRatePCT)
'      End If
'      UBCustRec(1).ProRatePCT = 100
'      '*************
'      UBCustRec(1).PrevBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
'      UBCustRec(1).CurrBalance = UBBillRec(1).TransAmt
'      UBBillRec(1).RunBalance = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
'      For RevCnt = 1 To MaxRevsCnt
'        UBCustRec(1).CurrRevAmts(RevCnt) = Round#(UBCustRec(1).CurrRevAmts(RevCnt) + UBBillRec(1).RevAmt(RevCnt) + UBBillRec(1).TaxAmt(RevCnt))
'      Next
'      UBBillRec(1).TransType = TranUtilityBill  'set transaction to Type 1
'      UBBillRec(1).TransDesc = "Utility Billing"
'      UBBillRec(1).TransDate = UBBillRec(1).BillDate
'
'      For MtrCnt = 1 To 7
'        CubMtr = False
'        If UBCustRec(1).LocMeters(MtrCnt).CurRead >= 0 Then
'          If Len(EstFlag$) > 0 Then
'            UBBillRec(1).EstRead(MtrCnt) = "Y"
'          End If
'          If UBCustRec(1).LocMeters(MtrCnt).MTRUnit = "C" Then
'            CubMtr = True
'          End If
'          ReadAmt& = UBBillRec(1).CurRead(MtrCnt) - UBBillRec(1).PrevRead(MtrCnt)
'          If ReadAmt& < 0 Then  'Meter rolled over or, been misread
'            MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(MtrCnt))) - 1)
'            ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MtrCnt)) + UBBillRec(1).CurRead(MtrCnt)
'          End If
'          If CubMtr Then
'            ReadAmt& = ReadAmt& * 7.481
'          End If
'          If ReadAmt& < 1 Then
'            ReadAmt& = 1
'          End If
'          If UBCustRec(1).LocMeters(MtrCnt).AvgUse < 1 Then
'            UBCustRec(1).LocMeters(MtrCnt).AvgUse = 1
'          End If
'          If UBCustRec(1).LocMeters(MtrCnt).UseCnt < 1 Then
'            UBCustRec(1).LocMeters(MtrCnt).UseCnt = 1
'          End If
'          TUse# = ReadAmt& + (UBCustRec(1).LocMeters(MtrCnt).AvgUse * UBCustRec(1).LocMeters(MtrCnt).UseCnt)
'          UBCustRec(1).LocMeters(MtrCnt).UseCnt = UBCustRec(1).LocMeters(MtrCnt).UseCnt + 1
'          UBCustRec(1).LocMeters(MtrCnt).AvgUse = TUse# / UBCustRec(1).LocMeters(MtrCnt).UseCnt
'          UBCustRec(1).LocMeters(MtrCnt).ReadFlag = ""
'          If SedgeFlag Then
'            UBCustRec(1).LocMeters(MtrCnt).CurRead = 0
'            UBCustRec(1).LocMeters(MtrCnt).PrevRead = 0
'            UBCustRec(1).LocMeters(MtrCnt).AvgUse = 0
'          End If
'        End If
'      Next
'      PrevLastTrans& = UBCustRec(1).LastTrans
'      UBBillRec(1).PrevTrans = PrevLastTrans&
'      UBCustRec(1).LastTrans = NumOfTranRecs&
'      If IndianFlag Then
'        UBCustRec(1).USERCODE1 = ""
'      End If
'      'DALE
'      Put UBCust, BillCnt, UBCustRec(1)
'      Put UBTran, NumOfTranRecs&, UBBillRec(1)
'      '**************
'    End If
'    ShowPctComp BillCnt, NumOfBillRecs
'  Next
'  Close
'  UBLog "  DONE: Posting Transactions."
'  UBLog "POSTED:" + Str$(PostedCnt&) + " New BILL Transactions."
'  'DALE
'  KillFile UBBillsFile
'  KillFile "UBBILLS.PRN"
'  '**************
'  UBLog "KILLED: UBBILLS.DAT & UBBILLS.PRN"
'
'  ShowProcessingScrn "Activating Pending Accounts."
'
'  UBLog "ACTIVATING ACCOUNTS:"
'
'  UBCust = FreeFile
'  Open "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
'  NumOfCust& = LOF(UBCust) / UBCustRecLen
'  For Cnt = 1 To NumOfCust&
'    Get UBCust, Cnt, UBCustRec(1)
'    If UBCustRec(1).Status = "P" Then
'      UBCustRec(1).Status = "A"
'      UBLog "ACTIVATED: " + Str$(Cnt) + "  " + UBCustRec(1).CUSTNAME
'      Activated = Activated + 1
'      Put UBCust, Cnt, UBCustRec(1)
'    End If
'    ShowPctComp Cnt, CInt(NumOfCust&)
'  Next
'  Close
'  UBLog "     DONE: Activating Accounts."
'  UBLog "ACTIVATED:" + Str$(Activated) + " Pending Accounts."
'  BlockClear
'  DisplayUBScrn "UPDATEOK"
'  WaitForAction
'
'ExitBillPost:
'  UBLog "OUT: Bill Posting." + CRLF$
'End Sub
  
Public Sub PreBillReport()
  'Acct$ = Space$(7)
  Dim Pro As String, CurRead As String, PreRead As String
  Dim Ok As Integer, Cnt As Integer, UBSetUpLen As Integer
  Dim AbortFlag As Boolean, MowFlag As Boolean
  Dim TennFlag As Boolean, TempRev As String
  Dim Temp2 As String, TownName As String
  Dim NumOfRevs As Integer, NumOfRates  As Integer
  Dim UBRateTblRecLen As Integer, RateFile As Integer
  Dim DoFuelAdjFlag As Boolean, SkipInactive As Boolean
  Dim SkipSeparator As Boolean, BookFlag As Boolean
  Dim ThisCycle As Integer, ThisBook As Integer
  Dim CycleFlag As Boolean, UsingAcct As Boolean
  Dim OKFlag As Boolean, DoneOne As Boolean
  Dim SeqFlag As String, IndexName As String
  Dim Choice As Integer, ExitFlag As Integer, MaxLines As Integer
  Dim FuelAdjAmt As Double, FlatAmt As Double
  Dim IdxTypeText As String, TheDate As String, FF As String
  Dim UBCustRecLen As Integer, UBBillRecLen As Integer
  Dim TBooks As Integer, NumOfRecs As Integer
  Dim UBBill As Integer, UBCust As Integer, UBRpt As Integer
  Dim ThisCustRec As Long, LineCnt As Integer
  Dim BillTo As String, BadBookFlag As Boolean
  Dim WhatBook As Integer, WhatRate As Integer
  Dim FRCnt As Integer, WhatService As Integer
  Dim Multi As Integer, TRevCnt As Integer
  Dim TRateCnt As Integer, PrintedRevAmt As Boolean
  Dim MINAMT As Long, MCCnt As Integer
  Dim CubMtr As Boolean, LocMeterType As String
  Dim MeterMulti As Long, MeterNum As String
  Dim Consump As Long, ThisMeterUseCnt As Integer
  Dim ReadAmt As Long, MaxMeterAmt As Long, AvgUse As Long
  Dim HiConsump As Long, LowConsump As Long
  Dim TTRevCnt  As Integer, ConsumpFlag As Boolean
  Dim CurReadAmt As Long, PreReadAmt As Long
  Dim ConsumpAmt As Long, Bills2Print As Integer
  Dim NONRate As Integer, NONRateCnt As Integer
  Dim CTaxAmt As Double, AcctBalance As Double
  Dim TXCnt As Integer, HasAPumpCode As Integer
  Dim TAcctBalance As Double, TBookGTot As Double
  Dim WhatPump As Integer, MPCnt As Integer
  Dim PumpMtrOK As Boolean ', CubMtr As Boolean
  Dim Book As String, RptText As String
  Dim RaCnt As Integer, ZCnt As Integer, RCnt As Integer
  Dim TotalFlatAmt As Double, TotalRevAmt As Double
  Dim TotalTaxAmt As Double, TestTot As Double
  Dim TBookAmt As Double, TBTaxAmt As Double
  Dim TPumps As Integer, TBCnt As Integer
  Dim TMMConsump As Double
  Dim CustPump As String, ThisPump As String
  Dim PageNo As Integer
  Dim CRLF As String
  
  CRLF$ = Chr$(13) + Chr$(10)
  Pro$ = Space$(4)
  CurRead$ = Space$(9)
  PreRead$ = Space$(9)
  
  UBLog "IN: Prebilling Report"
  
  If Exist("UBBILLS.DAT") And Exist("UBBILLS.PRN") Then
    UBLog "ERROR: UNPOSTED BILLING DETECTED!"
    UBLog "ASKING USER WANT TO CONTINUE?"
    Ok = PreBillYouSure%
    If Not Ok Then
      UBLog "USER ABORTED PREBILLING."
      AbortFlag = True
      GoTo ExitPreReport
    Else
      UBLog "USER WANTS TO CONTINUE!"
      Kill "UBBILLS.PRN"
    End If
  End If
  
  Temp2$ = Space$(12)
  NumOfRevs = MaxRevsCnt        'assume max munber of revenue sources
  
  NumOfRates = GetNumRateRecs%
  ReDim UBRateTbls(1 To NumOfRates) As UBRateTblRecType
  UBRateTblRecLen = Len(UBRateTbls(1))
  
  ReDim RateConsump(1 To NumOfRates) As Double
  
  RateFile = FreeFile
  Open "UBRATE.DAT" For Random Shared As RateFile Len = UBRateTblRecLen
  For Cnt = 1 To NumOfRates
    Get RateFile, Cnt, UBRateTbls(Cnt)
  Next
  Close
  
  'SortT UBRateTbls(1), NumOfRates, 0, UBRateTblRecLen, 0, 4
  
  ReDim ProrateServ(1 To 15) As Integer
  
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  TownName$ = UBSetUpRec(1).UTILNAME
  If InStr(TownName$, "MOWAS") > 0 Then
    MowFlag = True
  End If
  
  If UBSetUpRec(1).DEFSTATE = "TN" Then
    TennFlag = True
  End If
  
  ReDim RevDesc(1 To MaxRevsCnt) As String * 12
  For Cnt = 1 To MaxRevsCnt     'find last active revenue
    TempRev$ = Trim$(UBSetUpRec(1).Revenues(Cnt).REVNAME)
    If Len(TempRev$) = 0 Then
      NumOfRevs = Cnt - 1       'set actual number of revenues
      Exit For
    Else        'build revenue description lines
      LSet RevDesc(Cnt) = UCase$(TempRev$)
      If InStr(RevDesc(Cnt), "ELECTRIC") Then
        DoFuelAdjFlag = True
      End If
    End If
  Next
  
  '111398 Prorate
  For Cnt = 1 To MaxRevsCnt
    If UBSetUpRec(1).Revenues(Cnt).ProRate = "Y" Then
      ProrateServ(Cnt) = True
    End If
  Next
  
  If UBSetUpRec(1).SkipInactive = "Y" Then
    SkipInactive = True
  End If
  
  If UBSetUpRec(1).SkipSeparator = "Y" Then
    SkipSeparator = True
  End If
  

  If UBSetUpRec(1).PreByBook = "Y" Then
'Need to make a function to get the Book if they prebill
'by books
    'ThisBook = GetBillBook%
    
    If ThisBook = -1 Then
      BookFlag = False
    ElseIf ThisBook <= 0 Then
      GoTo ExitPreReport
    Else
      BookFlag = True
    End If
  ElseIf UBSetUpRec(1).BILLCYCL = "Y" Then

'need to make a function to get cycle
    'ThisCycle = GetBillCycle%
    If ThisCycle <= 0 Then
      GoTo ExitPreReport
    Else
      CycleFlag = True
    End If
  End If
  
  If UBSetUpRec(1).UseSeq = "Y" Then
    SeqFlag$ = "Y"
  End If
  
Restart:
  
  GetPreBillOrder Choice, ExitFlag, SeqFlag$
  
  If ExitFlag Then GoTo ExitPreReport
  
  If DoFuelAdjFlag Then
'need to make a function to get the Fuel adjustment amount.
    'FuelAdjAmt# = GetAdjFactor#
    UBLog "Fuel adjustment factor:" + Str$(FuelAdjAmt#)
  Else
    FuelAdjAmt# = 0
  End If
  
  If FuelAdjAmt# = -10000 Then GoTo Restart
  
  Select Case Choice
  Case 0
    ExitFlag = True
  Case 1        'Name
    IndexName$ = NameIndexFile
    OKFlag = True
  Case 2        'Acct
    IndexName$ = ""
    UsingAcct = True
    OKFlag = True
  Case 3        'Location
    IndexName$ = BookIndexFile
    OKFlag = True
  Case 4        'Postal Route
    IdxTypeText$ = "Postal Route"
    MakePostalIndex IdxTypeText$
    IndexName$ = TempIndexName
    OKFlag = True
  Case 5        'ZipCode
    IdxTypeText$ = "Zip-Code"
    If MowFlag Then
      MakeMowZipCodeIndex IdxTypeText$
    Else
      MakeZipCodeIndex IdxTypeText$
    End If
    IndexName$ = TempIndexName
    OKFlag = True
  Case 6        'Sequence number
    IdxTypeText$ = "Sequence Number"
    MakeSequenceIndex IdxTypeText$
    IndexName$ = TempIndexName
    OKFlag = True
  End Select
  
  MakeBillFile AbortFlag, FuelAdjAmt#, ThisCycle, ThisBook
  
  If AbortFlag Then GoTo ExitPreReport
  
  MaxLines = 53
  'format
  ReDim Fmt$(0 To 7)
  Fmt$(0) = String$(80, "-")
  Fmt$(1) = "#########.##"
  Fmt$(2) = "#########"
  Fmt$(3) = "######.##"
  Fmt$(4) = "###########"
'  Fmt$(5) = "$$,#########.##"
'  Fmt$(6) = "$$#######,.##"
  Fmt$(5) = "$############.##"
  Fmt$(6) = "$##########.##"
  Fmt$(7) = "  #####  "
  
  TheDate$ = "Date: " + Date$
  
  ReDim UBCustRec(1 To 2) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  
  ReDim UBBillRec(1) As UBTransRecType
  UBBillRecLen = Len(UBBillRec(1))
  
  ReDim FlatTotals(1 To NumOfRevs) As Double
  '021998 added flat revenue totals
  ReDim RevTotals(1 To NumOfRevs) As Double     'Revenue total amts
  '052097 added tax by revenue totals
  ReDim TaxTotals(1 To NumOfRevs) As Double     'Tax total amts
  ReDim ConsumpTot(1 To NumOfRevs, 1 To 2) As Double            'Consumption total amts
  ReDim RateConsump(1 To NumOfRates) As Double
  '012698 Added bill count by rate code
  ReDim RateCount(1 To NumOfRates) As Long
  ReDim RateTotals(1 To NumOfRates) As Double   'Rates total amts
  '052097 added tax by rate code totals
  ReDim RTaxTot(1 To NumOfRates) As Double      'Rates Tax total amts
  '052097 added tax by book totals to type def
  ReDim BookConsump(0 To 1) As BookConsumpType  'Consumption by book
  ReDim PumpConsump(0 To 1) As PumpConsumpType  'Consumption by pump code
  ReDim TaxExmp(0 To NumOfRevs) As Double
  
  TBooks = 0
  
  If UsingAcct Then
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  Else          'load the index
    UBLog "Loading index file: " + IndexName$
    NumOfRecs = FileSize(IndexName$) \ 4
    ReDim IndexArray(1 To NumOfRecs) As UBCustIndexRecType
    'FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
  End If
  
  UBBill = FreeFile
  Open UBBillsFile For Random Shared As UBBill Len = UBBillRecLen
  UBCust = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  UBRpt = FreeFile
  Open "UBPREBIL.RPT" For Output As UBRpt
  
'  BlockClear
'  ShowProcessingScrn "Processing Pre-Billing Report"
  UBLog "Writing prebilling report to disk."
  
  GoSub PrintPreHeader
  
  For Cnt = 1 To NumOfRecs
    If UsingAcct Then
      ThisCustRec& = Cnt
    Else
      ThisCustRec& = IndexArray(Cnt).RecNum
    End If
    
    Get UBCust, ThisCustRec&, UBCustRec(1)
    
    If UBCustRec(1).DelFlag Then
      GoTo SkipEM
    End If
    
    If SkipInactive And UBCustRec(1).Status <> "A" Then
      GoTo SkipEM
    ElseIf UBCustRec(1).Status = "F" Then       'skip over final's
      GoTo SkipEM
    ElseIf UBCustRec(1).Status = "B" Then       'skip over B-Status
      GoTo SkipEM
    End If
    
    If BookFlag Then
      If Val(UBCustRec(1).Book) <> ThisBook Then
        GoTo SkipEM
      End If
    End If
    
    If CycleFlag Then
      If UBCustRec(1).BILLCYCL <> ThisCycle Then
        GoTo SkipEM
      End If
    End If
    
    Get UBBill, ThisCustRec&, UBBillRec(1)
    
    If LineCnt > MaxLines Then
      Print #UBRpt, FF$
      GoSub PrintPreHeader
    End If
    
    If UBBillRec(1).ActiveFlag <> 0 Then
      If UBCustRec(1).BillTo = "O" Then
        BillTo$ = " O"
      Else
        BillTo$ = " C"
      End If
      GoSub GetWhatBook
      If BadBookFlag Then
        If ErrorScrn(2, ThisCustRec&) Then
          AbortFlag = True
          Exit For
        End If
      End If
      BookConsump(WhatBook).CustCnt = BookConsump(WhatBook).CustCnt + 1
      Print #UBRpt, UBCustRec(1).Status; Using(Fmt(7), ThisCustRec&);
      Print #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "  "; Left$(UBCustRec(1).CUSTNAME, 25); " "; Left$(UBCustRec(1).SERVADDR, 22); " ";
      RSet Pro$ = Format(UBBillRec(1).ProRatePCT, "###")
      
      Print #UBRpt, Pro$; "%";
      Print #UBRpt, BillTo$
      LineCnt = LineCnt + 1
      For FRCnt = 1 To 4
        WhatService = UBCustRec(1).FlatRates(FRCnt).REVSRC
        If UBCustRec(1).FlatRates(FRCnt).FRAMT <> 0 And WhatService > 0 Then
          Multi = UBCustRec(1).FlatRates(FRCnt).NumMin
          If Multi < 1 Then Multi = 1
          FlatAmt# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * Multi)
          '021998 Added flat rate summaries
          FlatTotals(WhatService) = Round#(FlatTotals(WhatService) + FlatAmt#)
        End If
      Next
      '102798 Added to skip accts that don't have a book/seq no. "J.R."
    ElseIf Len(Trim$(UBCustRec(1).Book)) = 0 And Len(Trim$(UBCustRec(1).SEQNUMB)) = 0 Then
      GoTo SkipEM
    End If
    WhatRate = 0
    DoneOne = False
    For TRevCnt = 1 To NumOfRevs
      WhatRate = 0
      If UBBillRec(1).RevAmt(TRevCnt) <> 0 Then
        DoneOne = False
        Print #UBRpt, RevDesc(TRevCnt);
        '102198 Moved out of meter loop, Stoped multi meter tax report bug
        If UBBillRec(1).TaxAmt(TRevCnt) > 0 Then
          TaxTotals(TRevCnt) = Round#(TaxTotals(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
        End If
        For TRateCnt = 1 To NumOfRates
          If UBRateTbls(TRateCnt).RATECODE = UBCustRec(1).Serv(TRevCnt).RATECODE Then
            MINAMT& = UBRateTbls(TRateCnt).MINUNITS
            WhatRate = TRateCnt
            '102198 Moved from meter loop, Stops multi meter tax report bug
            RTaxTot(WhatRate) = Round#(RTaxTot(WhatRate) + UBBillRec(1).TaxAmt(TRevCnt))
            Exit For
          End If
        Next
        If UBSetUpRec(1).Revenues(TRevCnt).UseMtr = "Y" Then
          RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
          BookConsump(WhatBook).RevAmt(TRevCnt) = Round#(BookConsump(WhatBook).RevAmt(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
          BookConsump(WhatBook).TaxAmt(TRevCnt) = Round#(BookConsump(WhatBook).TaxAmt(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
          '02-20-97 Add revenue totals by rate code
          If WhatRate > 0 Then
            RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
          End If
          PrintedRevAmt = False
          For MCCnt = 1 To 7
            CubMtr = False
            LocMeterType$ = Trim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
            MeterMulti& = UBCustRec(1).LocMeters(MCCnt).MTRMulti
            '063098 Added adjustment for cubic meters in consumption totals
            If UBCustRec(1).LocMeters(MCCnt).MTRUnit = "C" Then
              CubMtr = True
            End If
            If MeterMulti& <= 0 Then MeterMulti& = 1
            If (Len(LocMeterType$) > 0) And (LocMeterType$ = UBCustRec(1).Serv(TRevCnt).RMtrType) Then
              DoneOne = True
              MeterNum$ = Trim$(UBCustRec(1).Serv(TRevCnt).RATECODE)
              'use the Meternum$ to hold the rate code temporarily
              If Len(MeterNum$) > 0 Then
                If UBCustRec(1).LocMeters(MCCnt).NumUser > 1 Then
                  MeterNum$ = MeterNum$ + "*" + Trim$(Str$(UBCustRec(1).LocMeters(MCCnt).NumUser))
                End If
                RSet Temp2$ = MeterNum$
              End If
              ReadAmt& = UBBillRec(1).CurRead(MCCnt) - UBBillRec(1).PrevRead(MCCnt)
              If ReadAmt& < 0 Then              'Meter rolled over or, been misread
                MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(MCCnt))) - 1)
                ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MCCnt)) + UBBillRec(1).CurRead(MCCnt)
              End If
              If CubMtr Then
                ReadAmt& = ReadAmt& * 7.481
              End If
              RateConsump(WhatRate) = RateConsump(WhatRate) + (ReadAmt& * MeterMulti&)
              RateCount(WhatRate) = RateCount(WhatRate) + 1
              BookConsump(WhatBook).Consump(TRevCnt) = BookConsump(WhatBook).Consump(TRevCnt) + (ReadAmt& * MeterMulti&)
              ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + (ReadAmt& * MeterMulti&)
              Consump& = ReadAmt& * MeterMulti&
              ThisMeterUseCnt = UBCustRec(1).LocMeters(MCCnt).UseCnt
              If ThisMeterUseCnt <= 0 Then ThisMeterUseCnt = 1
              AvgUse& = UBCustRec(1).LocMeters(MCCnt).AvgUse
              If AvgUse& > 0 Then
                HiConsump& = Round#(AvgUse& * (UBSetUpRec(1).HighRead * 0.01))
                LowConsump& = Round#(AvgUse& * (UBSetUpRec(1).LowRead * 0.01))
              End If
'stop here
              Print #UBRpt, Tab(14); Temp2$; Tab(30); Using(Fmt$(2), UBBillRec(1).CurRead(MCCnt)); Tab(42); Using(Fmt$(2), UBBillRec(1).PrevRead(MCCnt)); Tab(54); Using(Fmt$(2), ReadAmt& * MeterMulti&);
              If UBCustRec(1).EstFlag = "E" Then
                Print #UBRpt, " E";             'Est. Reading
              ElseIf Consump& < LowConsump& Then
                Print #UBRpt, " L";             'Low reading
              ElseIf Consump& > HiConsump& Then
                Print #UBRpt, " H";             'High Reading
              End If
              If Consump& < MINAMT& Then
                Print #UBRpt, " M";             'Minium Usage
              End If
              If UBBillRec(1).RevAmt(TRevCnt) > 0 And PrintedRevAmt = False Then
                PrintedRevAmt = True
                Print #UBRpt, Tab(69); Using(Fmt$(3), UBBillRec(1).RevAmt(TRevCnt));
                If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
                  Print #UBRpt, "*";
                End If
              End If
              Print #UBRpt,
              LineCnt = LineCnt + 1
            End If
          Next
          '071197 Added this for mccormick. Has a sewer flat rate, Sewer is set up as
          '      a metered service but no meter on a flat rate charge. Charge was added
          '      to total, but didn't show on prebilling report.
          If Not DoneOne Then
            DoneOne = True
            Print #UBRpt, Tab(69); Using(Fmt$(3), UBBillRec(1).RevAmt(TRevCnt));
            If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
              Print #UBRpt, "*";
            End If
            'THIS WAS REMARKED OUT, I DON'T KNOW WHY?
            Print #UBRpt,
            ''''''''''''''''''''''''''''''''''''''
            LineCnt = LineCnt + 1
          End If
        Else    'it's a nonmetered service
          ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + 1
          If WhatRate > 0 Then
            RateConsump(WhatRate) = RateConsump(WhatRate) + 1
            RateCount(WhatRate) = RateCount(WhatRate) + 1
            RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
          End If
          BookConsump(WhatBook).Consump(TRevCnt) = BookConsump(WhatBook).Consump(TRevCnt) + 1
          BookConsump(WhatBook).RevAmt(TRevCnt) = Round#(BookConsump(WhatBook).RevAmt(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
          BookConsump(WhatBook).TaxAmt(TRevCnt) = Round#(BookConsump(WhatBook).TaxAmt(TRevCnt) + UBBillRec(1).TaxAmt(TRevCnt))
          RevTotals(TRevCnt) = Round#(RevTotals(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
          Print #UBRpt, Tab(69); Using(Fmt$(3), UBBillRec(1).RevAmt(TRevCnt));
          If UBBillRec(1).ProRatePCT < 100 And ProrateServ(TRevCnt) Then
            Print #UBRpt, "*";
          End If
        End If
        If Not DoneOne Then
          Print #UBRpt,
          LineCnt = LineCnt + 1
        End If
      End If
      If (TRevCnt = NumOfRevs) And UBBillRec(1).TransAmt = 0 Then
        If UBBillRec(1).TransAmt = 0 Then       'CONSUMPTION inactive account
          For TTRevCnt = 1 To NumOfRevs
            For MCCnt = 1 To 7
              LocMeterType$ = Trim$(UBCustRec(1).LocMeters(MCCnt).MTRType)
              If (Len(LocMeterType$) > 0) And (LocMeterType$ = UBCustRec(1).Serv(TTRevCnt).RMtrType) Then
                If UBBillRec(1).CurRead(MCCnt) < 0 Then
                  UBBillRec(1).CurRead(MCCnt) = 0
                End If
                If UBBillRec(1).PrevRead(MCCnt) < 0 Then
                  UBBillRec(1).PrevRead(MCCnt) = 0
                End If
                CurReadAmt& = UBBillRec(1).CurRead(MCCnt)
                PreReadAmt& = UBBillRec(1).PrevRead(MCCnt)
                If CurReadAmt& <> PreReadAmt& Then
                  If Not ConsumpFlag Then
                    Print #UBRpt, UBCustRec(1).Status; Using(Fmt(7), ThisCustRec&);
                    Print #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; "   "; Left$(UBCustRec(1).CUSTNAME, 25); "  "; Left$(UBCustRec(1).SERVADDR, 25)
                    LineCnt = LineCnt + 1
                  End If
                  ConsumpFlag = True
                  MeterNum$ = Trim$(UBCustRec(1).Serv(TTRevCnt).RATECODE)
                  If Len(MeterNum$) > 0 Then
                    If UBCustRec(1).LocMeters(MCCnt).NumUser > 1 Then
                      MeterNum$ = MeterNum$ + "*" + Trim$(Str$(UBCustRec(1).LocMeters(MCCnt).NumUser))
                    End If
                    RSet Temp2$ = MeterNum$
                  End If
                  ConsumpAmt& = CurReadAmt& - PreReadAmt&
                  '103098 Added meter roll over check to inactive consumption
                  If ConsumpAmt& < 0 Then       'Meter rolled over or, been misread
                    MaxMeterAmt& = 10& ^ (Len(Str$(PreReadAmt&)) - 1)
                    ConsumpAmt& = (MaxMeterAmt& - PreReadAmt&) + CurReadAmt&
                  End If
                  If InStr(UBCustRec(1).CUSTTYPE, "NON") Then
                    'For Nonprofits include consumption as normal   'cleveland
                    '040998 Made changes here
                    For NONRateCnt = 1 To NumOfRates
                      If UBRateTbls(NONRateCnt).RATECODE = UBCustRec(1).Serv(TTRevCnt).RATECODE Then
                        NONRate = NONRateCnt
                        Exit For
                      End If
                    Next
                    If NONRate > 0 Then
                      RateConsump(NONRate) = RateConsump(NONRate) + ConsumpAmt&
                    End If
                    ConsumpTot(TTRevCnt, 1) = ConsumpTot(TTRevCnt, 1) + ConsumpAmt&
                    BookConsump(WhatBook).Consump(TTRevCnt) = BookConsump(WhatBook).Consump(TTRevCnt) + ConsumpAmt&
                    '040998 Made changes here 'cleveland
                  Else          'add consumption to inactives
                    ConsumpTot(TTRevCnt, 2) = ConsumpTot(TTRevCnt, 2) + ConsumpAmt&
                  End If
                  Print #UBRpt, RevDesc(TTRevCnt); Tab(14); Temp2$; Tab(30); Using(Fmt$(2), CurReadAmt&); Tab(42); Using(Fmt$(2), PreReadAmt&); Tab(54); Using(Fmt$(2), ConsumpAmt&)
                  LineCnt = LineCnt + 1
                End If
              End If
            Next
          Next
        End If
        If ConsumpFlag And UBCustRec(1).Status <> "A" Then
          ConsumpFlag = False
          Print #UBRpt, "**** Consumption Noted on an Inactive Account. ****"
          LineCnt = LineCnt + 1
          If Not SkipSeparator Then
            Print #UBRpt, Fmt$(0)
            LineCnt = LineCnt + 1
          End If
        ElseIf ConsumpFlag Then
          'Customer Status is "A"
          'This happens when a cust has consumption and there rate code
          'has a zero calc amount. "i.e. a Church or other nonprofit"
          If InStr(UBCustRec(1).CUSTTYPE, "NON") Then
            Print #UBRpt, "*** NON-PROFIT ***"
            LineCnt = LineCnt + 1
          End If
          ConsumpFlag = False
          If Not SkipSeparator Then
            Print #UBRpt, Fmt$(0)
            LineCnt = LineCnt + 1
          End If
        End If
      ElseIf (TRevCnt = NumOfRevs) And UBBillRec(1).TransAmt > 0 Then
        '102998  Moved tax printing to here "now prints one tax line per customer"
        CTaxAmt# = 0
        For TXCnt = 1 To 15
          If UBBillRec(1).TaxAmt(TXCnt) > 0 Then
            CTaxAmt# = Round#(CTaxAmt# + UBBillRec(1).TaxAmt(TXCnt))
          End If
        Next
        If CTaxAmt# > 0 Then
          Print #UBRpt, " Tax"; Tab(69); Using(Fmt$(3), CTaxAmt#)
          LineCnt = LineCnt + 1
        End If
        Bills2Print = Bills2Print + 1
        AcctBalance# = UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance
        Print #UBRpt, Tab(5); "Current:"; Using(Fmt$(6), UBBillRec(1).TransAmt);
        If AcctBalance# <> 0 Then
          Print #UBRpt, Tab(30); "Previous:"; Using(Fmt$(6), AcctBalance#);
          TAcctBalance# = Round#(TAcctBalance# + AcctBalance#)
        End If
        Print #UBRpt, Tab(55); "Total:"; Tab(65); Using(Fmt$(6), (Round#(AcctBalance# + UBBillRec(1).TransAmt)))
        LineCnt = LineCnt + 1
        If Not SkipSeparator Then
          Print #UBRpt, Fmt$(0)
          LineCnt = LineCnt + 1
        End If
      End If
      
      If UBBillRec(1).TaxExempt = "Y" Then
        TaxExmp(TRevCnt) = Round#(TaxExmp(TRevCnt) + UBBillRec(1).RevAmt(TRevCnt))
      End If
    Next
    '020199 Moved pump code processing to here. Stops bug in getting true
    '       meter consumption figures.
    GoSub GetWhatPump
    If HasAPumpCode Then
      PumpConsump(WhatPump).CustCnt = PumpConsump(WhatPump).CustCnt + 1
      For MPCnt = 1 To 7
        PumpMtrOK = False
        CubMtr = False
        LocMeterType$ = Trim$(UBCustRec(1).LocMeters(MPCnt).MTRType)
        Select Case LocMeterType$
        Case "C", "S", "W"
          PumpMtrOK = True
        End Select
        If PumpMtrOK Then
          MeterMulti& = UBCustRec(1).LocMeters(MPCnt).MTRMulti
          If UBCustRec(1).LocMeters(MPCnt).MTRUnit = "C" Then
            CubMtr = True
          End If
          If MeterMulti& <= 0 Then MeterMulti& = 1
          ReadAmt& = UBBillRec(1).CurRead(MPCnt) - UBBillRec(1).PrevRead(MPCnt)
          If ReadAmt& < 0 Then  'Meter rolled over or, been misread
            MaxMeterAmt& = 10& ^ (Len(Str$(UBBillRec(1).PrevRead(MPCnt))) - 1)
            ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MPCnt)) + UBBillRec(1).CurRead(MPCnt)
          End If
          If CubMtr Then
            ReadAmt& = ReadAmt& * 7.481
          End If
          PumpConsump(WhatPump).Consump = PumpConsump(WhatPump).Consump + (ReadAmt& * MeterMulti&)
        End If
      Next
    End If
SkipEM:
'add a function to Ask user to abandon Printing
    
'    If AskAbandonPrint% Then
'      UBLog "ABORTED: Prebilling report"
'      UBLog "Closing files."
'      Close
'      AbortFlag = True
'      Exit For
'    End If
    ShowPctComp Cnt, NumOfRecs
  Next
  If AbortFlag Then GoTo ExitPreReport
  
  Print #UBRpt, FF$
  
  GoSub TitleLine
  Print #UBRpt, "Billing Grand Totals"
  If TennFlag Then
    Print #UBRpt, "                                Inactive          Taxed      NONTax     FlatRate"
    Print #UBRpt, "Revenue/Tax        Consump       Consump         Amount      Amount      Amount"
  Else
    Print #UBRpt, "                                 Inactive                             Flat Rate"
    Print #UBRpt, "Revenue/Tax    Consumption      Consumption            Amount           Amount"
  End If
  Print #UBRpt, Fmt$(0)
  
  TotalFlatAmt# = 0
  TotalRevAmt# = 0
  TotalTaxAmt# = 0
  
  For RaCnt = 1 To NumOfRevs
    If TennFlag Then
      Print #UBRpt, UBSetUpRec(1).Revenues(RaCnt).REVNAME; Using(Fmt$(4), ConsumpTot(RaCnt, 1)); Tab(30); Using(Fmt$(4), ConsumpTot(RaCnt, 2));
      If TaxTotals(RaCnt) > 0 Then
        Print #UBRpt, Tab(44); Using(Fmt$(1), Round#(RevTotals(RaCnt) - FlatTotals(RaCnt) - TaxExmp(RaCnt))); Tab(56); Using(Fmt$(1), TaxExmp(RaCnt)); Tab(68); FlatTotals(RaCnt)
      Else
        Print #UBRpt, Tab(44); Using(Fmt$(1), Round#(RevTotals(RaCnt) - FlatTotals(RaCnt))); Tab(68); Using(Fmt$(1), FlatTotals(RaCnt))
      End If
    Else
      Print #UBRpt, UBSetUpRec(1).Revenues(RaCnt).REVNAME; Using(Fmt$(4), ConsumpTot(RaCnt, 1)); Tab(33); Using(Fmt$(4), ConsumpTot(RaCnt, 2));
      Print #UBRpt, Tab(50); Using(Fmt$(1), RevTotals(RaCnt) - FlatTotals(RaCnt)); Tab(67); Using(Fmt$(1), FlatTotals(RaCnt))
    End If
    TotalFlatAmt# = Round#(TotalFlatAmt# + FlatTotals(RaCnt))
    TotalRevAmt# = Round#(TotalRevAmt# + RevTotals(RaCnt))
    If TaxTotals(RaCnt) > 0 Then
      If TennFlag Then
        Print #UBRpt, " Tax"; Tab(44); Using(Fmt$(1), TaxTotals(RaCnt))
      Else
        Print #UBRpt, " Tax"; Tab(50); Using(Fmt$(1), TaxTotals(RaCnt))
      End If
      TotalTaxAmt# = Round#(TotalTaxAmt# + TaxTotals(RaCnt))
    End If
  Next
  Print #UBRpt, Fmt$(0)
  Print #UBRpt, "  PREVIOUS: "; Using(Fmt$(6), TAcctBalance#);
  Print #UBRpt, Tab(32); "REVENUE TOTAL: "; Using(Fmt$(5), Round#(TotalRevAmt# - TotalFlatAmt#))
  Print #UBRpt, "BILL COUNT: "; Using(Fmt$(2), Bills2Print);
  Print #UBRpt, Tab(32); "   FLAT TOTAL: "; Using(Fmt$(5), TotalFlatAmt#)
  Print #UBRpt, Tab(32); "    TAX TOTAL: "; Using(Fmt$(5), TotalTaxAmt#)
  Print #UBRpt, Tab(32); "BILLING TOTAL: "; Using(Fmt$(5), Round#(TotalRevAmt# + TotalTaxAmt#))
  Print #UBRpt, FF$
  
  TotalRevAmt# = 0
  
  GoSub RptTotRateHeader
  
  For RaCnt = 1 To NumOfRates
    If (RateTotals(RaCnt) <> 0) Or (RateConsump(RaCnt) <> 0) Then
      If Len(Trim$(UBRateTbls(RaCnt).RATECODE)) > 0 Then
        Print #UBRpt, UBRateTbls(RaCnt).RATECODE; "    "; UBRateTbls(RaCnt).RATEDESC; Tab(39); Using(Fmt$(4), RateConsump(RaCnt));
        Print #UBRpt, Tab(55); Using(Fmt$(1), RateTotals(RaCnt));
        Print #UBRpt, Tab(69); Using(Fmt$(2), RateCount(RaCnt))
        LineCnt = LineCnt + 1
        TotalRevAmt# = Round#(TotalRevAmt# + RateTotals(RaCnt))
        If RTaxTot(RaCnt) > 0 Then
          Print #UBRpt, " Tax"; Tab(55); Using(Fmt$(1), RTaxTot(RaCnt))
          LineCnt = LineCnt + 1
        End If
        If LineCnt > MaxLines Then
          Print #UBRpt, FF$
          GoSub RptTotRateHeader
        End If
      End If
    End If
  Next
  
  Print #UBRpt, Fmt$(0)
  Print #UBRpt, Tab(36); "TAX TOTAL:"; Tab(52); Using(Fmt$(5), TotalTaxAmt#)
  Print #UBRpt, Tab(40); "TOTAL:"; Tab(52); Using(Fmt$(5), Round#(TotalRevAmt# + TotalTaxAmt#))
  Print #UBRpt, FF$
  
  'SortT BookConsump(1), TBooks, 0, Len(BookConsump(1)), 0, -1
  
  GoSub BookHeader
  
  For Cnt = 1 To TBooks
    TestTot# = 0
    For ZCnt = 1 To NumOfRevs
      TestTot# = Round#(TestTot# + BookConsump(Cnt).RevAmt(ZCnt))
    Next
    If TestTot# <> 0 Then
      If BookConsump(Cnt).Book < 10 Then
        Book$ = "0" + Trim$(Str$(BookConsump(Cnt).Book))
      Else
        Book$ = Trim$(Str$(BookConsump(Cnt).Book))
      End If
      Print #UBRpt, "Book: "; Book$; "    Customers:"; BookConsump(Cnt).CustCnt
      TBookAmt# = 0
      TBTaxAmt# = 0
      For RCnt = 1 To NumOfRevs
        Print #UBRpt, RevDesc(RCnt); Tab(30); Using(Fmt$(4), BookConsump(Cnt).Consump(RCnt));
        Print #UBRpt, Tab(59); Using("##########.##", BookConsump(Cnt).RevAmt(RCnt))
        TBookAmt# = Round#(TBookAmt# + BookConsump(Cnt).RevAmt(RCnt))
        If BookConsump(Cnt).TaxAmt(RCnt) > 0 Then
          Print #UBRpt, " Tax"; Tab(60); Using(Fmt$(1), BookConsump(Cnt).TaxAmt(RCnt))
          TBTaxAmt# = Round#(TBTaxAmt# + BookConsump(Cnt).TaxAmt(RCnt))
          LineCnt = LineCnt + 1
        End If
        LineCnt = LineCnt + 1
      Next
      TBookGTot# = Round#(TBookGTot# + TBookAmt# + TBTaxAmt#)
      Print #UBRpt, Tab(42); "Book Total:"; Tab(57); Using(Fmt$(5), Round#(TBookAmt# + TBTaxAmt#))
      If Cnt < TBooks Then
        Print #UBRpt, Fmt$(0)
      End If
      LineCnt = LineCnt + 1
    End If
    If LineCnt > MaxLines And Cnt < TBooks Then
      Print #UBRpt, FF$
      GoSub BookHeader
    End If
    
SkipThisBook:
  Next
  
  Print #UBRpt, Fmt$(0)
  Print #UBRpt, Tab(35); "Books GRAND Total:"; Tab(57); Using(Fmt$(5), TBookGTot#)
  Print #UBRpt, FF$
  
  If TPumps > 0 Then
    GoSub PumpHeader
    TMMConsump# = 0
    For Cnt = 1 To TPumps
      Print #UBRpt, PumpConsump(Cnt).PumpCode; Tab(30); Using("###########", PumpConsump(Cnt).CustCnt); Tab(60); Using("###########", PumpConsump(Cnt).Consump)
      TMMConsump# = TMMConsump# + PumpConsump(Cnt).Consump
    Next
    Print #UBRpt, Fmt$(0)
    Print #UBRpt, Tab(35); "Pump Code Total:"; Tab(60); Using("###########", TMMConsump#)
  End If
  
  Close
  
  UBLog "Finished writing prebilling report."
  
  Select Case Choice
  Case 1
    RptText$ = "(Customer"
  Case 2
    RptText$ = "(Account"
  Case 3
    RptText$ = "(Location"
  Case 4
    RptText$ = "(Postal RT."
  Case 5
    RptText$ = "(ZipCode"
  Case 6
    RptText$ = "(Sequence"
  End Select
  RptText$ = RptText$ + " Order)"
  
  Erase UBSetUpRec, RevDesc, UBRateTbls, RateConsump
  Erase Fmt$, UBCustRec, UBBillRec, FlatTotals
  Erase RevTotals, TaxTotals, ConsumpTot
  Erase RateTotals, RTaxTot, BookConsump, IndexArray
  Erase RateCount, ProrateServ
  Erase PumpConsump, TaxExmp
  
  If Not AbortFlag Then
    ViewPrint "UBPREBIL.RPT", "Pre-Billing Report " + RptText$
    'PrintRptFile "Pre-Billing Report " + RptText$, "UBPREBIL.RPT", LPTPort, RetCode, EntryPoint
    If BookFlag Then
      Kill UBBillsFile
    End If
  End If
  
  GoTo ExitPreReport
  
PrintPreHeader:
  GoSub TitleLine
  Print #UBRpt, "Stat  Act.  Locat    Customer Name             Service Address       Prorate%"
  Print #UBRpt, "Revenue            R-Code     Cur Read    Pre Read     Consump        Charges"
  Print #UBRpt, Fmt$(0)
  LineCnt = 5
  Return
  
GetWhatBook:
  BadBookFlag = False
  WhatBook = 0
  If Len(Trim$(UBCustRec(1).Book)) = 0 Then
    If UBCustRec(1).Status = "A" Then
      BadBookFlag = True
      'testing vvv
      WhatBook = 0
    End If
    GoTo ErrorBookExit
  End If
  
  ThisBook = Val(UBCustRec(1).Book)
  If TBooks > 0 Then
    For TBCnt = 1 To TBooks
      If BookConsump(TBCnt).Book = ThisBook Then
        WhatBook = TBCnt
        Exit For
      End If
    Next
    If WhatBook = 0 Then
      TBooks = TBooks + 1
      ReDim Preserve BookConsump(0 To TBooks) As BookConsumpType
      BookConsump(TBooks).Book = ThisBook
      WhatBook = TBooks
    End If
  Else
    TBooks = TBooks + 1
    BookConsump(TBooks).Book = ThisBook
    WhatBook = TBooks
  End If
  
ErrorBookExit:
  Return
  
GetWhatPump:
  HasAPumpCode = True           'assume they have a pump code
  WhatPump = 0
  If Len(Trim$(UBCustRec(1).PumpCode)) = 0 Then
    If UBCustRec(1).Status = "A" Then
      HasAPumpCode = False      'no pump code
      WhatPump = 0
    End If
    GoTo PumpCodeReturn
  End If
  
  CustPump$ = UCase$(Trim$(UBCustRec(1).PumpCode))
  If Len(CustPump$) > 0 Then
    For TBCnt = 1 To TPumps
      ThisPump$ = Trim$(PumpConsump(TBCnt).PumpCode)
      If ThisPump$ = CustPump$ Then
        WhatPump = TBCnt
        Exit For
      End If
    Next
    If WhatPump = 0 Then
      TPumps = TPumps + 1
      ReDim Preserve PumpConsump(0 To TPumps) As PumpConsumpType
      PumpConsump(TPumps).PumpCode = CustPump$
      WhatPump = TPumps
    End If
  Else
    TPumps = TPumps + 1
    PumpConsump(TPumps).PumpCode = CustPump$
    WhatPump = TPumps
  End If
  
PumpCodeReturn:
  Return
  
RptTotRateHeader:
  GoSub TitleLine
  Print #UBRpt,
  Print #UBRpt, "Report Totals by Rate Code"
  Print #UBRpt,
  Print #UBRpt, "Code      Rate Description            Consumption           Amount      Bills"
  Print #UBRpt, Fmt$(0)
  LineCnt = 5
  Return
  
BookHeader:
  GoSub TitleLine
  Print #UBRpt, "Report Totals by Book"
  Print #UBRpt,
  Print #UBRpt, "Book"
  Print #UBRpt, "Revenue                      Consumption                         Amount"
  Print #UBRpt, Fmt$(0)
  LineCnt = 7
  Return
  
PumpHeader:
  GoSub TitleLine
  Print #UBRpt, "Report Totals by Pump Code"
  Print #UBRpt,
  Print #UBRpt, "PumpCode                  Customer Count                    Consumption"
  Print #UBRpt, Fmt$(0)
  LineCnt = 6
  Return
  
TitleLine:
  PageNo = PageNo + 1
  Print #UBRpt, "Utility Pre-Billing Report.  "; TownName$; Tab(70); "Page: "; PageNo
  Print #UBRpt, TheDate$
  Return
  
ErrorAbortExit:
  Close
  
ExitPreReport:
  UBLog "OUT: Prebilling Report" + CRLF$
  
End Sub
  
Function PreBillYouSure()
  
'  LibName$ = "UBSETUP"
'  ScrnName$ = "PREBILOK"
'  NumScrns = 1
'
'  '--define the multi-choice fields
'  NumFlds = -1
'  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
'
'  '--define Quick Screen form editing arrays
'  ReDim frm(1) As FormInfo
'  ReDim Form$(NumFlds, 2)
'  ReDim Fld(NumFlds) As FieldInfo
'
'  '--for each screen, get first and last fields
'  StartEl = 0
'  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
'
'  '--Clear all fields
'  For F = 1 To NumFlds
'    LSet Form$(F, 0) = ""
'  Next
'
'  '--Set screen number to one and display screen
'  Scr = 1
'  BlockClear
'  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
'
'  ShowCursor
'  Action = 1
'
'  Do
'
'    EditForm Form$(), Fld(), frm(1), Cnf, Action
'    '--Check for Key presses
'    Select Case frm(1).KeyCode
'    Case -68    'F10Key
'      OKFlag = True
'      Exit Do
'    Case EscKey
'      ExitFlag = True
'    End Select
'  Loop Until frm(1).KeyCode = 27 Or ExitFlag
'
'  If Not OKFlag Then
'    PreBillYouSure = False
'  Else
'    PreBillYouSure = True
'  End If
'
End Function
  
'Sub RateCodeErrScrn(RATECODE$)
'  ReDim TempArray(0) As Integer
'  SaveScrn TempArray()
'  BlockClear
'  DisplayUBScrn "ERRSCRN1"
'  QPrintRC "RATE CODE:  " + Trim$(RATECODE$), 10, 22, -1
'  QPrintRC "Has an INVALID entry!", 10, 39, -1
'  QPrintRC "Correct and Print Again.", 12, 28, -1
'  WaitForAction
'  AbortFlag = True
'  RestScrn TempArray()
'  Erase TempArray
'End Sub
  

  Type FLen2
    V As String * 64
  End Type
  Dim PctC(1) As String * 3

Function Chk4DupeBookSeqNum(Book$, SeqNum$)
  
  Chk4DupeBookSeqNum = False    'assume it's ok
  TBookSeq& = QPValL(Book$ + SeqNum$)
  ReDim UBBookSeq(1) As BookSeqRecType
  BookSeqLen = Len(UBBookSeq(1))

  If FileSize("UBOOKSEQ.DAT") > 0 Then
    FOpenS "UBOOKSEQ.DAT", Handle               'open data file
    NumBookSeq = FLof(Handle) \ BookSeqLen
    ReDim UBBookSeq(1 To NumBookSeq) As BookSeqRecType
    FGetRTA Handle, UBBookSeq(1), 1&, NumBookSeq * BookSeqLen
    FClose Handle
    For Cnt = 1 To NumBookSeq
      If UBBookSeq(Cnt).BookSeq = TBookSeq& Then
        Ok = MsgBox%("UB.QSL", "DUPEBOOK")
        Chk4DupeBookSeqNum = True
        Exit For
      End If
    Next
  End If

  Erase UBBookSeq

End Function

Function ChkBillFile%()
  
  OKFlag = True 'assume all is well
  
  ReDim BillRec(1) As UBTransRecType
  RecLen = Len(BillRec(1))
  
  FHand = FreeFile
  Open UBBillsFile For Random Shared As FHand Len = RecLen
  NumOfRec& = LOF(FHand) \ RecLen
  Close FHand

  If NumOfRec& = 0 Then
    Kill UBBillsFile
    OKFlag = False
  End If
  
  ChkBillFile% = OKFlag

  Erase BillRec
  
End Function



Function CustHasMsg(RecNo&)

  ReDim MsgRec(1) As UBMessRecType
  MsgLen = Len(MsgRec(1))

  NumMsgRec& = FileSize&("UBMESAGE.DAT") / MsgLen

  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  
  If RecNo& > 0 Then
    UBFile = FreeFile
    Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
    Get UBFile, RecNo&, UBCustRec(1)
    Close UBFile
    MRec& = UBCustRec(1).MessageRec

    If MRec& > 0 And MRec& <= NumMsgRec& Then
      MsgFile = FreeFile
      Open "UBMESAGE.DAT" For Random Shared As MsgFile Len = MsgLen
      Get MsgFile, MRec&, MsgRec(1)
      Close MsgFile
      For zz = 1 To 15
        m$ = Trim$(MsgRec(1).MessLine(zz).Line)
        If Len(m$) > 0 Then
          GotMsg = True
          Exit For
        End If
      Next
    Else
      GotMsg = False
    End If
  Else
    GotMsg = False
  End If

  If GotMsg Then
    CustHasMsg = True
  Else
    CustHasMsg = False
  End If

End Function

Sub CustMessageSystem(RecNo&)
  
  CustRec& = RecNo&
  
  ReDim ScrnArray(0)
  
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  
  ReDim UBMessRec(1) As UBMessRecType
  UBMessRecLen = Len(UBMessRec(1))
  
  UBCust = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  Get UBCust, CustRec&, UBCustRec(1)
  Close
  
  LibName$ = "UB"
  ScrnName$ = "UBCUSMES"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  ReDim frm(1) As FormInfo
  ReDim Form$(NumFlds, 2)
  ReDim Fld(NumFlds) As FieldInfo
  frm(1).StayOnField = True
  
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  FirstTime = True
  
  Action = 1
  
  DisplayUBScrn ScrnName$
  QPrintRC Str$(CustRec&), 3, 20, -1
  QPrintRC UBCustRec(1).CUSTNAME, 4, 20, -1
  QPrintRC UBCustRec(1).Status, 3, 67, -1
  
  Do
    EditForm Form$(), Fld(), frm(1), Cnf, Action
    
    If FirstTime Then
      FirstTime = False
      GoSub LoadMessageInfo
      Action = 1
    End If
    
    Select Case frm(1).KeyCode
    Case F3Key
      GoSub ClearRecord
      GoSub ClearForm
      Action = 1
    Case F5Key
      GoSub SaveRecord
      GoSub PrintMessage
    Case F10Key
      SaveScrn ScrnArray()
      DisplayUBScrn "UPDATDSK"
      GoSub SaveRecord
      RestScrn ScrnArray()
      DisplayUBScrn "UPDATEOK"
      WaitForAction
      ExitFlag = True
      RestScrn ScrnArray()
      Done = True
    Case ESC
      Exit Sub
    Case Else
      Done = False
    End Select
  Loop Until Done
  
ExitMessageInquiry:
  Exit Sub
  '***************
  
LoadMessageInfo:
  MessageRecord = UBCustRec(1).MessageRec
  If MessageRecord > 0 Then
    UBMess = FreeFile
    Open "UBMESAGE.DAT" For Random Shared As UBMess Len = UBMessRecLen
    Get UBMess, MessageRecord, UBMessRec(1)
    Close
    BCopy VARSEG(UBMessRec(1)), VarPtr(UBMessRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), UBMessRecLen, 0
    Call UnPackBuffer(0, 0, Form$(), Fld())
  End If
Return
  
SaveRecord:
  UBMess = FreeFile
  Open "UBMESAGE.DAT" For Random Shared As UBMess Len = UBMessRecLen
  If MessageRecord = 0 Then
    MessageRecord = LOF(UBMess) / Len(UBMessRec(1)) + 1
  End If
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBMessRec(1)), VarPtr(UBMessRec(1)), UBMessRecLen, 0
  Put UBMess, MessageRecord, UBMessRec(1)
  Close
  
  UBCust = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
  Get UBCust, CustRec&, UBCustRec(1)
  UBCustRec(1).MessageRec = MessageRecord
  Put UBCust, CustRec&, UBCustRec(1)
  Close
Return
  
ClearRecord:
  If MessageRecord > 0 Then
    ReDim UBMessRec(1) As UBMessRecType
    UBMess = FreeFile
    Open "UBMESAGE.DAT" For Random Shared As UBMess Len = UBMessRecLen
    Put UBMess, MessageRecord, UBMessRec(1)
    Close
  End If
Return
  
ClearForm:
  For F = 1 To NumFlds
    LSet Form$(F, 0) = ""
  Next F
Return

PrintMessage:
  SaveScrn ScrnArray()
  Dash$ = String$(80, "-")
  ReDim UBSetUpRec(1) As UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  TownName$ = UBSetUpRec(1).UTILNAME
  Erase UBSetUpRec
  UBRpt = FreeFile
  Open "UBCUSMSG.RPT" For Output As UBRpt

  Print #UBRpt, "Customer Messages Listing."; Tab(64); "Date: "; Date$
  Print #UBRpt, "NAME: "; UBCustRec(1).CUSTNAME; "Acct:"; Str$(CustRec&)
  Print #UBRpt, "Message Text"; Tab(70); "Entry Date"
  Print #UBRpt, Dash$
  For MsgLine = 1 To 15
    Print #UBRpt, UBMessRec(1).MessLine(MsgLine).Line; Tab(70); UBMessRec(1).MessLine(MsgLine).LineDate
  Next
  Print #UBRpt, Dash$
  Print #UBRpt, Chr$(12)
  Close UBRpt
  PrintRptFile "Customer Message Listing.", "UBCUSMSG.RPT", 1, RetCode, EntryPoint
  RestScrn ScrnArray()
  Action = 1

Return
  
End Sub

Sub DisplayUBScrn(ScrnName$)
  LibFile2Scrn "UB", ScrnName$, MonoCode, Attribute%, ErrCode
End Sub

Function FmtBook$(Book$)
  Book$ = Trim$(Book$)
  BookLen = Len(Book$)
  
  Select Case BookLen
  Case 0
    FmtBook$ = "00"
  Case 1
    FmtBook$ = "0" + Book$
  Case Else
    FmtBook$ = Book$
  End Select
  
End Function

Function FmtSeqN$(SeqN$)
  
  SeqN$ = Trim$(SeqN$)
  SeqNLen = Len(SeqN$)
  
  Select Case SeqNLen
  Case 0
    FmtSeqN$ = "000000"
  Case 1 To 5
    FmtSeqN$ = "000000"
    Mid$(FmtSeqN$, (6 - SeqNLen) + 1) = SeqN$
  Case Else
    FmtSeqN$ = SeqN$
  End Select
  
End Function

Function GetCustMeterType(UBCustRec() As NewUBCustRecType, ThisMeter)
  Dim LMtrType As String
  Dim LMtrTypeLen As Integer, LThisMeter As Integer
  
  'Meter Types
  'CONST MtrWaterOnly = 1
  'CONST MtrSewerOnly = 2
  'CONST MtrCombined = 3
  'CONST MtrElectric = 4
  'CONST MtrDemand = 5
  'CONST MtrGas = 6
  'CONST MtrTouchRead = 7
  
  LMtrType$ = Trim$(UBCustRec(1).LocMeters(ThisMeter).MTRType)
  LMtrTypeLen = Len(LMtrType$)
  If LMtrTypeLen > 0 Then
    Select Case LMtrType$
    Case "W"
      LThisMeter = MtrWaterOnly
    Case "S"
      LThisMeter = MtrSewerOnly
    Case "C"
      LThisMeter = MtrCombined
    Case "E"
      LThisMeter = MtrElectric
    Case "D"
      LThisMeter = MtrDemand
    Case "G"
      LThisMeter = MtrGas
    Case "T"
      LThisMeter = MtrTouchRead
    Case Else
      LThisMeter = True
    End Select
    GetCustMeterType = LThisMeter
  Else
    GetCustMeterType = 0
  End If
  
End Function

'This function returns the number of customer records
Function GetNumOfCust()
  ReDim TCustRec(1) As NewUBCustRecType
  RecLen = Len(TCustRec(1))
  CFileSize& = FileSize("UBCUST.DAT")
  GetNumOfCust = CFileSize& \ RecLen
  Erase TCustRec
End Function

'This function return the number of rate codes
Function GetNumRateRecs()
  Dim UBRateTblRecLen As Integer
  ReDim UBRateTblRec(1) As UBRateTblRecType
  UBRateTblRecLen = Len(UBRateTblRec(1))
  GetNumRateRecs = FileSize("UBRATE.DAT") \ UBRateTblRecLen
  Erase UBRateTblRec
End Function

Function GetZipEDigit$(Zip$)
  
  ZipLen = Len(Zip$)
  ZipVal = 0
  
  DashPos = InStr(Zip$, "-")
  Do While DashPos
    Zip$ = Left$(Zip$, DashPos - 1) + Mid$(Zip$, DashPos + 1)
    DashPos = InStr(Zip$, "-")
  Loop
  
  For Cnt = 1 To ZipLen
    ZipVal = ZipVal + Val(Mid$(Zip$, Cnt, 1))
  Next
  
  If ZipVal Mod 10 > 0 Then
    Dif = 10 - (ZipVal Mod 10)
  Else
    Dif = 0
  End If
  
  GetZipEDigit$ = Trim$(Str$(Dif))
  
End Function

'Returns TRUE if this is a deleted account
Function IsDeleted%(AcctNo&)
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))
  FOpenS "UBCUST.DAT", C1Handle
  FGetRTA C1Handle, UBCustRec(1), AcctNo&, UBCustRecLen
  FClose C1Handle
  If UBCustRec(1).DelFlag <> 0 Then
    IsDeleted% = True
  Else
    IsDeleted% = False
  End If
  Erase UBCustRec
End Function

Sub LoadUBSetUpFile(UBSetUpRec() As UBSetupRecType, UBSetUpLen)
                       'use the length as an error flag
'  UBSetupLen = -1      'assume the file is not there, or 0 bytes.
  Dim FileHandle As Integer
  FileHandle = FreeFile
  UBSetUpLen = Len(UBSetUpRec(1))
    If Exist("UBSETUP.DAT") Then
    Open "UBSETUP.DAT" For Random As #FileHandle Len = UBSetUpLen              'open data file
    Get #FileHandle, 1, UBSetUpRec(1)
    Close FileHandle
  End If
  
End Sub


Function MakeMonth$(TDate$)
  Month = Val(Left$(TDate$, 2))
  Select Case Month
  Case 1
    MakeMonth$ = "January"
  Case 2
    MakeMonth$ = "February"
  Case 3
    MakeMonth$ = "March"
  Case 4
    MakeMonth$ = "April"
  Case 5
    MakeMonth$ = "May"
  Case 6
    MakeMonth$ = "June"
  Case 7
    MakeMonth$ = "July"
  Case 8
    MakeMonth$ = "August"
  Case 9
    MakeMonth$ = "September"
  Case 10
    MakeMonth$ = "October"
  Case 11
    MakeMonth$ = "November"
  Case 12
    MakeMonth$ = "December"
  End Select
End Function

Sub MakeMowZipCodeIndex(IndexText$)
  
  ShowProcessingScrn "Creating " + IndexText$ + " Index "
  QPrintRC "    Reading Customer Records     ", 11, 25, -1

  ReDim UBCustRec(1) As NewUBCustRecType
  CustRecLen = Len(UBCustRec(1))

  NumOfBillRec = FileSize("UBCUST.DAT") \ CustRecLen

  CHandle = FreeFile
  Open "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen

  ReDim ZipIndex(1 To NumOfBillRec) As MOWZipIndexType
  For BCnt = 1 To NumOfBillRec
    Get CHandle, BCnt, UBCustRec(1)
    ZipIndex(BCnt).ZIPCODE = UBCustRec(1).ZIPCODE
    ZipIndex(BCnt).RecNum = BCnt
    ShowPctComp BCnt, NumOfBillRec              'show user percentage complete
  Next
  Close

  QPrintRC "         Sorting Index.        ", 11, 25, -1
  SortT ZipIndex(1), NumOfBillRec, 0, 16, 0, 10
  QPrintRC "      Writing Index Records      ", 11, 25, -1

  IHandle = FreeFile
  Open TempIndexName For Output As IHandle
  Close IHandle

  IHandle = FreeFile
  Open TempIndexName For Random Shared As IHandle Len = 4
  For Cnt = 1 To NumOfBillRec
    Prec& = ZipIndex(Cnt).RecNum
    Put IHandle, Cnt, Prec&
    ShowPctComp Cnt, NumOfBillRec               'show user percentage complete
  Next
  Close IHandle

  Erase UBCustRec, ZipIndex

End Sub

Sub MakePostalIndex(IndexText$)
  
  ShowProcessingScrn "Creating " + IndexText$ + " Index"
  QPrintRC "    Reading Customer Records     ", 11, 25, -1
  
  ReDim UBCustRec(1) As NewUBCustRecType
  CustRecLen = Len(UBCustRec(1))
  
  NumCustRecs = GetNumOfCust%
  
  ReDim PostalIndex(1 To NumCustRecs) As UBPostalIndexType
  IndexRecLen = Len(PostalIndex(1))
  
  CHandle = FreeFile
  Open "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
  For Cnt = 1 To NumCustRecs
    Get CHandle, Cnt, UBCustRec(1)
    PostalIndex(Cnt).ZIPCODE = UBCustRec(1).ZIPCODE
    RSet PostalIndex(Cnt).Route = Trim$(UBCustRec(1).POSTRTE)
    PostalIndex(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumCustRecs                'show user percentage complete
  Next
  
  Close CHandle
  QPrintRC "         Sorting Index.        ", 11, 25, -1
  SortT PostalIndex(1), NumCustRecs, 0, 16, 10, 4
  QPrintRC "      Writing Index Records      ", 11, 25, -1
  IHandle = FreeFile

  FCreate TempIndexName
  
  Open TempIndexName For Random Shared As IHandle Len = 4
  For Cnt = 1 To NumCustRecs
    Prec& = PostalIndex(Cnt).RecNum
    Put IHandle, Cnt, Prec&
    ShowPctComp Cnt, NumCustRecs                'show user percentage complete
  Next
  Close IHandle
  
  Erase UBCustRec, PostalIndex
  
End Sub

Sub MakeSequenceIndex(IndexText$)
  ShowProcessingScrn "Creating " + IndexText$ + " Index"
  QPrintRC "    Reading Location Records     ", 11, 25, -1
  
  ReDim UBCustRec(1) As NewUBCustRecType
  CustRecLen = Len(UBCustRec(1))
  
  NumCustRecs& = GetNumOfCust%
  
  ReDim SequenceIndex(1 To NumCustRecs&) As UBSequenceIndexType
  IndexRecLen = Len(SequenceIndex(1))
  
  CHandle = FreeFile
  Open "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
  For Cnt = 1 To NumCustRecs&
    Get CHandle, Cnt, UBCustRec(1)
    SequenceIndex(Cnt).SeqNumber = UBCustRec(1).SEQ
    SequenceIndex(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumCustRecs&               'show user percentage complete
  Next
  
  Close CHandle
  
  QPrintRC "         Sorting Index.        ", 11, 25, -1
  
  SortT SequenceIndex(1), CInt(NumCustRecs&), 0, 16, 0, -2
  
  QPrintRC "      Writing Index Records      ", 11, 25, -1
  
  FCreate TempIndexName
  
  IHandle = FreeFile
  Open TempIndexName For Random Shared As IHandle Len = 4
  
  For Cnt = 1 To NumCustRecs&
    Prec& = SequenceIndex(Cnt).RecNum
    Put IHandle, Cnt, Prec&
    ShowPctComp Cnt, NumCustRecs&               'show user percentage complete
  Next
  Close IHandle
  
  Erase UBCustRec, SequenceIndex

End Sub

Sub MakeZipCodeIndex(IndexText$)
  
  ShowProcessingScrn "Creating " + IndexText$ + " Index "
  QPrintRC "    Reading Customer Records     ", 11, 25, -1
  
  'REDIM ZipIndex(1 TO 1)  AS PSAZipIndexType
  ReDim UBCustRec(1) As NewUBCustRecType
  CustRecLen = Len(UBCustRec(1))
  
  NumOfBillRec = FileSize("UBCUST.DAT") \ CustRecLen
  
  CHandle = FreeFile
  Open "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
  
  ReDim ZipIndex(1 To NumOfBillRec) As PSAZipIndexType
  
  For BCnt = 1 To NumOfBillRec
    Get CHandle, BCnt, UBCustRec(1)
    ZipIndex(BCnt).ZIPCODE = UBCustRec(1).ZIPCODE
    ZipIndex(BCnt).SName = UBCustRec(1).SEARCH
    ZipIndex(BCnt).RecNum = BCnt
    ShowPctComp BCnt, NumOfBillRec              'show user percentage complete
  Next
  
  Close
  
  QPrintRC "         Sorting Index.        ", 11, 25, -1
  
  SortT ZipIndex(1), NumOfBillRec, 0, 32, 0, 10
  
  First = 1
  Last = 1
  
  SZip$ = ZipIndex(1).ZIPCODE
  
  For ZCnt = 2 To NumOfBillRec
    EZip$ = ZipIndex(ZCnt).ZIPCODE
    If SZip$ <> EZip$ Then
      Last = ZCnt - 1
      GoSub SortThisZip
      First = ZCnt
      SZip$ = EZip$
    End If
    ShowPctComp ZCnt, NumOfBillRec              'show user percentage complete
  Next
  Last = ZCnt - 1
  GoSub SortThisZip
  
  QPrintRC "      Writing Index Records      ", 11, 25, -1
  
  IHandle = FreeFile
  Open TempIndexName For Output As IHandle
  Close IHandle
  
  IHandle = FreeFile
  Open TempIndexName For Random Shared As IHandle Len = 4
  For Cnt = 1 To NumOfBillRec
    Prec& = ZipIndex(Cnt).RecNum
    Put IHandle, Cnt, Prec&
    ShowPctComp Cnt, NumOfBillRec               'show user percentage complete
  Next
  Close IHandle
  
  Erase UBCustRec, ZipIndex
  
  Exit Sub
  
SortThisZip:
  If First < Last Then
    SortT ZipIndex(First), Last - First + 1, 0, 32, 10, 10
  End If
Return
  
End Sub

Function PromptSaveData%()
  
  ReDim TempScrn(0)
  SaveScrn TempScrn()
  
  LibName$ = "UB"
  SaveFlag = 2
  
  FormName$ = "SAVE1ST"
  NumFlds = LibNumberOfFields(LibName$, FormName$)
  
  ReDim frm(1) As FormInfo
  ReDim Form$(NumFlds, 2)       'DIM the form data array
  ReDim Fld(NumFlds) As FieldInfo               'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef LibName$, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  '----- Setup TYPE for setting and reading form editing information.
  frm(1).FldNo = 1              'Start editing on field #1
  frm(1).InsStat = False        'Set insert state (True = Insert on)
  frm(1).StartEl = 0            'Set form starting element to 0 and
  
  DisplayUBScrn FormName$
  
  Do
    EditForm Form$(), Fld(), frm(1), Cnf, Action
    Select Case frm(1).KeyCode
    Case F0Key
      SaveFlag = True
    Case EscKey
      SaveFlag = 1
    Case 88, 120                'X Key
      SaveFlag = False
    End Select
    
  Loop While SaveFlag = 2       'proper key not set
  
  PromptSaveData = SaveFlag
  CursorOff
  
  RestScrn TempScrn()
  
  Erase TempScrn, Form$, Fld, frm
  
End Function

Sub ReIndexSystem(PromptFlag%)
  
  UBLog " IN: Reindex Utility Files"
  
  BlockClear
  If PromptFlag% Then
    Ok = MsgBox%("UB", "MUSTEXIT")
    Select Case Ok
    Case 2
      GoTo ExitReindex
    End Select
  End If
  
  'BlockClear
  ReDim UBCustRec(1) As NewUBCustRecType
  UBCustRecLen = Len(UBCustRec(1))              'Length of Cust Record Structure
  
  ReDim UBTransRec(1) As UBTransRecType
  UBTranRecLen = Len(UBTransRec(1))             'Length of Tran Record Structure
  
  ShowProcessingScrn "Reading Customer Names"
  UBLog "BEGIN: Customer Name Reindex"
  UBFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  
  ReDim IdxBuff(1 To NumOfRecs&) As nUBCustReIndexRecType
  
  For Cnt = 1 To NumOfRecs&
    Get UBFile, Cnt, UBCustRec(1)
    IdxBuff(Cnt).SearchName = UBCustRec(1).SEARCH
    If UBCustRec(1).DelFlag Then
      IdxBuff(Cnt).DelFlag = "Y"
    Else
      IdxBuff(Cnt).DelFlag = ""
    End If
    IdxBuff(Cnt).Status = UBCustRec(1).Status
    IdxBuff(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumOfRecs&
  Next
  
  Close UBFile
  
  QPrintRC " Sorting Customer Names", 11, 29, -1
  
  SortT IdxBuff(1), CInt(NumOfRecs&), 0, 16, 0, 10
  
  GoSub ClearBlock
  QPrintRC "Writing Customer Index", 9, 30, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1
  
  KillFile "UBCUSTNM.IDX"
  UBFile = FreeFile
  Open "UBCUSTNM.IDX" For Random Shared As UBFile Len = 4
  For Cnt = 1 To NumOfRecs&
    Put UBFile, Cnt, IdxBuff(Cnt).RecNum
    ShowPctComp Cnt, NumOfRecs&
  Next
  Close UBFile
  
  GoSub ClearBlock
  QPrintRC "Writing Customer Search Data", 9, 27, 126
  QPrintRC "Processing:    % Complete", 13, 28, -1
  
  KillFile "UBCUSTSN.DAT"
  UBFile = FreeFile
  Open "UBCUSTSN.DAT" For Random Shared As UBFile Len = Len(IdxBuff(1))
  For Cnt = 1 To NumOfRecs&
    Put UBFile, Cnt, IdxBuff(Cnt)
    ShowPctComp Cnt, NumOfRecs&
  Next
  Close UBFile
  
  Erase IdxBuff
  UBLog "FINISH: Customer Name Reindex"
  GoSub ClearBlock
  
  QPrintRC "Reading Location Information", 9, 27, 126
  QPrintRC "Processing:    % Complete", 13, 28, -1
  UBLog "BEGIN: Book\Sequence Reindex"
  
  UBFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  
  ReDim LIdxBuff(1 To NumOfRecs&) As UBLocaReIndexRecType
  
  For Cnt = 1 To NumOfRecs&
    Get UBFile, Cnt, UBCustRec(1)
    LIdxBuff(Cnt).Book = UBCustRec(1).Book
    LIdxBuff(Cnt).SEQNUMB = UBCustRec(1).SEQNUMB
    LIdxBuff(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumOfRecs&
  Next
  
  Close UBFile
  
  QPrintRC " Sorting Locations Names", 11, 29, -1
  
  SortT LIdxBuff(1), CInt(NumOfRecs&), 0, 16, 0, 8
  'Array(1), NumElem, Dir, StructSize, MemOff, MemSize
  
  GoSub ClearBlock
  QPrintRC "Writing Location Index", 9, 30, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1
  'here
  KillFile "UBCUSTBK.IDX"
  
  UBFile = FreeFile
  Open "UBCUSTBK.IDX" For Random Shared As UBFile Len = 4
  
  For Cnt = 1 To NumOfRecs&
    Put UBFile, Cnt, LIdxBuff(Cnt).RecNum
    ShowPctComp Cnt, NumOfRecs&
  Next
  Close UBFile
  
  UBLog "FINISH: Book\Sequence Reindex"
  ReDim BookSeq(1) As BookSeqRecType
  
  KillFile "UBOOKSEQ.DAT"
  UBLog "BEGIN: Rebuild Book\Sequence List"
  BookHand = FreeFile
  Open "UBOOKSEQ.DAT" For Random Shared As BookHand Len = 4
  For Cnt = 1 To NumOfRecs&
    BookSeq(1).BookSeq = QPValL(LIdxBuff(Cnt).Book + LIdxBuff(Cnt).SEQNUMB)
    Put BookHand, Cnt, BookSeq(1)
  Next
  Close BookHand
  UBLog "FINISH: Rebuild Book\Sequence List"

  Erase LIdxBuff, BookSeq, IdxBuff
  Erase UBCustRec, UBTransRec
  
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
ExitReindex:
  UBLog "OUT: Reindex Utility Files" + CRLF$
  Exit Sub
  
ClearBlock:
  HideCursor
  Blank$ = Space$(40)
  For Cnt = 8 To 15
    QPrintRC Blank$, Cnt, 21, -1
  Next
  ShowCursor
Return
  
End Sub
'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
Public Function Round#(ByVal N#)
  Round# = (Int(N# * 100 + 0.5000001)) / 100
End Function

Sub ShowCustConsHist(CustRec&)
  
  ReDim TempScrn(0)
  SaveScrn TempScrn()
  
  ReDim Metered(1 To 15)
  ReDim UBSetUpRec(1) As UBSetupRecType
  UBSetUpLen = Len(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetUpLen, 1            'load it
  If InStr(UBSetUpRec(1).UTILNAME, "TROY") > 0 Then
    TroyFlag = True
  End If
  If InStr(UBSetUpRec(1).UTILNAME, "HAMLET") > 0 Then
    HamFlag = True
  End If

  NumOfRevs = MaxRevsCnt
  For RevCnt = 1 To 15
    RLen = Len(Trim$(Left$(UBSetUpRec(1).Revenues(RevCnt).REVNAME, 14)))
    If RLen >= 0 Then
      NumOfRevs = RevCnt - 1
      Exit For
    End If
    If UBSetUpRec(1).Revenues(RevCnt).UseMtr = "Y" Then
      Metered(RevCnt) = True
    End If
  Next
  
  ReDim MChoice(1 To 1) As FLen2
  ReDim UBTranRec(1) As UBTransRecType
  ReDim UBCustRec(1) As NewUBCustRecType
  
  UBCustRecLen = Len(UBCustRec(1))
  UBTranRecLen = Len(UBTranRec(1))
  
  UBFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  Get UBFile, CustRec&, UBCustRec(1)
  Close UBFile
  
  CurBal# = UBCustRec(1).CurrBalance
  PreBal# = UBCustRec(1).PrevBalance
  
  UBTran = FreeFile
  Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen
  
  PrevTranRec& = UBCustRec(1).LastTrans
  
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      Get UBTran, PrevTranRec&, UBTranRec(1)
      If UBTranRec(1).TransType = TranUtilityBill Or UBTranRec(1).TransType = TranUtilityBill + 100 Then
        For MtrCnt = 1 To 7
          If UBTranRec(1).MtrTypes(MtrCnt) <> 0 Then
            DCnt = DCnt + 1
            ReDim Preserve MChoice(1 To DCnt) As FLen2
            If HamFlag Then
              LSet MChoice(DCnt).V = Num2Date(UBTranRec(1).ReadDate)
            Else
              LSet MChoice(DCnt).V = Num2Date(UBTranRec(1).TransDate)
            End If
            Select Case UBTranRec(1).MtrTypes(MtrCnt)
            Case MtrWaterOnly
              MeterType$ = "Water"
            Case MtrSewerOnly
              MeterType$ = "Sewer"
            Case MtrCombined
              MeterType$ = "Combined"
            Case MtrElectric
              MeterType$ = "Electric"
            Case MtrDemand
              MeterType$ = "D Electric"
            Case MtrGas
              MeterType$ = "Gas Meter"
            Case MtrTouchRead
              MeterType$ = "Touch Read"
            Case MtrLightsService
              MeterType$ = "L Service"
            Case -1
              MeterType$ = "L Service"
            End Select

            Mid$(MChoice(DCnt).V, 13) = MeterType$
            Mid$(MChoice(DCnt).V, 26) = FUsing$(Str$(UBTranRec(1).CurRead(MtrCnt)), "##########")
            Mid$(MChoice(DCnt).V, 38) = FUsing$(Str$(UBTranRec(1).PrevRead(MtrCnt)), "##########")
            MeterConsp& = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
            If MeterConsp& < 0 Then
              MaxMeterAmt& = 10& ^ (Len(Str$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
              MeterConsp& = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
            End If
'working here
            MTRMulti# = 0
            For MCnt = 1 To 7
              If UBTranRec(1).MtrTypes(MtrCnt) = GetCustMeterType%(UBCustRec(), MCnt) Then
                MTRMulti# = UBCustRec(1).LocMeters(MCnt).MTRMulti
                If UBCustRec(1).LocMeters(MCnt).MTRUnit = "C" Then
                  MeterConsp& = MeterConsp& * 7.481
                  Exit For
                End If
              End If
            Next
            If MTRMulti# = 0 Then
              If TroyFlag Then
                MTRMulti# = 100
              Else
                MTRMulti# = 1
              End If
            End If
            Mid$(MChoice(DCnt).V, 52) = FUsing$(Str$(MTRMulti# * MeterConsp&), "##########")
          End If
        Next
      End If
      PrevTranRec& = UBTranRec(1).PrevTrans
    Loop
    
    Close UBTran
    RestScrn TempScrn()
    MPaintBox 3, 5, 22, 75, 8
    
    MaxLen = 62 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    If Choice < 1 Then
      Choice = 1                'Pre-load choice to highlight
    End If
    
    Title$ = Space$(MaxLen + 4)
    Balance$ = Title$
    LSet Title$ = " Trans Date   Meter Type      Current    Previous    Consumption"
    
    '--Find max menu width
    '--Center Menu within Screen
    
    Row = 4
    col = 8
    Row = 6
    BoxBot = 17 'limit the box length to go no lower than line 20
    
    TitleBox BoxBot + 3, col, MaxLen + 3, "Press <ESC> to continue.", Cnf
    
    QPrintRC Title$, Row - 1, col, 112
    MPaintBox Row, col + MaxLen + 4, Row, col + MaxLen + 5, 8
    
    Do
      LOCATE Row, col, 0
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      If Ky$ = Chr$(27) Then
        RestScrn TempScrn()
        Exit Do
      End If
    Loop
  Else
    Close UBTran
    Ok = MsgBox%("UB.QSL", "NOCTRANS")
    RestScrn TempScrn()
  End If
  
  RestScrn TempScrn()
  Erase Metered, UBSetUpRec, MChoice
  Erase TempScrn, UBTranRec, UBCustRec
  
Exit Sub
  
  
End Sub

Sub ShowCustHistory(CustRec&)
  
  u$ = Chr$(24)
  d$ = Chr$(25)
  
  ReDim TempScrn(0)
  SaveScrn TempScrn()
  
  DisplayUBScrn "UBCUHIST"
  
  ReDim RevText$(1 To MaxRevsCnt)
  ReDim Metered(1 To 15)
  ReDim UBSetUpRec(1) As UBSetupRecType
  UBSetUpLen = Len(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetUpLen, 1            'load it
  NumOfRevs = MaxRevsCnt
  For RevCnt = 1 To 15
    RevText$(RevCnt) = Left$(Trim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME), 14)
    If Len(RevText$(RevCnt)) = 0 Then
      NumOfRevs = RevCnt - 1
      Exit For
    End If
    If UBSetUpRec(1).Revenues(RevCnt).UseMtr = "Y" Then
      Metered(RevCnt) = True
    End If
  Next
  
  ReDim MChoice(1 To 1) As FLen2
  
  ReDim UBTranRec(1) As UBTransRecType
  ReDim UBCustRec(1) As NewUBCustRecType
  
  UBCustRecLen = Len(UBCustRec(1))
  UBTranRecLen = Len(UBTranRec(1))
  
  UBFile = FreeFile
  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
  Get UBFile, CustRec&, UBCustRec(1)
  Close UBFile
  
  CurBal# = UBCustRec(1).CurrBalance
  PreBal# = UBCustRec(1).PrevBalance
  
Top:
  
  UBTran = FreeFile
  Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen
  
  PrevTranRec& = UBCustRec(1).LastTrans
  
  If PrevTranRec& > 0 Then
    Do While PrevTranRec& > 0
      DCnt = DCnt + 1
      ReDim Preserve MChoice(1 To DCnt) As FLen2
      Get UBTran, PrevTranRec&, UBTranRec(1)
      LSet MChoice(DCnt).V = Num2Date(UBTranRec(1).TransDate)
      'MID$(MChoice(DCnt).V, 15) = UBTranRec(1).TransDesc
      GoSub GetTransType
      Mid$(MChoice(DCnt).V, 13) = TType$
      Mid$(MChoice(DCnt).V, 41) = FUsing(Str$(UBTranRec(1).TransAmt), "#####.##")
      'this will show th actual trans number in the list
      'MID$(MChoice(DCnt).V, 50) = FUsing(STR$(PrevTranRec&), "######")
      Mid$(MChoice(DCnt).V, 52) = FUsing(Str$(UBTranRec(1).RunBalance), "#####.##")
      Mid$(MChoice(DCnt).V, 61) = MKL$(PrevTranRec&)
      PrevTranRec& = UBTranRec(1).PrevTrans
    Loop
    
    Close UBTran
    
    RestScrn TempScrn()
    MPaintBox 3, 5, 22, 75, 8
    ReDim TempScrn2(0)
    SaveScrn TempScrn2()
    
HistTop:
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    If Choice < 1 Then
      Choice = 1                'Pre-load choice to highlight
    End If
    
    Title$ = Space$(MaxLen + 4)
    Balance$ = Title$
    LSet Title$ = "  Trans Date       Description           Trans Amt    Balance  "
    LSet Balance$ = " Balance:" + FUsing(Str$(CurBal# + PreBal#), ",#####.##") + "   Cur:" + FUsing(Str$(CurBal#), ",#####.##") + "  Prev:" + FUsing(Str$(PreBal#), ",#####.##")
    
    '--Find max menu width
    '--Center Menu within Screen
    
    Row = 4
    col = ((80 - 60) \ 2) - 1
    
    Row = 6
    BoxBot = 17 'limit the box length to go no lower than line 20
    
    'TitleBox BoxBot + 3, Col, MaxLen + 3, "       Press <ESC> to continue.", Cnf
    
    WazzWind BoxBot + 2, col, BoxBot + 5, MaxLen + 3 + col, 10, 4, True
    QPrintRC "  Use:  " + u$ + "-" + d$ + " to select.", BoxBot + 3, col + 3, 15
    QPrintRC u$, BoxBot + 3, col + 11, 14
    QPrintRC d$, BoxBot + 3, col + 13, 14
    
    QPrintRC "Total: " + Str$(DCnt), BoxBot + 4, col + 3, 15
    QPrintRC "Press:   [ESC] to continue.", BoxBot + 3, col + 33, 15
    QPrintRC "        [ENTER] for detail.", BoxBot + 4, col + 33, 15
    QPrintRC "ESC", BoxBot + 3, col + 43, 14
    QPrintRC "ENTER", BoxBot + 4, col + 42, 14
    
    QPrintRC Balance$, Row - 2, col, 112
    QPrintRC Title$, Row - 1, col, 112
    MPaintBox Row, col + MaxLen + 4, Row, col + MaxLen + 5, 8
    'FirstTime = True
    
    'SLEEP
    
    Do
      LOCATE Row, col, 0
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      If Ky$ = Chr$(27) Then
        RestScrn TempScrn()
        Exit Do 'choice = 0
      ElseIf Ky$ = Chr$(13) Then
        RestScrn TempScrn2()
        GoTo ShowTransDetail
      End If
    Loop        'UNTIL EditLocRec& > 0
  Else
    Close UBTran
    Ok = MsgBox%("UB.QSL", "NOCTRANS")
    RestScrn TempScrn()
  End If
  
  RestScrn TempScrn()
  Erase RevText$, Metered, UBSetUpRec, MChoice
  Erase TempScrn, UBTranRec, UBCustRec
  
  Exit Sub
  
ShowTransDetail:
  CursorOff
  TransRecNum& = CVL(Right$(MChoice(Choice).V, 4))
  UBTran = FreeFile
  Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen
  Get UBTran, TransRecNum&, UBTranRec(1)
  Close UBTran
  
  DisplayUBScrn "TRDETAIL"
  
  QPrintRC Num2Date(UBTranRec(1).TransDate), 3, 23, 15
  
  'CONST TranUtilityBill = 1          '   1=Utility bill
  'CONST TranLateCharge = 2           '   2=late charge
  'CONST TranReconnectFee = 3         '   3=reconnect fee
  'CONST TranBillPayment = 4          '   4=Bill Payment
  'CONST TranAppliedDeposit = 5       '   5=Applied Deposit
  'CONST TranPenaltyCharge = 6        '   6=Penalty Charge
  'CONST TranDepositPayment = 7       '   7=Deposit Payment
  'CONST TranDraftPayment = 8         '   8=Draft Payment
  'CONST TranRefundDeposit = 9       '    9=Refund Deposit
  'CONST TranBeginBalance = 10        '  10=Beginning Balance
  'CONST TranUpwardAdjustment = 11    '  11=Bill Adjustments
  'CONST TranDownwardAdjustment = 12  '  12=Bill Adjustments
  
  GoSub GetTransType
  
  QPrintRC FUsing$(Str$(UBTranRec(1).TransAmt), "#####.##"), 4, 25, 15
  
  QPrintRC TType$, 4, 50, 15
  QPrintRC UBTranRec(1).TransDesc, 3, 50, 15
  
  For RevCnt = 1 To NumOfRevs
    QPrintRC RevText$(RevCnt), RevCnt + 6, 8, 15
    QPrintRC FUsing$(Str$(UBTranRec(1).RevAmt(RevCnt)), "#####.##"), RevCnt + 6, 25, 15
    QPrintRC FUsing$(Str$(UBTranRec(1).TaxAmt(RevCnt)), "###.##"), RevCnt + 6, 36, 15
    '(Number$, Image$)
  Next
  
  For Cnt = 1 To 7
    If Metered(Cnt) Then
      QPrintRC FUsing$(Str$(UBTranRec(1).CurRead(Cnt)), "#########"), Cnt + 6, 42, 15
      QPrintRC FUsing$(Str$(UBTranRec(1).PrevRead(Cnt)), "#########"), Cnt + 6, 53, 15
      If Trim$(UBTranRec(1).EstRead(Cnt)) = "" Then
        QPrintRC "N", Cnt + 6, 70, 15
      Else
        QPrintRC "Y", Cnt + 6, 70, 15
      End If
    End If
  Next
  
  WaitForAction
  RestScrn TempScrn2()
  GoTo HistTop
  
GetTransType:
  
  Select Case UBTranRec(1).TransType
  Case TranUtilityBill, TranUtilityBill + 100
    TType$ = "Utility Bill "
  Case TranLateCharge, TranReconnectFee, TranLateCharge + 100, TranReconnectFee + 100
    TType$ = "Penalty, Reconnect Fee"
  Case TranBillPayment, TranBillPayment + 100
    TDesc$ = Trim$(UBTranRec(1).TransDesc)
    If InStr(UBTranRec(1).TransDesc, "PAYMENT") = 0 And Len(TDesc$) > 0 Then
      TType$ = "Utility Payment " + Left$(Trim$(UBTranRec(1).TransDesc), 10)
    Else
      TType$ = "Utility Payment"
    End If
  Case TranPenaltyPayment
    TType$ = "Penalty Payment"
  Case TranPenaltyCharge
    TType$ = "Penalty/Late Fee"
  Case TranAppliedDeposit
    TType$ = "Applied Deposit"
  Case TranDepositPayment, TranDepositPayment + 100
    TDesc$ = Trim$(UBTranRec(1).TransDesc)
    If InStr(UBTranRec(1).TransDesc, "DEPOSIT") = 0 And Len(TDesc$) > 0 Then
      TType$ = "Deposit Payment " + Left$(Trim$(UBTranRec(1).TransDesc), 10)
    Else
      TType$ = "Deposit Payment"
    End If
  Case TranDraftPayment
    TType$ = "Draft Payment"
  Case TranBeginBalance, TranBeginBalance + 100
    TType$ = "Beginning Balance"
  Case 9
    TType$ = "Deposit Refund"
  Case TranUpwardAdjustment
    TType$ = "Upward Adjustment"
  Case TranDownwardAdjustment
    TType$ = "Downward Adjustment"
  Case Else
    TType$ = Str$(UBTranRec(1).TransType) + " ???"
  End Select
  
Return
  
End Sub
Static Sub UBLog(Text$)

  Dim Today As String, TheTime As String
  Dim AmPm As String, SHour As String
    
  Dim THour As Integer, LogFile As Integer
    

'  IF NOT BeenDone THEN
'    BeenDone = True
    Today$ = Date$
    Today$ = Left$(Today$, 2) + Mid$(Today$, 4, 2) + Right$(Today$, 2)
'  END IF

  TheTime$ = Time$
  If Left$(TheTime$, 1) = "0" Then
    THour = Val(Mid$(TheTime$, 2, 1))
  Else
    THour = Val(Mid$(TheTime$, 1, 2))
  End If

  Select Case THour
  Case Is > 11
    THour = THour - 12
    If THour = 0 Then THour = 12
    AmPm$ = "pm"
  Case 1 To 12
    AmPm$ = "am"
  Case 0
    THour = 12
    AmPm$ = "am"
  End Select
  Select Case THour
    Case 1 To 9
      SHour$ = "0" + Trim$(Str$(THour))
    Case Else
      SHour$ = Trim$(Str$(THour))
  End Select
  TheTime$ = SHour$ + ":" + Mid$(TheTime$, 4) + AmPm$
  LogFile = FreeFile
  Open "UBLOG.DAT" For Append Shared As LogFile Len = 255
  Print #LogFile, "UB: "; Today$; " @"; TheTime$
  Print #LogFile, "    "; Text$
  Close #LogFile
  Text$ = "": TheTime$ = ""
End Sub


Static Sub ShowPctComp(ByVal RecNo%, ByVal NumOfRecs%)
 ' RSet PctC(1) = QPStrI$(Int((RecNo / NumOfRecs) * 100))
  'HideCursor
'  QPrintRC PctC(1), 13, 40, Cnf.HiLite
  'ShowCursor
  '  QPrintRC STR$(FRE("")), 25, 1, Cnf.HiLite
End Sub

Static Sub ShowPctCompL(ByVal RecNo&, ByVal NumOfRecs&)
'  RSet PctC(1) = QPStrL$(Int((RecNo& / NumOfRecs&) * 100))
  'HideCursor
'  QPrintRC PctC(1), 13, 40, Cnf.HiLite
  'ShowCursor
  '  QPrintRC STR$(FRE("")), 25, 1, Cnf.HiLite
End Sub

Sub ShowProcessingScrn(RptTitle$)
  TitleRow = 9
  TitleCol = 40 - (Len(RptTitle$) \ 2) + 1
  CursorOff
  BlockClear
  DisplayUBScrn "PRORPT"
  HideCursor
  QPrintRC RptTitle$, TitleRow, TitleCol, 126
  QPrintRC "Processing:    % Completed.", 13, 28, Cnf.HiLite
  ShowCursor
End Sub

Public Function Exist(FileName$)
  Dim FileHandle As Integer
  Dim FileSize As Long
  FileHandle = FreeFile
  Open FileName$ For Binary As FileHandle
  FileSize = LOF(FileHandle)
  Close FileHandle
  If FileSize > 0 Then
    Exist = True
  Else
    Exist = False
    Kill FileName$
  End If
End Function

Public Function FileSize&(FileName As String)
  Dim FileHandle  As Integer
  If Exist(FileName) Then
    FileHandle = FreeFile
    Open FileName For Binary As FileHandle
    FileSize& = LOF(FileHandle)
    Close FileHandle
  Else
    FileSize& = 0
  End If
End Function

Public Static Function Using$(ByVal Fmt As String, ByVal Number As Double)
  Dim TempNumber As String
  Dim FmtNumber As String
  Dim BuckPos As Integer, FmtLen As Integer, TempLen As Integer
  FmtLen = Len(Fmt)
  BuckPos = InStr(Fmt, "$")
  If BuckPos = 1 Then
    Fmt = Right$(Fmt, FmtLen - 1)
  ElseIf BuckPos > 1 Then
    Fmt = Left$(Fmt, BuckPos - 1) + Mid$(Fmt, BuckPos + 1)
  End If
  FmtNumber = Space$(Len(Fmt))
  TempNumber = Format(Number, Fmt)
  TempLen = Len(TempNumber)
  If TempLen >= 2 Then
    If Mid$(TempNumber, (TempLen - 1), 1) = "." Then
      TempNumber = TempNumber + "0"
    End If
  End If
  If Right$(TempNumber, 1) = "." Then
    TempNumber = TempNumber + "00"
  End If
  If BuckPos > 0 Then
    TempNumber = "$" + TempNumber
  End If
  RSet FmtNumber = TempNumber
  Using = FmtNumber
  
'Number = 5: Fmt = "$##,##0.00": Print Right(String(Len(Fmt), " ") & Format(Number, Fmt), Len(Fmt))
End Function

Public Sub ViewPrint(ReportFile As String, Title As String)
   frmViewPrint.ReportName = ReportFile$
   frmViewPrint.Caption = Title
   frmViewPrint.Show 1
End Sub

