Attribute VB_Name = "Module1"
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

