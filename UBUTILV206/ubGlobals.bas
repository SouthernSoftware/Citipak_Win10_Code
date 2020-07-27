Attribute VB_Name = "ubGlobals"
Option Explicit

Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)

Global DebugMode As Boolean
Global TOWNNAME As String
Global MaxLines As Integer
Global PageNo As Integer
Global SaveFlag As Integer
Global UBPath As String
Global LineCnt As Integer
Global FF As String
Global CrLf As String
Global Chr9 As String
Global DoItFlag As Boolean
Global Twiddle As String
Global tmpLastRate As Integer
Global screenW As Long
Public coladj As Double
Global OPERNUM As Integer
'020299
Global NameIndexFile As String
Global BookIndexFile As String
Global TempIndexName As String
Global UBCustFile    As String
Global UBOwnerFile   As String
Global SearchRec As Long
'added rptopt for graphics/text option -PS
Global rptopt As Integer
Global SavePay As Boolean
Global PrnRecp As Boolean
'added for receipt printer default value of yes(1) or no(0)
Global RecpDef As Integer
Global Const ServiceAddressIndexFile = "UBSVCADD.IDX"
Global Const UBBillsFile = "UBBILLS.DAT"
Global Const UBIBillFile = "UBIBILL.DAT"
Global Const UBFinPreRptFile = "UBPREFIN.RPT"
Global Const UBFinBillsFile = "UBFBILLS.DAT"
Global Const RePrintIdxFile = "UBREPRNT.IDX"

Global Const UBHHPathWayFile = "UBHHPATH.DAT"

Global Const MaxRevsCnt = 15

'Transaction Types
Global Const TranUtilityBill = 1          '   1=Utility bill
Global Const TranLateCharge = 2           '   2=late charge      'NOT USED
Global Const TranReconnectFee = 3         '   3=reconnect fee    'NOT USED
Global Const TranBillPayment = 4          '   4=Bill Payment
Global Const TranAppliedDeposit = 5       '   5=Applied Deposit
Global Const TranPenaltyCharge = 6        '   6=Penalty Charge
Global Const TranDepositPayment = 7       '   7=Deposit Payment
Global Const TranDraftPayment = 8         '   8=Draft Payment
Global Const TranRefundDeposit = 9        '   9=Refund Deposit
Global Const TranBeginBalance = 10        '  10=Beginning Balance
Global Const TranUpwardAdjustment = 11    '  11=Bill Adjustments
Global Const TranDownwardAdjustment = 12  '  12=Bill Adjustments
'added this for new over payment adjustment on Aug 11,2003
Global Const TranOverPayAdjustment = 33   '  33=OverPayment Adjustment
Global Const TranDepCreditRemoval = 37    '  37= Deposit Credit Removal - Not Interfaced W/GL
Global Const TranDepPaymentVoid = 39      '  39= Deposit Payment Void  - same gl as deposit refund
Global Const TranMiscPayment = 99         '  99=Misc Payment

'Meter Types
Global Const MtrWaterOnly = 1
Global Const MtrSewerOnly = 2
Global Const MtrCombined = 3
Global Const MtrElectric = 4
Global Const MtrDemand = 5
Global Const MtrGas = 6
Global Const MtrTouchRead = 7
Global Const MtrLightsService = 8
'Global UBSetUpRec(1) As UBSetupRecType

Type PrintBillInfoType
    FrstBill    As Long
    LastBill    As Long
    BillDate    As Integer
    PastDate    As Integer
    PRDate      As Integer
    CRDate      As Integer
    DrftDate    As Integer
    PastDate2   As Integer
    PrnOrder    As String * 25
    MsgLine1    As String * 25
    MsgLine2    As String * 25
    MsgLine3    As String * 25
    MsgLine4    As String * 25
End Type

Type BillMtrType
  CurrRead      As Long
  PrevRead      As Long
  MUsage        As Long
End Type

Type MetersOnBillType
  Mtr(1 To 7) As BillMtrType
End Type

 Type UBLateLetterType
  Head1    As String * 40
  Head2    As String * 40
  Head3    As String * 40
  Head4    As String * 40
  Head5    As String * 40
  Body(1 To 20) As String * 75
End Type

Type NoticeInfoType
  FromBC        As Integer
  ThruBC        As Integer
  NoticeDate    As Integer         '1
  PayByDate     As Integer         '2
 'FromDate      AS INTEGER         '3
 'TODate        AS INTEGER         '4
  MinBalance    As Double          '5
  BalanceType   As Integer         '6
  PrnOrder      As Integer         '7
  UseAFlag      As Integer         '8
  MsgLine1      As String * 25
  MsgLine2      As String * 25
  MsgLine3      As String * 25
  MsgLine4      As String * 25
  PrnCnt        As Integer
End Type

Type PenaltyInfoType
  PenDate    As Integer
  PenDesc    As String * 21
  RevSource  As Integer
  ChargeOn   As String * 17
'032299 Changed to a double variable
  PctCharge  As Double
  'PctCharge  AS INTEGER
  AmtCharge  As Double
  GreatLess  As String * 1
  MinBalance As Double
  CycFirst   As Integer
  CycLast    As Integer
  BookFirst  As Integer
  BookLast   As Integer
  PenCnt     As Integer
End Type

Type PSAZipIndexType
  ZIPCODE  As String * 10
  SName    As String * 10
  RecNum   As Long
  Pad      As String * 10
End Type

Type MOWZipIndexType
  ZIPCODE  As String * 10
  RecNum   As Long
  FillPad As String * 4
End Type

Type UBPostalIndexType
  ZIPCODE  As String * 10
  Route    As String * 4
  RecNum   As Long
End Type

Type UBServiceAddressIndexType
  ServiceAddress  As String * 14
  RecNum   As Long
End Type

Type UBSequenceIndexType
  SeqNumber As Long
  RecNum    As Long
  Fill      As String * 10      'This is to fill this to a 16 byte boundary
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
  BOOK       As String * 2
  SEQNUMB    As String * 6
  RecNum     As Long
  Pad        As String * 4
End Type

Type UBLocaReIndexRecTypeVB
  BookSEQNUMB    As String * 8
  RecNum         As Long
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
    RevName    As String * 20
    RATECODE   As String * 4
    RevMtrType As String * 1
End Type
'---------------------------------------------------------

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

'Type LocMeterType
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
'    MtrIDNO   As String * 11
'    MtrLat    As Double
'    MtrLng    As Double
'End Type

Type MonthlyPayType
    AMTOWED      As Double
    TotAmtPD     As Double
    PayAmt       As Double
    RevSource    As Integer
End Type
'-------------------------------------------------------------
Type NewUBCustRecType
    BOOK          As String * 2
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
    EPPAMT        As Double

    USEDRAFT      As String * 1
    BANKNAME      As String * 34
    BANKLOC       As String * 30
    TRANSIT       As String * 9
    BANKACCT      As String * 20
    BILLCMNT      As String * 25
    PAYCMNT       As String * 25
    PUMPCODE      As String * 4
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
    'EPPBalances(1 TO 15) AS DOUBLE
    PrevRevAmts(1 To 15) As Double
    DepositAmt    As Double
    DelFlag       As Integer
    PreNoteFlag   As Integer
    WOLastTrans   As Long            'work order last trans pointer
    EstFlag       As String * 1
    MessageRec    As Long            ' Points to Message Record
    OldRec        As Long
    EPPLastTran   As Long
    FillPad       As String * 7
End Type

'Type NewUBCustRecType
'    Book          As String * 2
'    SEQNUMB       As String * 6
'    Status        As String * 1
'    OPENDATE      As Integer
'    SEARCH        As String * 10
'    CustName      As String * 35
'    ADDR1         As String * 35
'    ADDR2         As String * 35
'    SERVADDR      As String * 35
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
''added GroupCoderec 2/1/05 for pointer to bookcode
'    GroupCodeRec  As Integer
'    Filler1       As String * 5
'   ' Filler1       As String * 7
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
'    serv(1 To 15)      As ServicesType
'    FlatRates(1 To 4)  As FlatRateType
''Page 4
'    Monthly(1 To 2)    As MonthlyPayType
'    MFEE1         As Double
'    MFEE2         As Double
'    LocMeters(1 To 7)  As LocMeterType
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
'    DPCode        As String * 2
'    FillPad       As String * 112
'    ChkByte       As String * 1
'End Type

Type GroupCodeRecType
    Deleted       As Integer
    GroupCODE     As String * 2
    GroupCodeName As String * 30
    xtrastuff     As String * 30
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

Type WorkOrderDefType
    Deleted           As Boolean
    WOType            As String * 20
    OrdersText        As WrkOrdTextType
    RepliesText       As WrkOrdTextType
    Xtra              As String * 20 'just in case
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

Type SetUpAcctType
   RevName    As String * 15
   DebitAcct  As String * 14
   CreditAcct As String * 14
End Type

Type RevSetUpType
    RevName As String * 15
    UseDep   As String * 1
    USERATE  As String * 1
    TAXRATE  As Single
    UseMtr   As String * 1
    DistOr   As Integer
    ProRate  As String * 1
End Type

Type UBSetupRecType
    UTILNAME        As String * 35
    DEFCITY         As String * 18
    DEFSTATE        As String * 2
    ZIPCODE         As String * 10
    PreByBook       As String * 1
    RecpPort        As String * 1
    RECPDEFT        As String * 1
    ESTREAD         As String * 1
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

'Added billsetup 9/29/03
Type UBBillSetupType
    Bill          As Integer
    LateNotice    As Integer
    PostBar       As String * 1
    AcctBar       As String * 1
    Permit        As String * 1
    ChargeOn      As String * 17
    PctCharge     As Double
    AmtCharge     As Double
    GreatLess     As String * 1
    MinBalance    As Double
    BL1Head1     As String * 30
    BL1Head2     As String * 40
    BL1Head3     As String * 40
    BL1Permit1   As String * 16
    BL1Permit2   As String * 16
    BL1Permit3   As String * 16
    BL1Permit4   As String * 16
    BL1Permit5   As String * 16
    BL1Opt1      As String * 20
    BL1Opt2      As String * 20
    BL1Opt3      As String * 22
    BL1Opt4      As String * 22
    BL1Opt5      As String * 22
    BL1Opt6      As String * 22
    BL1Opt7      As String * 22
    BL1Opt8      As String * 22
    BL1Opt9      As String * 22
    BL1Opt10     As String * 22
    AmtChge2     As Double
    reserved     As String * 22
End Type
'Added 6/29/04
Type UBBillLetterType
    IncLogoFlag  As Integer  ' 0 for not to print, 1 to print
    MtrNumFlag   As Integer   '1 for Meter Serial, 2 for Meter ID
    BL1Head1     As String * 30
    BL1Head2     As String * 40
    BL1Head3     As String * 40
    MsgOpt1      As String * 40  'set to print on top section
    MsgOpt2      As String * 40  'prints top section
    MsgPgph1     As String * 125 'prints under service charges
    MsgPgph2     As String * 125
    MsgPgph3     As String * 125
    MsgPgph4     As String * 125
    MsgPgph5     As String * 125
    MsgOpt3      As String * 50  'prints top of stub
    MsgOpt4      As String * 30  'prints mid of stub
    MsgOpt5      As String * 30  'prints mid right lower side of stub
    Padding      As String * 95
End Type

Type UBHHPathRecType
    PathWay As String * 48
End Type

'Type UBSensusRecType
'    PathWay As String * 20
'End Type

'Type UBLogiconRecType
'    PathWay As String * 20
'End Type

'Type UBPC3000RecType
'    PathWay As String * 20
'End Type

Type DistArrayType
   DistOrder As Integer
   DistCnt   As Integer
End Type

'Note:  if transaction is an adjustment then
'       CurRead field will contain the adjust amount
Type UBTransRecType
   TransDate              As Integer      '
   TransType              As Integer      '
   TransDesc              As String * 21  'may change
   Transamt               As Double       'total revenue amount
   RevAmt(1 To 15)        As Double       'Revenue amounts
   TaxAmt(1 To 15)        As Single       'Tax Amounts
'01-20-97 Added meter types field to hold meter type at time of transaction
   MtrTypes(1 To 7)       As Integer
'*******************
   CurRead(1 To 7)        As Long         'Last/Current meter readings
   PrevRead(1 To 7)       As Long         'Previous readings
   ESTREAD(1 To 7)        As String * 1   'Y/N Flags for meter est's
   BillNumber             As Long         'Number on the bill that Printed
   ReadDate               As Integer
   BillDate               As Integer
   PastDueDate            As Integer
   DraftDate              As Integer      '
'111398
   ProRatePCT             As Integer
   ChkByte                As String * 1   'Added check byte
   EPPFlag                As String * 1   'Equal Payment Flag
   CustStatus             As String * 1   'Customer Status at Time of Transaction
'020199
   EPPTrans               As Long         'Pointer to Equal Pay trans
   PenAtBill              As Single       'Used to flag IRR Meter (Sunset)
'****************
   PayTypeCode            As Integer      'Payment Type:  1=Cash, 2=Check, 3=Cash/Check, 4=Charge
   OperatorNumber         As Integer      '
   CustAcctNo             As Long         'Pointer to RecNo in ubcust.dat
   PrevTrans              As Long
   VoidFlag               As Integer       'Changed for wadesboro
   FromCMFlag             As Integer
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

Type OldMessLineType
  Line As String * 59
  LineDate As String * 10
End Type

Type MessLineType
  Msg As String * 69
End Type

Type UBMessRecType
  MessLine(1 To 15) As MessLineType
  CustRec As Long
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
    BOOK             As Integer
    CustCnt          As Long
    Consump(1 To 15) As Double
    RevAmt(1 To 15)  As Double
    TaxAmt(1 To 15)  As Double
End Type

Type PumpConsumpType
    PUMPCODE         As String * 4
    CustCnt          As Long
    Consump          As Double
End Type

Type UBOwnerRecType
    OwnLName  As String * 20
    OwnFName  As String * 15
    ADDR1     As String * 35
    ADDR2     As String * 35
    CITY      As String * 18
    STATE     As String * 2
    ZIPCODE   As String * 10
    HPHONE    As String * 14
    WPHONE    As String * 14
    ChkByte   As String * 1
End Type

Type CMTransRecType
    TransDate    As Integer
    TransAmount  As Double
    TransCash    As Double
    TransCheck   As Double
    TransAmtOwed As Double
    TransDesc    As String * 25
    TransSource  As Integer           '1-Misc 24-Util 27-UtilDep 31-Tax 131-Newtax 41-License 141-NewBL 51-decal
    ''''''''''''''''''''''''''''''''''201-void Misc 224-void util 227-void dep 241-void lic 231-void tax
    ''''''''''''''''''''''''''''''''''251-void Decal
    TransName    As String * 25
    TransAcctNum As Long               'Holds Master Acct Record Number in Mod
    TransDetNum  As Long               'Holds Record Number of Transaction Det
    TransRevAmt(1 To 15) As Double
    TransOperNum As Long
    Trans2GL     As String * 1
    TransTender  As Integer     'Type: 1=Cash, 2=Check, 3=Cash/Check, 4=Charge
'added charge 4 above and transvoid for new void payment procedure PS 4/14/04
    TransVoidNum As Long        'Voided trans link to record voided or void trans
    ChkByte      As String * 1
    TransPad     As String * 18
End Type
Type MiscCodeRecType
    MiscCode As String * 7
    Description As String * 25
    GlAcctNumb As String * 14
    InActiveFlag As String * 1
    NotUsed As String * 17
End Type
Type ARNewCatCodeRecType
    CatCode    As String * 5    'Not Used in Version 8.5 work2 directory
    CodeType   As String * 1    ' F=Flat M=Multiplier S=Step
    CODEDESC   As String * 35
    Fee        As Single
    BaseAmt1   As Single
    Recpt1     As Double
    Percent1   As Single
    Maximum1   As Double
    BaseAmt2   As Single
    Recpt2     As Double
    Percent2   As Single
    Maximum2   As Double
    BaseAmt3   As Single
    Recpt3     As Double
    Percent3   As Single
    Maximum3   As Double
    BaseAmt4   As Single
    Recpt4     As Double
    Percent4   As Single
    Maximum4   As Double
    BaseAmt5   As Single
    Recpt5     As Double
    Percent5   As Single
    Maximum5   As Double
    REVGLNUM   As Long
    CASHACCT   As Long
    ARGLACCT   As Long

    BaseAmt6   As Single
    Recpt6     As Double
    Percent6   As Single
    Maximum6   As Double
    RateStep   As Long
    Extra      As String * 36
End Type

