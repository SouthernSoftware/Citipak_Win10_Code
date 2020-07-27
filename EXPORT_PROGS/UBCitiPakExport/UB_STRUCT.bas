Attribute VB_Name = "UB_STRUCT"
Option Explicit
Public Const UBGrpCde = "UBGrpCde.DAT"
Public Const UBSetup = "UBSetup.DAT"
Public Const UBDraftRec = "UBSDRAFT.DAT"
Public Const UBRateCodes = "UBRate.DAT"
Public Const UBLateLetter = "UBLatlet.DAT"
Public Const UBLaserBill = "UBBilLtr.DAT"
Public Const UBWoTrans = "UBWRKORD.DAT"
Public Const UBData = "UBdata\"
Public Const UBOwner = "UBOWNER.DAT"
Public Const UBMessage = "UBMESAGE.DAT"
Public Const UBCust = "UBCUST.DAT"
Public Const UBTransRec = "UBTRANS.DAT"

Global DebugMode As Boolean
Global TownName As String
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
Global WDflag As Boolean  'for wadesboro log
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
'added for receipt printer default for Control Codes Yes(1) or No(0)
Global CntrlDef As Integer
Global PrnVali As Boolean
'added for validation default  yes(1) no(0)
Global ValiDef As Integer
Global BnkAcctNum As String

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
Global Const MtrIrrigation = 9
'Global UBSetUpRec(1) As UBSetupRecType

Type PrintBillInfoType
    FrstBill    As Long
    LastBill    As Long
    BillDate    As Integer
    DueDate     As Integer
    PastDate    As Integer
    PRDate      As Integer
    CRDate      As Integer
    DrftDate    As Integer
    PastDate2   As Integer
    PrnOrder    As String * 25
    MsgLine1    As String * 75
    MsgLine2    As String * 75
    MsgLine3    As String * 75
    MsgLine4    As String * 75
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
Type UBZipLocationIndexType
  ZIPLocat As String * 18
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
    Ratecode As String * 4
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
    Ratecode   As String * 4
    RevMtrType As String * 1
End Type
'------------------------------------------------------------------
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
'    MTRType   As String * 1
'    MtrUnit   As String * 1
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

'-----------------------------------------------------------------------
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
    ProratePCT    As Integer
    HHMSG1        As String * 20
    HHMSG2        As String * 20
    HHMSG3        As String * 20
'Page 3
    Serv(1 To 15)      As ServicesType   'Nochange
    FlatRates(1 To 4)  As FlatRateType   'Nochange
'Page 4
    Monthly(1 To 2)    As MonthlyPayType 'Nochange
    MFEE1         As Double
    MFEE2         As Double
    LocMeters(1 To 7)  As LocMeterType   'change
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
'    BOOK          As String * 2
'    SEQNUMB       As String * 6
'    Status        As String * 1
'    OPENDATE      As Integer
'    SEARCH        As String * 10
'    CUSTNAME      As String * 35
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
'    SEQ           As Long
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
'    BANKNAME      As String * 34
'    BANKLOC       As String * 30
'    TRANSIT       As String * 9
'    BANKACCT      As String * 20
'    BILLCMNT      As String * 25
'    PAYCMNT       As String * 25
'    PUMPCODE      As String * 4
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
'    LocMeters(1 To 7)  As LocMeterType
''END OF Quick Screen Form
'    CustPIN       As Long
'    LastTrans     As Long
'    CurrBalance   As Double
'    PrevBalance   As Double
'    CurrRevAmts(1 To 15) As Double   'includes the tax amount
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

Type GroupCodeRptType
    RecordNum   As Integer
    GroupCode   As String * 2
End Type
Type RateCodeRptType  'used in consump rpt for temp file for rates selected
    RecordNum   As Integer
    Ratecode    As String * 4
End Type

Type GroupCodeRecType
    Deleted       As Integer
    GroupCode     As String * 2
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
    Prorate  As String * 1
End Type

Type UBSetupRecType
    UTILNAME        As String * 35
    DEFCITY         As String * 18
    DEFSTATE        As String * 2
    ZIPCODE         As String * 10
    PreByBook       As String * 1
    'RecpPort        As String * 1 ' change to LockBoxDef on 1/25/05
    LockBoxDef      As String * 1  ' 6 for 6digit acct, 8 for 8digit acct file type struc
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
    RtePrint     As Integer
    Reserved     As String * 20
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
'added 8/25/06 for Middletown 21Line blank paper bill format
'this type structure will allow generic use bill with defaults
Type UBBill21LineType
    MsgORdates   As String * 1  'M for message line 1 or D for service dates
    TxtLine1     As String * 40   'lines 1-11 print top right side of form
    TxtLine2     As String * 40
    TxtLine3     As String * 40
    TxtLine4     As String * 40
    TxtLine5     As String * 40
    TxtLine6     As String * 40
    TxtLine7     As String * 40
    TxtLine8     As String * 40
    TxtLine9     As String * 40
    TxtLine11    As String * 40
    TxtBLine1    As String * 130  'prints bottom of form
    TxtBLine2    As String * 130  'prints bottom of form

End Type
'added 8/25/06 for Middletown 21Line blank paper late notice format
'this type structure will allow generic use notice with defaults
Type UBLateNotice21LineType
    TxtLineH1    As String * 40  'msg on left of notice top
    TxtLineH2    As String * 40
    TxtLine1     As String * 40   'lines 1-11 print top right side of form
    TxtLine2     As String * 40
    TxtLine3     As String * 40
    TxtLine4     As String * 40
    TxtLine5     As String * 40
    TxtLine6     As String * 40
    TxtLine7     As String * 40
    TxtLine8     As String * 40
    TxtLine9     As String * 40
    TxtLine10    As String * 40
    TxtLine11    As String * 40
    TxtBLine1    As String * 130  'prints bottom of form
    TxtBLine2    As String * 130  'prints bottom of form

End Type

'added 7/30/04 for Tray Labels for Postal Zip sorting Utility Bills
'Type PostTrayLabelType
'    '
'End Type
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
    BOOK       As String * 2
'mod for Wrongsville beech!
    MtrSize    As String * 2
    AvgUse     As String * 8
    Future     As String * 19
'    Future     As String * 29
    Recend     As String * 1               'Must be 'X'
    CrLf       As String * 2
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
    BOOK       As String * 2
    Future     As String * 29
    Recend     As String * 1               'Must be 'X'
    CrLf       As String * 2
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
    BOOK          As Integer
    CurRead       As Double
    PastRead      As Double
    ReadDate      As Integer
    ReadTime      As String * 5
    Note1         As String * 20
    Note2         As String * 20
    Note3         As String * 20
    NoteStatus    As String * 1    'T=Temp Note  P=Perm Not'
'    MtrIDMST      As String * 0
'    MtrIDNO       As String * 1
End Type

Type UBDGRecType
    PathWay As String * 20
End Type

Type UBDGHHRecType        ' File Layout for Sending Out Records
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
    CrLf As String * 2
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
    CrLf As String * 2
End Type

Type ItronBType                         'Customer Record Two
    RecordCode As String * 1
    CUSTNAME As String * 32
    CUSTADDR As String * 32
    Filler As String * 13
    CrLf As String * 2
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
    CrLf As String * 2
End Type

Type ItronZType
    RecordCode As String * 1
    Route As String * 8
    NumberAccts As String * 4
    NumberMeters As String * 4
    Filler As String * 61
    CrLf As String * 2
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

''Note:  if transaction is an adjustment then
''       CurRead field will contain the adjust amount
'Type UBTransRecType
'   TransDate              As Integer      '
'   TransType              As Integer      '
'   TransDesc              As String * 21  'may change
'   TransAmt               As Double       'total revenue amount
'   RevAmt(1 To 15)        As Double       'Revenue amounts
'   TaxAmt(1 To 15)        As Single       'Tax Amounts
''01-20-97 Added meter types field to hold meter type at time of transaction
'   MtrTypes(1 To 7)       As Integer
''*******************
'   CurRead(1 To 7)        As Long         'Last/Current meter readings
'   PrevRead(1 To 7)       As Long         'Previous readings
'   ESTREAD(1 To 7)        As String * 1   'Y/N Flags for meter est's
'   BillNumber             As Long         'Number on the bill that Printed
'   ReadDate               As Integer
'   BillDate               As Integer
'   PastDueDate            As Integer
'   DraftDate              As Integer      '
''111398
'   ProRatePCT             As Integer
'   ChkByte                As String * 1   'Added check byte
'   EPPFlag                As String * 1   'Equal Payment Flag
'   CustStatus             As String * 1   'Customer Status at Time of Transaction
''020199
'   EPPTrans               As Long         'Pointer to Equal Pay trans
'   PenAtBill              As Single       'Used to flag IRR Meter (Sunset)
''****************
'   PayTypeCode            As Integer      'Payment Type:  1=Cash, 2=Check, 3=Cash/Check, 4=Charge
'   OperatorNumber         As Integer      '
'   CustAcctNo             As Long         'Pointer to RecNo in ubcust.dat
'   PrevTrans              As Long
'   VoidFlag               As Integer      'Changed for wadesboro
'   FromCMFlag             As Integer
'   ActiveFlag             As Integer      'Valid transaction flag
'   RunBalance             As Double
'   CheckAmount            As Double
'   CashAmount             As Double
'   BillMsg                As String * 20
'   ApplyDepFlag           As String * 1
'   Posted2GL              As String * 1
'   PrevDate               As Integer
'   PenalFlag              As String * 1
'   TaxExempt              As String * 1
'   NONProfit              As String * 1
'End Type
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
   ProratePCT             As Integer
   Filler1                As String * 1
   EPPFlag                As String * 1   'Equal Payment Flag
   CustStatus             As String * 1   'Customer Status at Time of Transaction
'020199
   'PenAtBill              AS DOUBLE
   EPPTrans               As Long         'Pointer to Equal Pay trans
   Filler2                As String * 4
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
    Ratecode As String * 4
    RATEDESC As String * 29
    ChkByte  As String * 1
    MINAMT   As Double
    MINUNITS As Long
    MaxAmt   As Double
    TblBreaks(1 To 10) As TblBreakRecType
End Type

Type oUBRateTblRecType
    Ratecode As String * 4
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
Type BookGroupType
    BOOK             As String * 2
    CustCnt          As Long
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

Type MtrNotesType
  Note    As String * 20
End Type

Type UBHuskyHHRecType                'File Layout for Sending Out Records
  CUSTNAME      As String * 20
  ServAddress   As String * 20
  UCode1        As String * 2
  UCode2        As String * 2
  MeterID       As String * 8
  LowRead       As Long
  HighRead      As Long
  Account       As String * 8
  ReadFlag      As String * 1        'Y/N
  MeterType     As String * 1
  BOOK          As Integer
  CurRead       As Long
  PastRead      As Long
  ReadDate      As Integer
  'ReadTime      AS STRING * 5
  NoteStatus    As String * 1    'T=Temp Note  P=Perm Note
  Notes(1 To 3)  As MtrNotesType
End Type
'Added the following types when did the cycle count sum report, PS 5-7-03
Type MtrDateSortType
  MtrDate As Integer
  RecNum   As Integer
End Type

Type CycleType
   CustCnt As Long
   PendCnt As Long
End Type
'062603 Added draft info
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

  Type DraftRptType
    TRANSIT  As String * 9
    BANKNAME As String * 14
    'BANKNAME AS STRING * 34
    CustAcct As String * 5
    CUSTNAME As String * 23
    AcctType As String * 1
    BillAmt  As String * 9
    BANKACCT As String * 20
  End Type
Type UBConsumpRptType   'Report for Stat Consumption Top Ten
    CustAcct      As Long
    ConsumpAmt    As Double
    CUSTNAME      As String * 20
    SvcAddr       As String * 20
    AvgAmt        As Double
End Type

  Type BDRptType
    BANKNAME  As String * 14
    CustRec   As Integer
    TransRec  As Long
  End Type

  Type BankTotalsType
    BANKNAME  As String * 14
    Amount    As Double
  End Type

Type PaidOwedType
    AMTOWE1  As Double
    AMTPD1   As Double
End Type

Type UBPaymentRecType
    OPERNUM         As Integer
    PAYDATE         As Integer
    CustAcct        As Long
    CUSTNAME        As String * 24
    CUSTADDR        As String * 24
    CUSTCMNT        As String * 32
'052598 Added tax exempt flag changed cust comment len to 32
    TaxExempt       As String * 1
    AMTOWED         As Double
    TENDERTY        As String * 12
    CASHAMT         As Double
    CHKAMT          As Double
    AMTRECD         As Double
    CHANGE          As Double
    DESC            As String * 19
    PaidOwed(1 To 15)   As PaidOwedType
    TOTOWED         As Double
    AMTPAID         As Double
'112801 Added cust status changed desc to 19
    Status          As String * 1
End Type

Type PayListType
  CustRec   As Long
  Listrec   As Long
End Type

Type BookTotalType
  Count   As Long
  Cash    As Double
  Check   As Double
  Charge  As Double
  CHANGE  As Double
End Type

Type CMOperRecType
    OperatorNumber As Integer
    OperatorName As String * 30
    OperatorPassword As String * 8
    NotUsed As String * 24
End Type

Type MiscCodeRecType
    MiscCode As String * 7
    Description As String * 25
    GlAcctNumb As String * 14
    NotUsed As String * 18
End Type

Type RMReceiptRecType
    RecName As String * 25
    RecAddress As String * 25
    RecDesc As String * 25
    RecAmtOwed As Double
    RecPayType As Integer
    RecCashAmt As Double
    RecCheckAmt As Double
    RecChangeDue As Double
    RecDate As String * 10
    RecOperator As String * 2
    RecptNumber As Single
    RecBlank As String * 1
    RecBalance As Single
End Type

Type CMTransRecType
    TransDate    As Integer
    TransAmount  As Double
    TransCash    As Double
    TransCheck   As Double
    TransAmtOwed As Double
    TransDesc    As String * 25
    TransSource  As Integer            '1-Misc 2-Util 3-Tax 4-License 5-decal
    TransName    As String * 25
    TransAcctNum As Long               'Holds Master Acct Record Number in Mod
    TransDetNum  As Long               'Holds Record Number of Transaction Det
    TransRevAmt(1 To 15) As Double
    TransOperNum As Long
    Trans2GL      As String * 1
    TransPad     As String * 25
End Type

Type CMConfigType
    TownName As String * 30
    CASHACCT As String * 14
    LPTPORT  As String * 1
    PrnDefYN As String * 1
    ENDMSG   As String * 30
End Type
  Type FLen2
    V As String * 64
  End Type

'This is for new local receipt setup file stored on each computer on
'drive c:\
'added ctlDefYN on 7/27/04
'added RValidate 5/2/05 for validation flag
Type ReceiptPRNType
  RcpPort   As String * 40
  PrnDefYN  As Integer
  CtlDefYN  As Integer
  PaymDate  As Integer    'For Changing Default Date During Daily Entry
  RValidate As Integer
  ZExtra    As String * 16
End Type
Type CMBankAcctRecType
    COMPACCT As String * 20
End Type

Type UBIntermecHHRecType         '
' First two fields are required by CE File I/O dll.
  CEVariant       As String * 2  'MUST BE  CHR$(8) + CHR$(0)
  CEStrLen        As String * 2  'MUST BE CHR$(165) + CHR$(0)
'**************************************************************
  CUSTNAME      As String * 20
  ServAddress   As String * 20
  MeterID       As String * 8
  LowRead       As String * 9
  HighRead      As String * 9
  Account       As String * 8
  ReadFlag      As String * 1         'Y/N
  MeterType     As String * 1
  BOOK          As String * 2
  CurRead       As String * 9
  PastRead      As String * 9
  ReadDate      As String * 8   ' fmt mmddyyyy
  Note1         As String * 20
  Note2         As String * 20
  Note3         As String * 20
  NoteStatus    As String * 1  'T=Temp  P=Perm
End Type

Type SReadType
   BOOK        As String * 2
   SEQ         As String * 7
   CUSTNAME    As String * 30
   SERVADDR    As String * 20
   CurrRead    As String * 10
   ReadDate    As String * 10
   CrLf        As String * 2
End Type

Type SchlumHHType
  'ne=No Equivalent, r=Required, o=Optional
  'r=right justified l=left justified z=zero filled s=space filled
  Route        As String * 10  'r       uBook-Sequence   LS
  WalkSeq      As String * 4   'r
  PageNum      As String * 4   'r  18
  ReadSeq      As String * 2   'r
  HHID         As String * 6   'o  26
  ReadDir      As String * 1   'o ne
  NumDial      As String * 1   'o ne
  IDExpected   As String * 13  'r meter serial#    43
  IDCaptured   As String * 13  'o
  IDOverride   As String * 13  'o                  67
  Decimals     As String * 1   'o ne               68
  MtrRead      As String * 10  'o                  78
  ReadOVRide   As String * 10  'o ne               88
  HighLimit    As String * 10  'o                  98
  LowLimit     As String * 10  'o                  108
  Date2Read    As String * 6   'o
  Date2Exp     As String * 6   'o
  NoteCodes    As String * 8   'o ne
  LocatCode    As String * 2   'o
  MtrRCode     As String * 2   'o
  RecType      As String * 2   'o   'reading status
  RecStatus    As String * 1   'o ne internal debugging
  ReadDate     As String * 6   'o must have  Actual read date
  ReadTime     As String * 6   'o ne
  ReadType     As String * 1   'o Reading class
  NetNumb      As String * 2   'o ne
  ReadAtmpt    As String * 1   'o ne
  UserChar     As String * 7   '????? no code in schlum manual
  HHManufac    As String * 1   'o ne
  ActStatus    As String * 1   'o Account Status  A/I
  MTRType      As String * 1   'o Type for automatic read
  ReadFailCode As String * 1   'o Reading Failure code
  PrevRead     As String * 10  'o Previous Mtr reading
  PrevDate     As String * 6   'o Previous read date
  HHDisp1      As String * 24  'r Service Addr
  HHDisp2      As String * 24  'r Remote Location (Book-Sequence)
  HHDisp3      As String * 24  'r More Location info
  HHDisp4      As String * 24  'r Other Info (Maybe name)
  Notes1       As String * 24  'o Notes Fields
  Notes2       As String * 24  'o Notes Fields
  Notes3       As String * 24  'o Notes Fields
  Notes4       As String * 24  'o Notes Fields
  Notes5       As String * 24  'o Notes Fields
  Notes6       As String * 24  'o Notes Fields
  Notes7       As String * 24  'o Notes Fields
  Notes8       As String * 24  'o Notes Fields
  OpCode       As String * 1   'r Reserved
  UBAcctNo     As String * 6   'Sosoft Reserved Account Number
  MtrSlot      As String * 1   'Sosoft Reserved Meter Slot
  UtilFld      As String * 33  'o For us to use
  CrLf         As String * 2   'r CrLf terminator
End Type

' These Sensus Layout Files are Spec'd Out Exactly to Gilbert SC

Type UBGilSensusReadRecType         ' File Layout for Sending Out Records
    CustLastName   As String * 25
    CustFirstName  As String * 25
    MeterID        As String * 8
    Account        As String * 8
    LowRead        As String * 8
    HighRead       As String * 8
    SensusType     As String * 1        ' B=Touch Read : M=Manual
    PastRead       As String * 8
    CurRead        As String * 8
    ServAddress    As String * 20
    LocationNumber As String * 10
    Message        As String * 30
'040604
    MtrIDMST       As String * 10
    MtrIDNO        As String * 1
    MtrLat         As String * 11
    MtrLng         As String * 11
End Type

Type UBGilSensusGetReadRecType      ' File Layout For Retreiving Records
    CustLastName   As String * 25
    CustFirstName  As String * 25
    MeterID        As String * 8
    Account        As String * 8
    LowRead        As String * 8
    HighRead       As String * 8
    SensusType     As String * 1        ' B=Touch Read : M=Manual
    PastRead       As String * 8
    CurRead        As String * 8
    ServAddress    As String * 20
    LocationNumber As String * 10
    fil1           As String * 1
    ReadDate       As String * 6
    Message        As String * 19   'was 26
    DateRead       As String * 4
    MtrLat         As String * 11
    MtrLng         As String * 11
End Type

'10/18/04 Old ESensus type structure.
Type UBOESensusReadRecType         ' File Layout for Sending Out Records
    CustLastName As String * 25
    CustFirstName As String * 25
    MeterID As String * 8
    Account As String * 8
    LowRead As String * 8
    HighRead As String * 8
    SensusType As String * 1        ' B=Touch Read : M=Manual
    PastRead As String * 8
    CurRead As String * 8
    ServAddress As String * 20
    LocationNumber As String * 10
    Message As String * 30
End Type

Type UBOESensusGetReadRecType      ' File Layout For Retreiving Records
    CustLastName As String * 25
    CustFirstName As String * 25
    MeterID As String * 8
    Account As String * 8
    LowRead As String * 8
    HighRead As String * 8
    SensusType As String * 1        ' B=Touch Read : M=Manual
    PastRead As String * 8
    CurRead As String * 8
    ServAddress As String * 20
    LocationNumber As String * 10
    fil1           As String * 1
    ReadDate       As String * 6
    Message As String * 19   'was 26
    DateRead As String * 4
End Type

'This is for temporary Tax bill for Spruce Pine and Fairmont
Type TxBillDefaultsType
    TxtHead1 As String * 50
    TxtHead2 As String * 50
    txtOpt1 As String * 40
    TxtOpt2 As String * 40
    TxtOpt3 As String * 40
    TxtOpt4 As String * 40
    txtPgph0 As String * 125
    txtPgph1 As String * 125
    txtPgph2 As String * 125
    txtPgph3 As String * 125
    txtPgph4 As String * 125
    txtPgph5 As String * 125
    txtPgph6 As String * 125
    txtPgph7 As String * 125
    TxtOpt5 As String * 75
    txtHead4 As String * 40
    txtHead5 As String * 40
    txtHead6 As String * 40
    TxtOpt6 As String * 45
    TxtOpt7 As String * 75
    dologo As Integer  '0 for no 1 for yes
End Type

Type BillOutServType
   ServText               As String * 20
   ServAmt                As String * 10
End Type

Type BillOutRecType   '
   AcctNo                 As String * 8
   LocationNum            As String * 9        'format ##-######
   CUSTNAME               As String * 35
   ADDR1                  As String * 35
   ADDR2                  As String * 35
   SERVADDR               As String * 35
   CITY                   As String * 18
   STATE                  As String * 2
   ZIPCODE                As String * 10
   BillType               As String * 1        'N=Normal F=Final
   DepAppAmt              As String * 10       'if final applied deposit amt.
   PrevDue                As String * 15       'total revenue amount
   CurrDue                As String * 15       'total revenue amount
   TotalDue               As String * 15       'total revenue amount
   MTRType                As String * 1        'w=water s=sewer c=combined
   MTRUnit                As String * 1
   CurrDate               As String * 10
   PrevDate               As String * 10
   ServDays               As String * 4
   CurrRead               As String * 10       'Last/Current meter readings
   PrevRead               As String * 10       'Previous readings
   Consump                As String * 10
   IServDays              As String * 4
   ICurrRead              As String * 10       'Last/Current meter readings
   IPrevRead              As String * 10       'Previous readings
   IConsump               As String * 10
   ServInfo(1 To 15)      As BillOutServType
   BillDate               As String * 10
   PastDueDate            As String * 10
   DraftDate              As String * 10
   MsgLine1               As String * 22     'was 25
   MsgLine2               As String * 22     'was 25
   MsgLine3               As String * 22     'was 25
   MsgLine4               As String * 22     'was 25
   MtrNumb                As String * 12     'Added Meter Number
   CrLf                   As String * 2
End Type
Type BillOut2MeterType
   MTRType                As String * 1        'w=water s=sewer c=combined
   MTRUnit                As String * 1
   CurrRead               As String * 10       'Last/Current meter readings
   PrevRead               As String * 10       'Previous readings
   Consump                As String * 10
   MtrIDNum               As String * 12
End Type
Type BillOutRec2Type   '
   AcctNo                 As String * 8
   LocationNum            As String * 9        'format ##-######
   Cycle                  As String * 2
   CUSTNAME               As String * 35
   ADDR1                  As String * 35
   ADDR2                  As String * 35
   SERVADDR               As String * 35
   CITY                   As String * 18
   STATE                  As String * 2
   ZIPCODE                As String * 10
   BillType               As String * 1        'N=Normal F=Final
   DepAppAmt              As String * 10       'if final applied deposit amt.
   PrevDue                As String * 15       'total revenue amount
   CurrDue                As String * 15       'total revenue amount
   TotalDue               As String * 15       'total revenue amount
   CurrDate               As String * 10
   PrevDate               As String * 10
   ServDays               As String * 4
   MtrInfo(1 To 7)        As BillOut2MeterType
   ServInfo(1 To 15)      As BillOutServType
   BillDate               As String * 10
   PastDueDate            As String * 10
   DraftDate              As String * 10
   MsgLine1               As String * 22     'was 25
   MsgLine2               As String * 22     'was 25
   MsgLine3               As String * 22     'was 25
   MsgLine4               As String * 22     'was 25
   CrLf                   As String * 2
End Type


Type MowBillOutRecType   'Mow
   AcctNo                 As String * 8
   LocationNum            As String * 9        'format ##-######
   CUSTNAME               As String * 35
   ADDR1                  As String * 35
   ADDR2                  As String * 35
   SERVADDR               As String * 35
   CITY                   As String * 18
   STATE                  As String * 2
   ZIPCODE                As String * 10
   BillType               As String * 1        'N=Normal F=Final
   DepAppAmt              As String * 10       'if final applied deposit amt.
   PrevDue                As String * 15       'total revenue amount
   CurrDue                As String * 15       'total revenue amount
   TotalDue               As String * 15       'total revenue amount
   MTRType                As String * 1        'w=water s=sewer c=combined
   MTRUnit                As String * 1
   CurrDate               As String * 10
   PrevDate               As String * 10
   ServDays               As String * 4
   CurrRead               As String * 10       'Last/Current meter readings
   PrevRead               As String * 10       'Previous readings
   Consump                As String * 10
   IServDays              As String * 4
   ICurrRead              As String * 10       'Last/Current meter readings
   IPrevRead              As String * 10       'Previous readings
   IConsump               As String * 10

   ServInfo(1 To 15)      As BillOutServType
   'this is the structure above

   BillDate               As String * 10
   PastDueDate            As String * 10
   DraftDate              As String * 10
   MsgLine1               As String * 25     'was 25
   MsgLine2               As String * 25     'was 25
   MsgLine3               As String * 25     'was 25
   MsgLine4               As String * 25     'was 25
   CrLf                   As String * 2
End Type

Type OJUBPaymentRecType
    OPERNUM         As Integer
    PAYDATE         As Integer
    CustAcct        As Long
    CUSTNAME        As String * 24
    CUSTADDR        As String * 24
    CUSTCMNT        As String * 32
'052598 Added tax exempt flag changed cust comment len to 32
    TaxExempt       As String * 1
    AMTOWED         As Double
    TENDERTY        As String * 12
    CASHAMT         As Double
    CHKAMT          As Double
    AMTRECD         As Double
    CHANGE          As Double
    DESC            As String * 20
    PaidOwed(1 To 15)   As PaidOwedType
    TOTOWED         As Double
    AMTPAID         As Double

End Type

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

Type LockBoxRecType
  AcctNum     As String * 6     '1-6
  Amount      As String * 10    '7-16
  TenderType  As String * 2     '17/18 CH for Check, CA for cash
  ChkNum      As String * 8     '19/26 chk num
  PaymentDate As String * 10    '27-36 01/01/1999
  Fill1       As String * 44    '37/80
  CrLf        As String * 2     '81/82 CrLf
End Type

'First Citizens
Type LockBoxRecTypeFC
  AcctNum     As String * 8     '1-6   Utility Billing Account Number
  Amount      As String * 10    '7-16  Decimal assumed (2.00 = 200) in record
  TenderType  As String * 2     '17-18 CH for Check, CA for cash
  ChkNum      As String * 8     '19-26 Check Number
  PaymentDate As String * 10    '27-36 Payment Date. Formate: MM-DD-YYYY
  Blank       As String * 42    '37-80 N/A
  CrLf        As String * 2     '81-82 CR+LF
End Type
'' UBBADGER.BI
Type UBBadgerRecType
  Fill1          As String * 8   'fill spaces
  CUSTNAME       As String * 20
  SERVADDR       As String * 20  'lj
  MtrNum1        As String * 9   'Meter serial number
  Multi          As String * 4   'n/a fill spaces
  Status         As String * 1   'Acct Status 'A' or 'I'
  ReadCode       As String * 1   'Read Code fill space
  ServFreq       As String * 2   'Serv Code   'set to '1M'
  DNI            As String * 2   'Dialog ID    'space fill
  MtrNum2        As String * 9   'Same data as field above
  NumDials       As String * 1   'N/A fill space
  HiRead         As String * 9   'High reading limit
  LoRead         As String * 9   'Low reading limit
  CurrRead       As String * 9   'Current reading or space fill
  ReadTime       As String * 8   'read time  space fill
  ReadCode2      As String * 2   'space fill 'NOT same as above'
  CmntCode       As String * 2   'space fill
  Fill2          As String * 4   'space fill
  Account        As String * 15  'location  '01-000001-01
  ReadDate       As String * 8   'Read date ddmmyyyy
  DevCode        As String * 1   'don't have a clueEnd Type
  MMILat         As String * 6   'space fill
  MMILong        As String * 6   'space fill
  MMIChanl       As String * 6   'space fill
  CircleCode     As String * 2   'space fill
  SEQNUMB        As String * 6   'Read sequence number
  MfgModel       As String * 20  'space fill
  UserField      As String * 30  'we will use
  ReadID         As String * 3   'space fill
  ReadCo1        As String * 2   'space fill
  ReadCo2        As String * 2   'space fill
  ReadCo3        As String * 2   'space fill
  MMIReadCode    As String * 6   'space fill
  Pad            As String * 19  'space fill
  CrLf           As String * 2   'Carriage return line feed sequence
End Type
Type BADReadRecType
  Fill          As String * 128
End Type

Type BADReadRecType0
 RecordID       As String * 3
'***
 Pad1           As String * 11 '14
 Account        As String * 20 '34
'***
 CurRead        As String * 10 '44
 pad2           As String * 3  '47
 ReadDate       As String * 8  '56
 Pad            As String * 73 '
End Type

Type FDRTYPE
 RecordID       As String * 3
 TABLE          As String * 1
 PROBE          As String * 1
 VERS           As String * 5
 Reserved       As String * 3
 CYCLES         As String * 2
 RADIOREADYN    As String * 1
 WANDYN         As String * 1
 EXTFORMATYN    As String * 1
 Pad            As String * 108
End Type

Type CDRTYPE
 RecordID       As String * 3
 CycleNumber    As String * 2
 NumberCycles   As String * 4
 CycleDate      As String * 8
 Pad            As String * 109
End Type

Type RDRType
 RecordID        As String * 3
 RouteNumber     As String * 8
 SurveyYN        As String * 1
 RouteMessYN     As String * 1
 Keys            As String * 4
 Readings        As String * 4
 Demands         As String * 4
 Keyed           As String * 4
 Probed          As String * 4
 Radio           As String * 4
 Customers       As String * 4
 Meters          As String * 4
 TimeAllowed     As String * 6
 Gas             As String * 4
 Water           As String * 4
 Electric        As String * 4
 Location        As String * 4
 Extra           As String * 4
 Region          As String * 2
 ZONE            As String * 2
 Office          As String * 2
 BillCycle       As String * 2
 DropCycle       As String * 2
 WandReads       As String * 4
 AMR             As String * 1
 Pad             As String * 40
End Type

Type CUSType
 RecordID       As String * 3
 RouteNumber    As String * 8
 Meters         As String * 3
 AccountNumber  As String * 20
 Name           As String * 20
 ADDRESS1       As String * 20
 ADDRESS2       As String * 20
 Reserved       As String * 2
 Group          As String * 1
 CustInfo       As String * 20
 ExCustRec      As String * 1
 Segment        As String * 4
 UtilityID      As String * 2
 PassThur       As String * 1
 Pad            As String * 1
End Type

Type MTRType
 RecordID       As String * 3   '3
 RouteNumber    As String * 8   '11
 Reads          As String * 3   '14
 Reserved       As String * 2   '16
 Group          As String * 1   '17
 Reserved1      As String * 1   '18
 Reserved2      As String * 7   '25
 Survey         As String * 1   '26
 Reserved3      As String * 2   '28
 Survey2        As String * 1   '29
 BillCode       As String * 1   '30
 MtrStatus      As String * 1   '31
 OpticalProbe   As String * 14  '45
 MtrNumber      As String * 12  '57
 Reserved4      As String * 2   '59
 MeterType      As String * 2   '61
 MeterSeq       As String * 8   '69
 MeterInfo      As String * 20  '89
 Reserved5      As String * 1   '90
 Location       As String * 2   '92
 Reserved6      As String * 1   '93
 ReadInst1      As String * 2   '95
 Reserved7      As String * 1   '96
 ReadInst2      As String * 2   '98
 SpecMessage    As String * 1   '99
 Reserved8      As String * 1   '100
 SpecMessageYN  As String * 1   '101
 MtrCat         As String * 1   '102
 LocExtraMtr    As String * 1   '103
 TimeCode       As String * 3   '106
 MtrAudit1      As String * 2   '108
 MtrAudit2      As String * 2   '110
 MtrAudit3      As String * 2   '112
 MtrAudit4      As String * 2   '114
 Pad            As String * 12  '126
End Type

Type RDGType
 RecordID       As String * 3
 RouteNumber    As String * 8  '11
 Text           As String * 4  '15
 Prompt         As String * 1  '16
 ReadDir        As String * 1  '17
 Compare        As String * 3  '20
 Validation     As String * 3  '23
 Reserved       As String * 1  '24
 Channel        As String * 2  '26
 Dials          As String * 2  '28
 Decimals       As String * 2  '30
 ReadMethod     As String * 1  '31
 PrevRead       As String * 10 '41
 High1          As String * 10 '51
 Low1           As String * 10 '61
 MtrConstant    As String * 6  '67
 ConstantFlag   As String * 1  '68
 HHFFlag        As String * 1  '69
 PosCreep       As String * 5  '74
 Estimates      As String * 1  '75
 Reserved2      As String * 1  '76
 ReadType       As String * 2  '78
 MaxPercent     As String * 6  '84
 MinPercent     As String * 6  '
 NegCreep       As String * 5  '95
 Pad            As String * 31 '
End Type

Type RFFType
 RecordID       As String * 3
 RouteNumber    As String * 8
 RFERT          As String * 8
 Reserved       As String * 6
 RFProgram      As String * 4
 Reserved2      As String * 7
 GeoArea        As String * 2
 RFFreq         As String * 12
 RFTone         As String * 4
 Reserved3      As String * 10
 Tamper         As String * 2
 ConcIndicator  As String * 1
 Pad            As String * 59
End Type

Type WRRType
 RecordID       As String * 3
 RouteNumber    As String * 8
 DeviceID       As String * 14
 WandProg       As String * 4
 Resv1          As String * 5
 Resv2          As String * 1
 NodeNumb       As String * 2
 NoTamper       As String * 1  'y/n "Y" = no Check
 Fill1          As String * 88
End Type

Type UBJettHHRecType         '
' First two fields are required by CE File I/O dll.
  CEVariant       As String * 2  'This can be blank or anything your heart desires
  CEStrLen        As String * 2  'same as above
'**************************************************************
  CUSTNAME      As String * 20
  ServAddress   As String * 20
  MeterID       As String * 8
  LowRead       As String * 9
  HighRead      As String * 9
  Account       As String * 8
  ReadFlag      As String * 1         'Y/N
  MeterType     As String * 1
  BOOK          As String * 2
  CurRead       As String * 9
  PastRead      As String * 9
  ReadDate      As String * 8   ' fmt mmddyyyy
  ReadTime      As String * 6   ' fmt hhmmss
  Note1         As String * 20
  Note2         As String * 20
  Note3         As String * 20
  NoteStatus    As String * 1  'T=Temp  P=Perm
End Type

'-----------------------------------------------------------------------

Type EPNewUBCustRecType
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
    ProratePCT    As Integer
    HHMSG1        As String * 20
    HHMSG2        As String * 20
    HHMSG3        As String * 20
'Page 3
    Serv(1 To 15)      As ServicesType   'Nochange
    FlatRates(1 To 4)  As FlatRateType   'Nochange
'Page 4
    Monthly(1 To 2)    As MonthlyPayType 'Nochange
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




