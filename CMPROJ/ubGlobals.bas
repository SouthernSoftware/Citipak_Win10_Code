Attribute VB_Name = "ubGlobals"
Option Explicit
Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)
Global GLSetup As GLSetupRecType
Global GLAcctidx As GLAcctIndexType
Global GLAcct As GLAcctRecType
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
Global ScreenW As Long
Global OperNum As Integer
Global ThisCustXNum As Integer 'used when opening frmBLTransHistJr modally
Global GCustNum As Long
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
Global PrnVali As Boolean
Global BnkAcctNum As String
'added for receipt printer default value of yes(1) or no(0)
Global RecpDef As Integer
'added for receipt printer default for Control Codes Yes(1) or No(0)
Global CntrlDef As Integer
'added for validation default  yes(1) no(0)
Global ValiDef As Integer

Global intHasTaxes As Integer

Global Const ServiceAddressIndexFile = "UBSVCADD.IDX"
Global Const UBBillsFile = "UBBILLS.DAT"
Global Const UBIBillFile = "UBIBILL.DAT"
Global Const UBFinPreRptFile = "UBPREFIN.RPT"
Global Const UBFinBillsFile = "UBFBILLS.DAT"
Global Const RePrintIdxFile = "UBREPRNT.IDX"
Global Const BLCatCodeName = "ARCODE.DAT"
Global Const BLCustFileName = "ARCUST.DAT"
Global Const BLTransFileName = "ARTRANS.DAT"
Global Const CatCodeIdxName = "arcatcodeidx.dat"
Global Const CustNameIdx = "arcustnameidx.dat"
Global Const LicNumIdx = "arlicnumidx.dat"
Global Const CustNumIdx = "arcustnumidx.dat"
Global Const CustSearchNameIdx = "arsrhidx.dat"
Global Const BLTransTempPost = "artmppst.dat"
Global Const BLPayFileName = "AREDPY"
Global Const BLTownSetUpName = "artownsu.dat"
Global Const BLTempPenaltyCharges = "artmppen.dat"

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
Type GLSetupRecType                 'V205 added new fields noted below
   UserName    As String * 30
   TotAcctLen  As Integer
   FundLen     As Integer
   AcctLen     As Integer
   DetLen      As Integer
   CASHACCT    As String * 14
   APAcct      As String * 14
   EncAcct     As String * 14
   FBAcct      As String * 14
   FYBeg       As Integer
   FYEnd       As Integer
   NYBeg       As Integer
   NYEnd       As Integer
   CDCash      As String * 14  'new
   CDDue       As String * 14
   CDActive    As String * 1
   CRCashAcct  As String * 14
   DeptCode    As String * 1
   LPDate      As Integer
   HPDate      As Integer
   CDCashAcct  As String * 14
   CDsbCash    As String * 14
   APChkCode   As Integer
   POStop      As Boolean          'new 7/22/02 for potab on invoice entry
 'Fields added for V205
   PSLFlag     As Integer   '1 for default to Yes, 0 for No
   DupInvFlag  As Integer   '1 to allow duplicates, 0 for No
   CRBank      As Integer   'banknum as default on entry
   CDBank      As Integer   'banknum for default on entry
   ChkBank     As Integer   'banknum for default on check printing
   pad         As String * 20
   ChkVer      As String * 4   ' for "V205"
End Type
Type GLAcctIndexType                'Account Index: 16 bytes
   AcctNum     As String * 14       'Formatted account Number string
   RecNum      As Integer           'Pointer to record
   '*****
End Type
Type GLAcctRecType                  'Account Record Type: ? bytes
   Deleted     As Integer           'Active Account Flag
   Num         As String * 14       'Formatted Account Number
   Title       As String * 30       'Account Description
   Typ         As String * 1        'Account Type
   FrstTran    As Long              'Pointer to First Trans
   LastTran    As Long              'Pointer to Last Trans
   PYAct       As Double            'Prior Year Actual
   BegBal      As Double            'Beginning Balance
   Bgt         As Double            'Budget Amount
   Bal         As Double            'Running Balance
   Encumb      As Double            'Encumbered Amount
   MTD         As Double            'Month to Date Bal (calc as needed)
   YTD         As Double            'Year to Date Bal (calc as needed)
   NYEst       As Double            'Bgt Estimate
   NYReq       As Double            'New Year Bgt Requested
   NYRec       As Double            'New Year Bgt Recommended
   NYApp       As Double            'New Year Bgt Approved
   FrstBTran   As Integer           'Pointer to First Budget Trans
   LastBTran   As Integer           'Pointer to Last Budget Trans
   FrstPTran   As Integer           'Pointer to First Budget Trans
   LastPTran   As Integer           'Pointer to Last Budget Trans
   'Res         AS STRING * 25       'Reserved for future needs
   Work        As Double            'Temp added 08/17/96 for closeout
'edit the res added function rec pointer 6/11/04
   FNCTRec     As Long
   Res         As String * 12
   ChkByte     As String * 1    'this is updated at GASB34 conversion with chr$(1)
   Marked      As Integer           '
End Type

'old one below
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
Type CMBankAcctRecType
    COMPACCT As String * 20
End Type

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
  pad        As String * 4
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

Type LocMeterType
    MtrNum    As String * 12
    MTRMulti  As Integer
    MtrType   As String * 1
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
    MtrIDNO   As String * 11
    MtrLat    As Double
    MtrLng    As Double
End Type

Type MonthlyPayType
    AmtOwed      As Double
    TotAmtPD     As Double
    PayAmt       As Double
    RevSource    As Integer
End Type

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
'Change for extra meter fields 7/7/04
Type NewUBCustRecType
    BOOK          As String * 2
    SEQNUMB       As String * 6
    Status        As String * 1
    OPENDATE      As Integer
    SEARCH        As String * 10
    CustName      As String * 35
    Addr1         As String * 35
    Addr2         As String * 35
    SERVADDR      As String * 35
    City          As String * 18
    State         As String * 2
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
    Seq           As Long
'Page 2
    CASHONLY      As String * 1
    LATEFEE       As String * 1
    CUTOFFYN      As String * 1
    TAXEXPT       As String * 1
    SRCIT         As String * 1
    EPPFlag       As String * 1
'032299 Modified for Bank draft account type
'    EPPAMT        AS DOUBLE
'added GroupCoderec 2/1/05 for pointer to bookcode
    GroupCodeRec  As Integer
    Filler1       As String * 5
   ' Filler1       As String * 7
    USEDRAFT      As String * 1
    AcctType      As String * 1
'032299 Inserted account type
    BankName      As String * 34
    BANKLOC       As String * 30
    TRANSIT       As String * 9
    BankAcct      As String * 20
    BILLCMNT      As String * 25
    PAYCMNT       As String * 25
    PumpCode      As String * 4
    USERCODE1     As String * 4
    USERCODE2     As String * 2
    ProratePCT    As Integer
    HHMSG1        As String * 20
    HHMSG2        As String * 20
    HHMSG3        As String * 20
'Page 3
    serv(1 To 15)      As ServicesType
    FlatRates(1 To 4)  As FlatRateType
'Page 4
    Monthly(1 To 2)    As MonthlyPayType
    MFEE1         As Double
    MFEE2         As Double
    LocMeters(1 To 7)  As LocMeterType
'END OF Quick Screen Form
    CustPin       As Long
    LastTrans     As Long
    CurrBalance   As Double
    PrevBalance   As Double
    CurrRevAmts(1 To 15) As Double
    PrevRevAmts(1 To 15) As Double
    DepositAmt    As Double
    DelFlag       As Integer
    PreNoteFlag   As Integer
    WOLastTrans   As Long            'work order last trans pointer
    EstFlag       As String * 1
    MessageRec    As Long            ' Points to Message Record
    OldRec        As Long
    EPPLastTran   As Long
    NewNotes      As Integer
    DPCode        As String * 2
    FillPad       As String * 112
    ChkByte       As String * 1
End Type
'Type GroupCodeIndexType
'    RecordNum   As Integer
'    GroupCODE   As String * 2
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

''  Type NewUBCustRecType
''    Book          As String * 2
''    SEQNUMB       As String * 6
''    Status        As String * 1
''    OPENDATE      As Integer
''    SEARCH        As String * 10
''    CustName      As String * 35
''    ADDR1         As String * 35
''    ADDR2         As String * 35
''    ServAddr      As String * 35
''    CITY          As String * 18
''    STATE         As String * 2
''    ZIPCODE       As String * 10
''    HPHONE        As String * 14
''    WPHONE        As String * 14
''    SOSEC         As String * 11
''    DRVLIC        As String * 16
''    CUSTTYPE      As String * 3
''    Addr911       As String * 14
''  '051498 added bill to field. Removed 1 byte from 911 addr
''    BillTo        As String * 1
''  '********************************************************
''    BILLCOPY      As Integer
''    POSTRTE       As String * 4
''    BILLCYCL      As Integer
''    ZONE          As String * 3
''    Seq           As Long
''  'Page 2
''    CASHONLY      As String * 1
''    LATEFEE       As String * 1
''    CUTOFFYN      As String * 1
''    TAXEXPT       As String * 1
''    SRCIT         As String * 1
''    EPPFlag       As String * 1
''  '032299 Modified for Bank draft account type
''  '    EPPAMT        AS DOUBLE
''    Filler1       As String * 7
''    USEDRAFT      As String * 1
''    AcctType      As String * 1
''  '032299 Inserted account type
''    BankName      As String * 34
''    BANKLOC       As String * 30
''    TRANSIT       As String * 9
''    BankAcct      As String * 20
''    BILLCMNT      As String * 25
''    PAYCMNT       As String * 25
''    PumpCode      As String * 4
''    USERCODE1     As String * 4
''    USERCODE2     As String * 2
''    ProRatePCT    As Integer
''    HHMSG1        As String * 20
''    HHMSG2        As String * 20
''    HHMSG3        As String * 20
''  'Page 3
''    Serv(1 To 15)      As ServicesType
''    FlatRates(1 To 4)  As FlatRateType
''  'Page 4
''    Monthly(1 To 2)    As MonthlyPayType
''    MFEE1         As Double
''    MFEE2         As Double
''    LocMeters(1 To 7)  As LocMeterType
''  'END OF Quick Screen Form
''    CustPIN       As Long
''    LastTrans     As Long
''    CurrBalance   As Double
''    PrevBalance   As Double
''    CurrRevAmts(1 To 15) As Double
''    PrevRevAmts(1 To 15) As Double
''    DepositAmt    As Double
''    DelFlag       As Integer
''    PreNoteFlag   As Integer
''    WOLastTrans   As Long            'work order last trans pointer
''    EstFlag       As String * 1
''    MessageRec    As Long            ' Points to Message Record
''    OldRec        As Long
''    EPPLastTran   As Long
''    NewNotes      As Integer
''    FillPad       As String * 4
''    ChkByte       As String * 1
''  End Type
''
''  Type WrkOrdTextType
''  Text(1 To 6)  As String * 67
''  End Type
''
''  Type WorkOrderRecType
''    CustRec           As Long
''    ENTRYDATE         As Integer
''    OrdersText        As WrkOrdTextType
''    RepliesText       As WrkOrdTextType
''    CompleteByDate    As Integer
''    CompletedDate     As Integer
''    PrevTransRec      As Long
''  End Type
''
''  Type WorkOrderDefType
''    Deleted           As Boolean
''    WOType            As String * 20
''    OrdersText        As WrkOrdTextType
''    RepliesText       As WrkOrdTextType
''    Xtra              As String * 20 'just in case
''  End Type

Type Newport
    Acct As String * 6
    Name As String * 31
    Address As String * 25
    City As String * 15
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
    reserved     As String * 30
End Type


Type DistArrayType
   DistOrder As Integer
   DistCnt   As Integer
End Type


Type UBDGProcRecType        ' File Layout for Sending Out Records
    RouteID As String * 20
    SvcTyp As String * 1
    CustName As String * 25
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
   DraftDate              As Integer      '
'111398
   ProratePCT             As Integer
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
    PumpCode         As String * 4
    CustCnt          As Long
    Consump          As Double
End Type

Type UBOwnerRecType
    OwnLName  As String * 20
    OwnFName  As String * 15
    Addr1     As String * 35
    Addr2     As String * 35
    City      As String * 18
    State     As String * 2
    ZIPCODE   As String * 10
    HPHONE    As String * 14
    WPHONE    As String * 14
    ChkByte   As String * 1
End Type

Type MtrNotesType
  Note    As String * 20
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
  
Type PaidOwedType
    AMTOWE1  As Double
    AMTPD1   As Double
End Type
  
Type UBPaymentRecType
    OperNum         As Integer
    payDate         As Integer
    CustAcct        As Long
    CustName        As String * 24
    CustAddr        As String * 24
    CUSTCMNT        As String * 32
'052598 Added tax exempt flag changed cust comment len to 32
    TaxExempt       As String * 1
    AmtOwed         As Double
    TenderTY        As String * 12
    CashAmt         As Double
    ChkAmt          As Double
    AmtRecd         As Double
    Change          As Double
    Desc            As String * 19
    PaidOwed(1 To 15)   As PaidOwedType
    TotOwed         As Double
    AmtPaid         As Double
'112801 Added cust status changed desc to 19
    Status          As String * 1
End Type
Type PayListType
  CustRec   As Long
  ListRec   As Long
End Type

Type BookTotalType
  Count   As Long
  Cash    As Double
  Check   As Double
  Charge  As Double
  Change  As Double
End Type

'CM
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
    InactiveFlag As String * 1
    NotUsed As String * 17
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
Type CMSetupType
    CMTOWNNAME   As String * 30
    GLInterface  As String * 1
    Pass4Voids   As String * 1  'Y -yes, N- no, F - full access only
    VoidPW       As String * 10
    Pass4Adj     As String * 1  'Y -yes, N- no, F - full access only
    AdjPW        As String * 10
    Filler       As String * 75  '128
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
  PaymDate  As Integer         'For Changing Default Date During Daily Entry
  RValidate As Integer
  ZExtra    As String * 16
End Type

Type struct
  who As String * 14
  'change RecNum from integer on 3-1-04 PKS
  RecNum As Long
End Type

Type ARCustRecType
    CUSTNUMB As String * 10
    SORTNAME As String * 10
    BILLNAME As String * 35
    ADDRESS1 As String * 35
    ADDRESS2 As String * 35
    City     As String * 20
    State    As String * 2
    ZIPCODE  As String * 10
    CustName As String * 35
    Contact  As String * 30
    BILLCAT1     As String * 5
    DESC1        As String * 35
    REV1         As Long
    Fee1         As Double 'latest charge
    FeeLicBal1   As Double
    FeeLicPay1   As Double
    BILLCAT2     As String * 5
    DESC2        As String * 35
    REV2         As Long
    Fee2         As Double
    FeeLicBal2   As Double
    FeeLicPay2   As Double
    BILLCAT3     As String * 5
    DESC3        As String * 35
    REV3         As Long
    Fee3         As Double
    FeeLicBal3   As Double
    FeeLicPay3   As Double
    BILLCAT4     As String * 5
    Desc4        As String * 35
    REV4         As Long
    Fee4         As Double
    FeeLicBal4   As Double
    FeeLicPay4   As Double
    BILLCAT5     As String * 5
    Desc5        As String * 35
    REV5         As Long
    Fee5         As Double
    FeeLicBal5   As Double
    FeeLicPay5   As Double
    IssuanceFee  As Double
    CustLocation As String * 1
    WPHONE       As String * 14
    FeeAmt       As Double
    LICENSE      As String * 12
    valid        As Integer
    Inactive     As String * 1    '"Y" if account is inactive
    ProRate      As Integer       'prorate percentage
    AcctBal      As Double
    IssueLicense As String * 1    'y/n
    Deleted      As String * 1    '(yY)=deleted, anything else isn't
    FirstTrans   As Long
    LastTrans    As Long
    LicBal       As Double
    FeeBal       As Double
    PenBal       As Double
    RoomtoGrow   As String * 136
    ChkByte      As String * 1
    IssuanceBal  As Double
    IssuancePay  As Double
    ServAdd      As String * 35
    SSNFID       As String * 15 'Paula...add this
End Type

Type ARNewCatCodeRecType
    CATCODE    As String * 5    'Not Used in Version 8.5 work2 directory
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

Type ARCustIDXRecType
    IDXName     As String * 10
    IDXRECORD   As Integer
    ExtraRoom   As String * 52
End Type

Type ARTransRecType
    CustomerNumber      As String * 10
    TransDate           As Integer
    TransAmount         As Double
    TransType           As Integer
    TransDesc           As String * 35 '5
    CashAmount          As Double
    ChkAmount           As Double
    BalanceAfterTrans   As Double
    Posted2GL           As String * 1
    CatCodeRec1         As Long  '10         'Place to Grab G/L Acct #'s
    CatCodeRec2         As Long           'Place to Grab G/L Acct #'s
    CatCodeRec3         As Long           'Place to Grab G/L Acct #'s
    CatCodeRec4         As Long           'Place to Grab G/L Acct #'s
    CatCodeRec5         As Long           'Place to Grab G/L Acct #'s
    CatLicAmt1          As Double '15
    CatLicAmt2          As Double
    CatLicAmt3          As Double
    CatLicAmt4          As Double
    CatLicAmt5          As Double
    CatLicBal1          As Double '25
    CatLicBal2          As Double
    CatLicBal3          As Double
    CatLicBal4          As Double
    CatLicBal5          As Double
    PenBal              As Double
    LicBal              As Double
    IssBal           As Double
    FeeAmt              As Double
    LicAmt              As Double
    PenAmt              As Double
    IssAmt              As Double
    ExtraRoom           As String * 8
    NextTrans           As Long
    DetailTransType     As Integer 'used for reading transaction types inside BL program for reports...not GL
    'Codes for General Ledger:
    '1 = all non-penalty charges; 2 = all payments; 6 = all penalty charges; 13 = adjust payment down
    '23 = adjust billing down; 24 = adjust billing up
    'Codes for internal Business License:
    '101 = Charge Penalty ; 110 = Charge Lic; 201 = Pay Penalty; 210 = Pay Lic; 211 = Pay Lic and Penalty; 301 = Adjust Down Pen; 310 = Adjust Down Lic
    '311 = Adjust Down Pen and Lic; '401 = Adjust Up Pen; 410 = Adjust Up Lic; 411 = Adjust Up Lic and Penalty
End Type


Type AREditPaymentRecType
    TranType        As Integer
    TranDate        As Integer
    CustNumber      As String * 10
    CustName        As String * 35
    Add1            As String * 35
    City            As String * 25
    State           As String * 2
    ZIPCODE         As String * 10
    Amount          As Double
    CASHCHK         As String * 9
    CashAmt         As Double
    ChkAmt          As Double
    CREDITAM        As Double
    AmtPaid         As Double
    Change          As Double
    ISSUELIC        As String * 1
    SetFee          As String * 1
    ISSueFEE        As Double
    Desc            As String * 20
    LICDUE          As Double
    LICDUE1         As Double
    LICDUE2         As Double
    LICDUE3         As Double
    LICDUE4         As Double
    LICDUE5         As Double
    LICPAID         As Double
    LICPAID1        As Double
    LICPAID2        As Double
    LICPAID3        As Double
    LICPAID4        As Double
    LICPAID5        As Double
    TotDue          As Double
    TotPaid         As Double
    CatDesc1        As String * 35
    CatDesc2        As String * 35
    CatDesc3        As String * 35
    CatDesc4        As String * 35
    CatDesc5        As String * 35
    PENDUE          As Double
    PENPAID         As Double
    ISSDUE          As Double
    ISSPAID         As Double
End Type

Type CatCodeIdxType
  CatCodeRec As Integer
  CatCodeNum As String * 20
End Type


Type CustNameIdxType
   BillingName As String * 35
   CustRec As Integer
End Type

Type CustLicNumIdxType
   LicNum As String * 12
   CustRec As Integer
End Type

Type CustNumIdxType
   CUSTNUMB As String * 10
   CustRec As Integer
End Type

Type CustSearchNameIdxType
   SORTNAME As String * 10
   CustRec As Integer
End Type

Type TransIdxType
  TransWho As String * 35
  TransRecNum As Double
  TransAmt As Double
End Type
Type TempPenaltyCharges
    CustomerNumber      As String * 10
    TransDate           As Integer
    TransAmount         As Double
    PenAmt              As Double
End Type

Type TempTransPostType
    CustomerNumber      As String * 10
    TransDate           As Integer
    TransAmount         As Double
    TransType           As Integer
    TransDesc           As String * 35
    BalanceAfterTrans   As Double
    Posted2GL           As String * 1
    CatCodeRec1         As Long           'Place to Grab G/L Acct #'s
    CatCodeRec2         As Long           'Place to Grab G/L Acct #'s
    CatCodeRec3         As Long           'Place to Grab G/L Acct #'s
    CatCodeRec4         As Long           'Place to Grab G/L Acct #'s
    CatCodeRec5         As Long           'Place to Grab G/L Acct #'s
    CatFee1             As Double
    CatFee2             As Double
    CatFee3             As Double
    CatFee4             As Double
    CatFee5             As Double
    CatFeeBal1             As Double
    CatFeeBal2             As Double
    CatFeeBal3             As Double
    CatFeeBal4             As Double
    CatFeeBal5             As Double
    LICENSE             As String * 12
    valid               As Integer
    LicBal              As Double
    AcctBal             As Double
    PenBal              As Double
    Prev                As Long
    CreditUsed          As Boolean
    IssFee              As Double
    IssFeeBal           As Double
End Type

Type TownSetUpType
    TownName As String * 38 'allow for TOWN OF
    TownAdd1 As String * 30
    TownAdd2 As String * 30
    Contact As String * 30
    City As String * 30
    State As String * 2
    ZIPCODE As String * 10
    TownPhone As String * 14
    AppForm As Integer
    DLQNotice As Integer
    SpareSpace As String * 60
    AppAdd1 As String * 30
    AppCity As String * 30
    AppState As String * 2
    AppPhone As String * 14
    AppAdminName As String * 25
    AppAdminTitle As String * 25 '17
    AppBaseFee(1 To 10) As Double
    AppCentsPer(1 To 4) As Double
    AppGrsRcpts(1 To 4) As Double '29
    AppFirstDay As String * 7
    AppLastDay As String * 7
    AppTownOf As String * 38
    AppZip As String * 10
    AppPct As Double
    AppGrsPct As Double
    AppDenom As Integer
    AppNumer As Integer '37
    AppColFee As Double
    AppPayBy As Integer
    AppDiscPct As Double
    AppDiscMonth As String * 9
    AppDiscDay As Integer
    AppPenMonth As String * 9
    AppPenDay As Integer
    AppFiscMonth As String * 9
    AppFiscDay As Integer
    AppMayorCouncil As String * 25
    AppWholeMonth As Integer
    AppWholeDay As Integer '52
    AppRetailMonth As Integer
    AppRetailDay As Integer
    AppFinMonth As Integer
    AppFinDay As Integer
    AppContMonth As Integer
    AppContDay As Integer
    AppRepairMonth As Integer
    AppRepairDay As Integer
    AppStartMonth As String * 9
    AppStartDay As Integer
    AppLicRetMonth As String * 9
    AppLicRetDay As Integer
    AppAdoptDate As Integer
    AppCityOrd As String * 40
    AppYrUpDown(1 To 10) As String * 4
    DlqTownName As String * 38
    DlqAdd1 As String * 30 '68
    DlqCity As String * 30
    DlqState As String * 2
    DlqZip As String * 10
    DlqPhone As String * 14
    DlqPhone2 As String * 14
    DlqFax As String * 14
    DlqAdminName As String * 25
    DlqAdminTitle As String * 25
    DlqFirstDay As String * 9
    DlqLastDay As String * 9
    DlqFirstHour As String * 9
    DlqLastHour As String * 9
    DlqClerkName As String * 25
    DlqMayorCouncil As String * 25 '82
    LicNumPermYN  As String * 3
    UseAmtPctYN   As String * 3
    PENREVGLNUM   As Long
    PENRECGLNUM   As Long
    PENCASHACCT   As Long
    IssFee        As Double
    AcctMeth      As String * 1
    LaserLtr      As String * 1
    GL2Cats       As String * 1
End Type

Type TempCustRecType
    CustRecNum As Integer
    AppType As Integer
    ThisYear As String * 4
    Fee(1 To 5) As Double
    CATCODE(1 To 5) As String * 5
    CatDesc(1 To 5) As String * 35
    MiscNum As Double
    AmtPct As String * 3
    IssFee As Double
End Type

'Decal files
Type DCSetupType
    DCTNNAME     As String * 30
    GLInterface  As String * 1
    AppType      As Integer
    DCVers       As String * 3
    Taxbalchk    As String * 1
    DefLook      As String * 1
    Filler       As String * 90  '128
End Type
Type DCExpireDate
    ExpireDate As String * 10
End Type

Type DCCustRecType
    CUSTNUMB     As String * 10
    SORTNAME     As String * 10
    BILLNAME     As String * 35
    ADDRESS1     As String * 35
    ADDRESS2     As String * 35
    City         As String * 20
    State        As String * 2
    ZIPCODE      As String * 10
    SOSEC        As String * 11
    DRVLIC       As String * 12
    DATEOPED     As Integer
    CASHONLY     As String * 1
    resident     As String * 1
    Owner        As String * 1
    HPHONE       As String * 14
    WPHONE       As String * 14
    LICENSE      As String * 12
    valid        As Integer
    AcctBal      As Double
    Deleted      As String * 1      'rem y=deleted :AnyThing Else is Non-Delet
    FirstTrans   As Long
    LastTrans    As Long
    FirstCar     As Long
    LastCar      As Long
    SocSec1      As String * 11
    OtherName    As String * 25
    RoomtoGrow   As String * 224
End Type

Type DCCatCodeRecType
    CATCODE    As String * 3
    CODEDESC   As String * 35
    APPNUMB    As Integer
    BILLCODE   As Integer
    REVGLNUM   As String * 14
    CASHACCT   As String * 14
    Fee        As Single
    InactiveFlag As String * 1
    Extra      As String * 53
End Type

Type DCCustIDXRecType
    IDXName     As String * 10
    IDXRECORD   As Long
    ExtraRoom   As String * 52
End Type

Type ZipIndexType
    IDXName     As String * 10
    IDXRECORD   As Long
End Type

Type DCTempIDXRecType
    IDXRECORD   As Long
End Type

Type DCTransRecType
    CustomerNumber As String * 10
    TransDate As Integer
    TransAmount As Double
    TransType As Integer     '1-charge,2-pay,3-voidchrg,4-voidpay
    TRVinDesc As String * 40
    CashAmount As Double
    ChkAmount As Double
    BalanceAfterTrans As Double
    makemodel As String * 25
    StateTag As String * 35
    ExpireDate As Integer
    Sticker As String * 12
    NextTrans As Long
    OperNum   As Long
    GLInterfaced  As String * 1
    DecalCat As String * 5
    TransTender  As Integer     'Type: 1=Cash, 2=Check, 3=Cash/Check, 4=Charge
''added tendertype and 3,4 transtypes above and  chkbyte to prevent reconverting PS 7/8/05
    VoidFlag As String * 1   'Y if voided
    ChkByte  As String * 1   'this is chr$(1)
    ExtraDesc As String * 26   'added extra
    VehRecord As Long
    ExtraRoom As String * 48
End Type
Type DCEditPaymentRecType
  VehRecord As Long
  CustNumber As String * 10
  CustName As String * 35
  CustAddr As String * 35
  TranDate As Integer
  Amount As Double
  DecalCat As String * 5
  Sticker As String * 12
  VinDesc As String * 35
  ExpDate As Integer
  makemodel As String * 25
  StateTag As String * 25
  resident As String * 1
  Owner As String * 1
  PersBuss As String * 1
  PayDesc As String * 23
  CashAmt   As Double
  CheckAmt  As Double
  Change    As Double
  OperNum   As Long
  TransTender  As Integer     'Type: 1=Cash, 2=Check, 3=Cash/Check, 4=Charge
  VoidFlag As String * 1   'Y if voided
  Notes As String * 39
End Type

Type DCVehType
  DecalCat As String * 5
  makemodel As String * 25
  StateTag As String * 35
  ExpireDate As Integer
  Sticker As String * 12
  valid As String * 1         'y/n means is it current
  Active As String * 1        'y/n  n=deleted record
  Notes As String * 39
  PBFlag  As String * 1
  Desc As String * 40
  Fee As Single
  MasterRecord As Long
  NextRec As Long
  MoreRoom As String * 83
End Type

