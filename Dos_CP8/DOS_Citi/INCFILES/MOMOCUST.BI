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
  RecNum   As Integer
  Pad      As String * 10
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
  Book       As String * 2
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
    MtrUnit   As String * 1
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
    CustName      As String * 35
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
    Filler1       As String * 7
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
