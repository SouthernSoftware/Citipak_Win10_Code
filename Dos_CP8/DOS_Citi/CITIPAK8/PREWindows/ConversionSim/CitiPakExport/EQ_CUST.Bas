Attribute VB_Name = "Module1"
'EPP
Const NameIndexFile = "UBCUSTNM.IDX"
Const BookIndexFile = "UBCUSTBK.IDX"
Const TempIndexName = "UBTEMP.IDX"
Const ServiceAddressIndexFile = "UBSVCADD.IDX"

Const UBBillsFile = "UBBILLS.DAT"
Const UBIBillFile = "UBIBILL.DAT"
Const UBFinPreRptFile = "UBPREFIN.RPT"
Const UBFinBillsFile = "UBFBILLS.DAT"
Const RePrintIdxFile = "UBREPRNT.IDX"


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

Type nUBCustReIndexRecType
  SearchName As String * 10
  DelFlag    As String * 1
  Status     As String * 1
  RecNum     As Long
End Type

Type UBLocaReIndexRecType
  BOOK       As String * 2
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
    RateCode As String * 4
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
    RateCode   As String * 4
    RevMtrType As String * 1
End Type

'Type LocMeterType
'    MTRNUM    As String * 12
'    MTRMulti  As Integer
'    MTRType   As String * 1
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
'End Type

Type MonthlyPayType
    AMTOWED      As Double
    TotAmtPD     As Double
    PayAmt       As Double
    RevSource    As Integer
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
'    EPPAMT        As Double
'
'    USEDRAFT      As String * 1
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
'    CurrRevAmts(1 To 15) As Double
'    'EPPBalances(1 TO 15) AS DOUBLE
'    PrevRevAmts(1 To 15) As Double
'    DepositAmt    As Double
'    DelFlag       As Integer
'    PreNoteFlag   As Integer
'    WOLastTrans   As Long            'work order last trans pointer
'    EstFlag       As String * 1
'    MessageRec    As Long            ' Points to Message Record
'    OldRec        As Long
'    EPPLastTran   As Long
'    FillPad       As String * 7
'End Type

Type WrkOrdTextType
  Text(1 To 6)  As String * 67
End Type

'Type WorkOrderRecType
'    CustRec           As Long
'    ENTRYDATE         As Integer
'    OrdersText        As WrkOrdTextType
'    RepliesText       As WrkOrdTextType
'    CompleteByDate    As Integer
'    CompletedDate     As Integer
'    PrevTransRec      As Long
'End Type

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

