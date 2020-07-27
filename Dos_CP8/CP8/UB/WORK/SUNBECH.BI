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
    ServAddr      As String * 35
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
