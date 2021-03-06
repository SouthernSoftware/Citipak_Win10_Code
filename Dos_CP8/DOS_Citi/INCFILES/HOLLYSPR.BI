CONST NameIndexFile = "UBCUSTNM.IDX"
CONST BookIndexFile = "UBCUSTBK.IDX"
CONST TempIndexName = "UBTEMP.IDX"
CONST ServiceAddressIndexFile = "UBSVCADD.IDX"

CONST UBBillsFile = "UBBILLS.DAT"
CONST UBIBillFile = "UBIBILL.DAT"
CONST UBFinPreRptFile = "UBPREFIN.RPT"
CONST UBFinBillsFile = "UBFBILLS.DAT"
CONST RePrintIdxFile = "UBREPRNT.IDX"

TYPE AlarmAccountIdxType
  LocAcct   AS LONG
  CustRec   AS LONG
END TYPE

'Transaction Types
 CONST TranUtilityBill = 1          '   1=Utility bill
 CONST TranLateCharge = 2           '   2=late charge      'NOT USED
 CONST TranReconnectFee = 3         '   3=reconnect fee    'NOT USED
 CONST TranBillPayment = 4          '   4=Bill Payment
 CONST TranAppliedDeposit = 5       '   5=Applied Deposit
 CONST TranPenaltyCharge = 6        '   6=Penalty Charge
 CONST TranDepositPayment = 7       '   7=Deposit Payment
 CONST TranDraftPayment = 8         '   8=Draft Payment
 CONST TranRefundDeposit = 9        '   9=Refund Deposit
 CONST TranBeginBalance = 10        '  10=Beginning Balance
 CONST TranUpwardAdjustment = 11    '  11=Bill Adjustments
 CONST TranDownwardAdjustment = 12  '  12=Bill Adjustments
'added this for new over payment adjustment on Aug 11,2003
 CONST TranOverPayAdjustment = 33   '  33=OverPayment Adjustment
 CONST TranDepCreditRemoval = 37    '  37= Deposit Credit Removal - Not Interfaced W/GL
 CONST TranDepPaymentVoid = 39      '  39= Deposit Payment Void  - same gl as deposit refund
 CONST TranMiscPayment = 99         '  99=Misc Payment

'Meter Types
 CONST MtrWaterOnly = 1
 CONST MtrSewerOnly = 2
 CONST MtrCombined = 3
 CONST MtrElectric = 4
 CONST MtrDemand = 5
 CONST MtrGas = 6
 CONST MtrTouchRead = 7
 CONST MtrLightsService = 8
 CONST MtrIrrigation = 9

TYPE ServicesType
    Ratecode AS STRING * 4
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
  Ratecode   AS STRING * 4
  RevMtrType AS STRING * 1
END TYPE

TYPE LocMeterType
    MtrNum    AS STRING * 12
    MTRMulti  AS INTEGER
    MTRType   AS STRING * 1
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
'added GroupCoderec 2/1/05 for pointer to bookcode
    GroupCodeRec  AS INTEGER
    Filler1       AS STRING * 5
   ' Filler1       As String * 7
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
    serv(1 TO 15)      AS ServicesType
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
    CurrRevAmts(1 TO 15) AS DOUBLE   'includes the tax amount
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

