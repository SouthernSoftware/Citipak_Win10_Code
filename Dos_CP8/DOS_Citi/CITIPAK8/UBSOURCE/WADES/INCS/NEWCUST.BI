CONST NameIndexFile = "UBCUSTNM.IDX"
CONST BookIndexFile = "UBCUSTBK.IDX"
CONST TempIndexName = "UBTEMP.IDX"
CONST ServiceAddressIndexFile = "UBSVCADD.IDX"

CONST UBBillsFile = "UBBILLS.DAT"
CONST UBIBillFile = "UBIBILL.DAT"
CONST UBFinPreRptFile = "UBPREFIN.RPT"
CONST UBFinBillsFile = "UBFBILLS.DAT"
CONST RePrintIdxFile = "UBREPRNT.IDX"

TYPE PSAZipIndexType
  ZIPCODE  AS STRING * 10
  SName    AS STRING * 10
  RecNum   AS INTEGER
  pad      AS STRING * 10
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
  Fill     AS STRING * 10      'This is to fill this to a 16 byte boundary
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

TYPE UBCustReIndexRecType
  SearchName AS STRING * 10
  First      AS STRING * 2
'  DelFlag    AS STRING * 1
'  Status     AS STRING * 1
  RecNum     AS LONG
END TYPE

TYPE UBLocaReIndexRecType
  BOOK       AS STRING * 2
  SEQNUMB    AS STRING * 6
  RecNum     AS LONG
  pad        AS STRING * 4
END TYPE

TYPE UBPINType
  PIN AS LONG
END TYPE

TYPE BookSeqRecType
  BookSeq  AS LONG
END TYPE

TYPE ServicesType
    RateCode AS STRING * 4
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
    REVNAME    AS STRING * 20
    RateCode   AS STRING * 4
    RevMtrType AS STRING * 1
END TYPE

TYPE LocMeterType
    MTRNUM    AS STRING * 12
    MTRMulti  AS INTEGER
    MTRType   AS STRING * 1
    MTRUnit   AS STRING * 1
    NumUser   AS INTEGER
    InsDate   AS INTEGER
    CurRead   AS LONG
    PrevRead  AS LONG
    CurDate   AS INTEGER
    PastDate  AS INTEGER       'hidden & protected
    ReadFlag  AS STRING * 1    'hidden & protected
    AvgUse    AS LONG          'hidden & protected
    UseCnt    AS INTEGER       'hidden & protected
END TYPE

TYPE MonthlyPayType
    AMTOWED      AS DOUBLE
    TotAmtPD     AS DOUBLE
    PayAmt       AS DOUBLE
    RevSource    AS INTEGER
END TYPE

TYPE NewUBCustRecType
    BOOK          AS STRING * 2
    SEQNUMB       AS STRING * 6
    Status        AS STRING * 1
    OPENDATE      AS INTEGER
    SEARCH        AS STRING * 10
    CUSTNAME      AS STRING * 35
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
    SEQ           AS LONG

'Page 2
    CASHONLY      AS STRING * 1
    LATEFEE       AS STRING * 1
    CUTOFFYN      AS STRING * 1
    TAXEXPT       AS STRING * 1
    SRCIT         AS STRING * 1
    EPP           AS STRING * 1
    EPPAMT        AS DOUBLE

    USEDRAFT      AS STRING * 1
    BANKNAME      AS STRING * 34
    BANKLOC       AS STRING * 30
    TRANSIT       AS STRING * 9
    BANKACCT      AS STRING * 20
    BILLCMNT      AS STRING * 25
    PAYCMNT       AS STRING * 25
    PUMPCODE      AS STRING * 4
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
    FillPad       AS STRING * 11
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

