CONST NameIndexFile = "UBCUSTNM.IDX"
CONST BookIndexFile = "UBCUSTBK.IDX"
CONST TempIndexName = "UBTEMP.IDX"

CONST UBBillsFile = "UBBILLS.DAT"
CONST UBFinPreRptFile = "UBPREFIN.RPT"
CONST UBFinBillsFile = "UBFBILLS.DAT"
CONST RePrintIdxFile = "UBREPRNT.IDX"

TYPE UBPostalIndexType
  ZIPCODE  AS STRING * 10
  Route    AS STRING * 4
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

TYPE UBCustReIndexRecType
  SearchName AS STRING * 10
  First      AS STRING * 2
  RecNum     AS LONG
END TYPE

TYPE UBLocaReIndexRecType
  BOOK       AS STRING * 2
  SEQNUMB    AS STRING * 6
  RecNum     AS LONG
  Pad        AS STRING * 4
END TYPE

TYPE UBPINType
  PIN AS LONG
END TYPE

TYPE RevDataType
    REVNAME    AS STRING * 20
    RateCode   AS STRING * 4
    RevMtrType AS STRING * 1
END TYPE

TYPE FlatRateDataType
    FRDESC   AS STRING * 20
    FRAMT    AS DOUBLE
    FRFREQ   AS STRING * 1
    REVSRC   AS INTEGER
END TYPE

TYPE MeterDescType
    MTRNUM  AS STRING * 12     'meter number
    CURRD   AS DOUBLE          'Curr Read
'est flag   y/n
    CRDATE  AS INTEGER         'Curr date Read Date
'Past read date
'Past read amt
    MULT    AS INTEGER         'meter mulit
    MTRTYP  AS STRING * 1      'SoSoft defined meter type
    MTRUNT  AS STRING * 1      'Meter unit  /gal etc
    USER    AS INTEGER         'Minimum
    INDATE  AS INTEGER         'installed date
END TYPE

TYPE MonthlyPayType
    AMTOWED      AS DOUBLE
    TotAmtPD     AS DOUBLE
    PayAmt       AS DOUBLE
    RevSource    AS INTEGER
END TYPE

TYPE UBCustRecType
    STATUS   AS STRING * 1
    OPENDATE AS INTEGER
    SEARCH   AS STRING * 10
    CUSTNAME AS STRING * 35
    ADDR1    AS STRING * 35
    ADDR2    AS STRING * 35
    CITY     AS STRING * 18
    STATE    AS STRING * 2
    ZIPCODE  AS STRING * 10
    CUSTTYPE AS STRING * 3
    PostRte  AS STRING * 4
    BILLCOPY AS INTEGER
    HPHONE   AS STRING * 14
    WPHONE   AS STRING * 14
    SOSEC    AS STRING * 11
    DRVLIC   AS STRING * 16
    TAXEXPT  AS STRING * 1
    SRCIT    AS STRING * 1
    EPP      AS STRING * 1
    EPPAMT   AS DOUBLE
'page 2
    CASHONLY     AS STRING * 1
    LATEFEE      AS STRING * 1
    CUTOFFYN     AS STRING * 1

    Monthly(1 TO 2) AS MonthlyPayType

    DRAFTYPE     AS STRING * 1
    BANKNAME     AS STRING * 34
    BANKLOC      AS STRING * 30
    TRANSIT      AS STRING * 9
    BANKACCT     AS STRING * 20
    BILLCMNT     AS STRING * 25
    PAYCMNT      AS STRING * 25
'END OF QS FORM
    CustPIN           AS LONG
    OwnFrstLocRec     AS LONG
    OwnLastLocRec     AS LONG
    CustFrstLocRec    AS LONG
    CustLastLocRec    AS LONG
    Pend2Rec          AS LONG
    LastTrans         AS LONG
    CurrBalance       AS DOUBLE
    PrevBalance       AS DOUBLE
    CurrRevAmts(1 TO 15) AS DOUBLE
    DepositAmt        AS DOUBLE
    DelFlag           AS INTEGER
    PreNoteFlag       AS INTEGER
    PrevLocRec        AS LONG
    WOLastTrans       AS LONG
    'work order last trans pointer
    FillPad           AS STRING * 20
END TYPE

