TYPE WrightCustType
  Cycle        AS STRING * 1
  Route        AS STRING * 2
  CustNO       AS STRING * 6
  CustName     AS STRING * 23
  Addr1        AS STRING * 23
  Addr2        AS STRING * 23
  City         AS STRING * 21
  State        AS STRING * 2
  ZipCode      AS STRING * 9   'actual two fields "5 & 4"
  ServAddr     AS STRING * 23
  Status       AS STRING * 1
  TaxExmpt     AS STRING * 1
  DraftYN      AS STRING * 1
  ELCode       AS STRING * 2
  ELMeterNo    AS STRING * 5
  ELMulti      AS STRING * 5
  ELPrevRead   AS STRING * 5
  ELCurrRead   AS STRING * 5
  SecLght1     AS STRING * 2
  SecLght1t    AS STRING * 1
  SecLght2     AS STRING * 2
  SecLght2t    AS STRING * 1
  SecLghtU     AS STRING * 5
  SecLghtCH    AS STRING * 6
  WaterClass   AS STRING * 2
  SewerClass   AS STRING * 2
  MeterIDate   AS STRING * 6   'MO DA YR
  WatPrevRead  AS STRING * 8
  WatCurrRead  AS STRING * 8
  MiscCharge   AS STRING * 8
  WatMinimum   AS STRING * 5
  AmountDue    AS STRING * 10
  ServFrom     AS STRING * 6   'MO DA YR
  ServTo       AS STRING * 6   'same
  GarbRate     AS STRING * 2
  ELRate       AS STRING * 1
  WatRate      AS STRING * 1
  SewRate      AS STRING * 1
  BankCode     AS STRING * 2
  StreetCode   AS STRING * 1
  RouteSeqNo   AS STRING * 4
  LocSearch    AS STRING * 15
  NameSearch   AS STRING * 15
  InputFee     AS STRING * 6
  InputMisc    AS STRING * 8
  CustType     AS STRING * 2
  MeterNo      AS STRING * 9
  MeterLoc     AS STRING * 16
  MeterSize    AS STRING * 2
  NumOfDials   AS STRING * 1
  CustIDate    AS STRING * 6
  CrLf         AS STRING * 2
END TYPE

