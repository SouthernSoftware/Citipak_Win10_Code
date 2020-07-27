TYPE PaidOwedType
   AmtOwed   AS DOUBLE
   AmtPaid   AS DOUBLE
END TYPE

TYPE TaxPaymentRecType
    OperNum  AS INTEGER
    PayDate  AS INTEGER
    CustAcct AS LONG
    CustName AS STRING * 24
    CustAddr AS STRING * 24
    AmtOwed  AS DOUBLE
    TenderTY AS STRING * 12
    CashAmt  AS DOUBLE
    ChkAmt   AS DOUBLE
    AmtRecd  AS DOUBLE
    Change   AS DOUBLE
    DESC     AS STRING * 20
    PaidOwed(1 TO 7)  AS PaidOwedType
   'end of quick screen form
    TotOwed  AS DOUBLE
    AmtPaid  AS DOUBLE
    TotPaid         AS DOUBLE
    LastPayRec      AS LONG          'Pointer to first payment list record
    NumPayRec       AS INTEGER       'Count of payment list records
    CustPIN         AS LONG
    BillType        AS STRING * 1       'R or P
END TYPE



TYPE PayListType
  BillRec       AS LONG      'Pointer to bill trans rec this payment is for
  Principle1    AS DOUBLE
  Principle2    AS DOUBLE
  Principle3    AS DOUBLE
  Principle4    AS DOUBLE
  Principle5    AS DOUBLE
  Interest     AS DOUBLE
  Penalty       AS DOUBLE
  TotPaid       AS DOUBLE    'amount paid to this bill rec (parital payment)
  CustRec       AS LONG      'backup pointer to cust rec
  PrevListRec   AS LONG      'pointer to next paylist rec
END TYPE

TYPE DMVHeader
  RecordType AS STRING * 2              'Must be H
  Batch      AS STRING * 7
  Jury       AS STRING * 5              'AAND for AshLand  LURG for Lunenburg Cty
  TaxYear    AS STRING * 5              'YYYY Format
  SubDate    AS STRING * 9              'MMDDYYYY Format
  TotalVeh   AS STRING * 8
  TotalAmt   AS STRING * 13
  Filler     AS STRING * 301
END TYPE

TYPE DMVRecord
  RecordType AS STRING * 2              'Must be D
  LineNumber AS STRING * 21
  SSN1       AS STRING * 10
  LastName1  AS STRING * 36
  FirstName1 AS STRING * 21
  Init1      AS STRING * 2
  SSN2       AS STRING * 10
  LastName2  AS STRING * 36
  FirstName2 AS STRING * 21
  Init2      AS STRING * 2
  Addr1      AS STRING * 33
  Addr2      AS STRING * 33
  City       AS STRING * 18
  State      AS STRING * 3
  Zip        AS STRING * 10
  VIN        AS STRING * 19
  VehValue   AS STRING * 8      'Whole Dollars
  PPTaxPd    AS STRING * 9      'Assume 2 Decimals
  PPTaxReimb AS STRING * 7      'Assume 2 Decimals
  PPTaxStart AS STRING * 7      'YYYYMM Format
  PPTaxEnd   AS STRING * 7      'YYYYMM Format
  Jury       AS STRING * 5      'AAND For Ashland
  SubDate    AS STRING * 9      'YYYYMMDD Format
  Filler     AS STRING * 21
END TYPE

TYPE DMVInformationType
 PerRate AS SINGLE
 Batch   AS LONG
 JCode   AS STRING * 4
END TYPE

