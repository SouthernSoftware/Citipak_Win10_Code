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
    DiscAmt  AS DOUBLE
    AmtRecd  AS DOUBLE
    Change   AS DOUBLE
    DESC     AS STRING * 20
    PaidOwed(1 TO 3)  AS PaidOwedType
    TotOwed  AS DOUBLE
    AmtPaid  AS DOUBLE
'end of quick screen form
    TotPaid         AS DOUBLE
    LastPayRec      AS LONG          'Pointer to first payment list record
    NumPayRec       AS INTEGER       'Count of payment list records
    CustPIN         AS LONG
END TYPE

TYPE PayListType
  BillRec       AS LONG      'Pointer to bill trans rec this payment is for
  Principle1    AS DOUBLE
  Interest1     AS DOUBLE
  Collection    AS DOUBLE
  TotPaid       AS DOUBLE    'amount paid to this bill rec (parital payment)
  CustRec       AS LONG      'backup pointer to cust rec
  PrevListRec   AS LONG      'pointer to next paylist rec
END TYPE

