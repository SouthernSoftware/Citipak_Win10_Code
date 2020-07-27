TYPE K401RptType
  EmpName AS STRING * 35
  SSN     AS STRING * 11
  VAmt    AS STRING * 11
  LAmt    AS STRING * 11
  MAmt    AS STRING * 11
  Batch   AS STRING * 11
  HDate   AS STRING * 10
END TYPE

TYPE DedRptType
  EmpName AS STRING * 35
  SSN     AS STRING * 11
  DAmt    AS STRING * 23
END TYPE

TYPE EmpTransHistRptType
  EmpName        AS STRING * 35
  SSN            AS STRING * 11
  CheckNum       AS LONG
  CheckDate      AS INTEGER   'yea
  GrossPay       AS DOUBLE    'this is add earnings and wages combined
  'GrossWage      AS DOUBLE   'this is also in the transaction history
  TotDedAmt      AS DOUBLE
  NetPay         AS DOUBLE
end type
