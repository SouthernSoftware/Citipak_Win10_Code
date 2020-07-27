TYPE HeaderRecType
  ID          AS STRING * 5
  Fill1       AS STRING * 1
  PCN1        AS STRING * 5
  Fill2       AS STRING * 1   'Spaces
  Fill3       AS STRING * 2   'Zeros
  PCN2        AS STRING * 5
  DataType    AS STRING * 1   '"E"
  Fill4       AS STRING * 8   'spaces
  FileDate    AS STRING * 6   'mmddyy
  Fill5       AS STRING * 3
  PayrollName AS STRING * 32
  Fill6       AS STRING * 11
  CrLf        AS STRING * 2
END TYPE


TYPE DetailRecType
  ID          AS STRING * 1
  Batch       AS STRING * 5
  PCN         AS STRING * 5
  ProcDate    AS STRING * 6
  SSN         AS STRING * 9
  EmpName     AS STRING * 30
  EmpVolDed   AS STRING * 8
  EmpLoanPay  AS STRING * 8
  EmpContAmt  AS STRING * 8
  CrLf        AS STRING * 2
END TYPE

TYPE TrailerRecType
  ID          AS STRING * 1
  TotVolDED   AS STRING * 11
  TotLoanAmt  AS STRING * 11
  TotContAmt  AS STRING * 11
  TotDRecs    AS STRING * 6
  Filler      AS STRING * 40
  CrLf        AS STRING * 2
END TYPE

