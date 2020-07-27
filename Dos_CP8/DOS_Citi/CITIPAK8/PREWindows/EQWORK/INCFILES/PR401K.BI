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

