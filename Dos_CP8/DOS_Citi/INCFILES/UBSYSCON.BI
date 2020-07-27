TYPE SysconHHType
  FirstName       AS STRING * 20
  LastName        AS STRING * 20
  TreadCode       AS STRING * 1
  Addr            AS STRING * 20
  RDate           AS STRING * 6
  RTime           AS STRING * 6
  AvgUse          AS STRING * 8
  CurrRead        AS STRING * 8
  PrevRead        AS STRING * 8
  AcctNo          AS STRING * 8
  Serial          AS STRING * 10
  HighRead        AS STRING * 8
  LowRead         AS STRING * 8
  Fill            AS STRING * 122
  X               AS STRING * 1
  CrLf            AS STRING * 2
END TYPE

'HH inout structure
' 20 AS trfirstname$
' 20 AS trlname$
' 1 AS treadcode$
' 20 AS tradd1$
' 6 AS trdate$
' 6 AS trtime$
' 8 AS travguse$
' 8 AS trcurread$
' 8 AS trpastread$
' 8 AS tracctnum$
' 10 AS trserial$
' 8 AS trhighread$
' 8 AS trlowread$
' 123 AS trunused$
' 2 AS crlf$


