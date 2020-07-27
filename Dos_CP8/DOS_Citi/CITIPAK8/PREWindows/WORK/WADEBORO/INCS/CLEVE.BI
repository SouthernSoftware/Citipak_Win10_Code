TYPE WB8Cust1RecType
  RAmt1    AS SINGLE       ' 4 AS RAMT1$
  RAmt2    AS SINGLE       ' 4 AS RAMT2$
  RAmt3    AS SINGLE       ' 4 AS RAMT3$
  RDecs1   AS STRING * 25  '25 AS RDESC1$
  RDecs2   AS STRING * 25  '25 AS RDESC2$
  RDecs3   AS STRING * 25  '25 AS RDESC3$
  RAcct1   AS STRING * 9   ' 9 AS RACCT1$
  RAcct2   AS STRING * 9   ' 9 AS RACCT2$
  RAcct3   AS STRING * 9   ' 9 AS RACCT3$
  OAmt1    AS SINGLE       ' 4 AS OAMT1$
  OAmt2    AS SINGLE       ' 4 AS OAMT2$
  ODecs1   AS STRING * 25  '25 AS ODESC1$
  ODecs2   AS STRING * 25  '25 AS ODESC2$
  OAcct1   AS STRING * 9   ' 9 AS OACCT1$
  OAcct2   AS STRING * 9   ' 9 AS OACCT2$
  Fill1    AS STRING * 66  '66 AS blank$
END TYPE

TYPE WB8CustRecType
  LastName       AS STRING * 25   '25 AS lname$
  FirstName      AS STRING * 25   '25 AS firstname$
  Add1           AS STRING * 30   '30 AS ADD1$
  Add2           AS STRING * 30   '30 AS ADD2$
  City           AS STRING * 20   '20 AS CITY$
  State          AS STRING * 2    '2 AS ST$
  Zip            AS STRING * 9    '9 AS ZP$
  Phone          AS STRING * 12   '12 AS phone$

  pad1           AS INTEGER       '2 AS BOOK$
  Meter          AS STRING * 12   '12 AS meter$
  WON            AS STRING * 8    '8 AS WON$
  MinStall       AS DOUBLE        '8 AS MINSTALL$
  AvgUse         AS DOUBLE        '8 AS AVGUSE$
  AvgCnt         AS INTEGER       '2 AS AVGCT$
  CurBal         AS DOUBLE        '8 AS curbal$
  PastDue        AS DOUBLE        '8 AS pastdue$
  Cycle          AS STRING * 1    '1 AS cycle$
  CurrRead       AS DOUBLE        '8 AS curread$
  PastRead       AS DOUBLE        '8 AS pastread$
  FirstTrans     AS DOUBLE        '8 AS firsttr$
  LastTrans      AS DOUBLE        '8 AS LASTTR$
  WaterDep       AS SINGLE        '4 AS waterdep$
  IRev1          AS DOUBLE        '8 AS IRev1$
  IRev2          AS DOUBLE        '8 AS IRev2$
  IRev3          AS DOUBLE        '8 AS IRev3$
  IRev4          AS DOUBLE        '8 AS IRev4$
  IRev5          AS DOUBLE        '8 AS IRev5$
  IRev6          AS DOUBLE        '8 AS IRev6$
  IRev7          AS DOUBLE        '8 AS irev7$
  IRev8          AS DOUBLE        '8 AS IREV8$

  CustType       AS STRING * 2    '2 AS CUSTTYPE$
  MeterNum       AS STRING * 14   '14 AS METERNUM$
  Addr3          AS STRING * 30   '30 AS ADD3$
  Minimums       AS INTEGER       '2 AS minimums$
  rate1          AS STRING * 2    '2 AS rate1$
  rate2          AS STRING * 2    '2 AS rate2$
  rate3          AS STRING * 2    '2 AS rate3$
  rate4          AS STRING * 2    '2 AS rate4$
  rate5          AS STRING * 2    '2 AS rate5$
  rate6          AS STRING * 2    '2 AS rate6$
  rate7          AS STRING * 2    '2 AS rate7$
  rate8          AS STRING * 2    '2 AS rate8$
  MeterReadDate  AS STRING * 6    '6 AS METERREADDATE$
  MeterEST       AS STRING * 1    '1 AS METEREST$
  WorkOrder      AS STRING * 1    '1 AS workorder$
  WorkOrderNum   AS SINGLE        '4 AS workordernumber$
  WorkOrderBDate AS STRING * 6    '6 AS workorderbeg$
  WorkOrderEDate AS STRING * 6    '6 AS workorderend$
  MessFirst      AS SINGLE        '4 AS messfirst$
  MessLast       AS SINGLE        '4 AS messlast$
  Comment        AS STRING * 30   '30 AS comment$
  SeqNum         AS SINGLE        '4 AS seqnum$
  CustID         AS STRING * 10   '10 AS CUSTID$
  ProRate        AS STRING * 1    '1 AS prorate$
  DepDate        AS STRING * 6    '6 AS depdate$
  FirstTwo       AS SINGLE        '4 AS firstwo$
  LastTwo        AS SINGLE        '4 AS lastwo$
  MtrTime        AS STRING * 6    '6 AS mtrtime$
  Pad2           AS STRING * 41   '41 AS notused$
END TYPE

TYPE WB8TransRecType
  TransDate      AS STRING * 8    '8 AS TDATE$
  TransAmt       AS DOUBLE        ' 8 AS tamt$
  TransType      AS INTEGER       ' 2 AS TTYPE$
  TransDesc      AS STRING * 15   ' 15 AS TDESC$
  CustRec        AS SINGLE        ' 4 AS tmrec$
  NextTrans      AS SINGLE        ' 4 AS tnexttr$
  TCurrAmt       AS SINGLE        ' 4 AS tcuramt$
  RevAmt1        AS SINGLE        ' 4 AS wbrev1$
  RevAmt2        AS SINGLE        ' 4 AS wbrev2$
  RevAmt3        AS SINGLE        ' 4 AS wbrev3$
  RevAmt4        AS SINGLE        ' 4 AS wbrev4$
  RevAmt5        AS SINGLE        ' 4 AS wbrev5$
  RevAmt6        AS SINGLE        ' 4 AS wbrev6$
  RevAmt7        AS SINGLE        ' 4 AS wbrev7$
  RevAmt8        AS SINGLE        ' 4 AS wbrev8$
  TEstimate      AS STRING * 1    ' 1 AS TESTIMATE$
  TCurrRead      AS DOUBLE        ' 8 AS TCURREAD$
  TPastRead      AS DOUBLE        ' 8 AS TPASTREAD$
  TConsump       AS DOUBLE        ' 8 AS TCONSUMP$
  TBook          AS INTEGER       ' 2 AS TBOOK$
  TCustType      AS INTEGER       ' 2 AS TCUSTTYPE$
  TCustLoc       AS STRING * 1    ' 1 AS TCUSTLOC$
  TStatus        AS STRING * 1    ' 1 AS TSTATUS$
  FillPad        AS STRING * 20   ' 20 AS wbblank$
END TYPE

