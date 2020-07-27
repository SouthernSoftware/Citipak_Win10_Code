
' OPEN "R", 1, "wbcust.dat", 256

TYPE WBSpruceCust0RecType
 LastName       AS STRING * 25     '25 AS lname$
 FirstName      AS STRING * 25     '25 AS firstname$
 Add1           AS STRING * 30     '30 AS ADD1$
 Add2           AS STRING * 30     '30 AS add2$
 City           AS STRING * 20     '20 AS CITY$
 State          AS STRING * 2      '2 AS ST$
 Zip            AS STRING * 5      '5 AS ZP$
 Phone          AS STRING * 12     '12 AS PHONE$

 Acct           AS SINGLE          '4 AS acct$
 Book           AS INTEGER         '2 AS BOOK$
 Meter          AS STRING * 12     '12 AS meter$
 Rate           AS INTEGER         '2 AS RATE$
 WON            AS STRING * 8      '8 AS WON$   Inside or Outside
 MInstall       AS STRING * 8      '8 AS MINSTALL$ Date meter installed
 AvgUse         AS DOUBLE          '8 AS AVGUSE$
 AvgCT          AS INTEGER         '2 AS AVGCT$
 CurBal         AS DOUBLE          '8 AS curbal$
 PastDue        AS DOUBLE          '8 AS pastdue$  past due amt!!!
 Cycle          AS STRING * 1      '1 AS cycle$
 Fill1          AS STRING * 5      '5 AS NOTUSED$
 CurRead        AS DOUBLE          '8 AS CURREAD$
 PastRead       AS DOUBLE          '8 AS PASTREAD$
 FirstTR        AS DOUBLE          '8 AS firsttr$
 LastTR         AS DOUBLE          '8 AS lasttr$
 WaterDEP       AS SINGLE          '4 AS waterdep$
 Fill2          AS STRING * 3
END TYPE

TYPE WBSpruceCust1RecType
  RAMT1        AS SINGLE           '  4
  RAMT2        AS SINGLE           '  4
  RAMT3        AS SINGLE           '  4
  RDESC1       AS STRING * 25      '  25
  RDESC2       AS STRING * 25      '  25
  RDESC3       AS STRING * 25      '  25
  RACCT1       AS STRING * 9       '  9
  RACCT2       AS STRING * 9       '  9
  RACCT3       AS STRING * 9       '  9
  OAMT1        AS SINGLE           '  4
  OAMT2        AS SINGLE           '  4
  ODESC1       AS STRING * 25      '  25
  ODESC2       AS STRING * 25      '  25
  OACCT1       AS STRING * 9       '  9
  OACCT2       AS STRING * 9       '  9
  MINIMUM      AS INTEGER          '  2
  CUSTTYPE     AS INTEGER          '  2
  METERNUM     AS STRING * 14      '  14
  add3         AS STRING * 26      '  26
  carrier      AS STRING * 4       '  4
  flow         AS STRING * 1       '  1
  flowcode     AS STRING * 1       '  1
  multacctS    AS INTEGER          '  2
  masteracct   AS STRING * 1       '  1
  slavecode    AS STRING * 1       '  1
  readmeter    AS DOUBLE           '  8
  Fill         as string * 4
END TYPE


