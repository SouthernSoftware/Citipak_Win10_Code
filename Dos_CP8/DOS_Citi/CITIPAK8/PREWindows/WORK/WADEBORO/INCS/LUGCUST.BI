TYPE LugOffWCustType
 CAcct           AS STRING * 6
 LastName        AS STRING * 32
 FrstName        AS STRING * 22
 Addr1           AS STRING * 22
 City            AS STRING * 16
 State           AS STRING * 2
 Zip             AS STRING * 10
 ServAddr        AS STRING * 22
 MrtNumber       AS STRING * 16
 PrevBal         AS STRING * 8    'format: #####.##  t9
 BalDue          AS STRING * 8    'format: #####.##  t8
 WatCharge       AS STRING * 8    'format: #####.##  t8
 PayThisMth      AS STRING * 8    'format: #####.##  t8
 CredThisMth     AS STRING * 8    'format: #####.##  t8
 LateCharge      AS STRING * 8    'format: #####.##  t8
 Payment         AS STRING * 8    'format: #####.##  t8
 Credit          AS STRING * 8    'format: #####.##  t8
 LPayAmt         AS STRING * 8    'format: #####.##  t8
 LPayDate        AS STRING * 8    'format: ##/##/##  t8
'good b229
 LCredAmt        AS STRING * 8    'format: #####.##  t8
 LCredDate       AS STRING * 8    'format: ##/##/##  t8
 YTDPayAmt       AS STRING * 8    'format: #####.##  t8
 YTDCharge       AS STRING * 8    'format: #####.##  t8
 YTDCredit       AS STRING * 8    'format: #####.##  t8
 TapCharge       AS STRING * 8    'format: #####.##  t8
 DepAmt          AS STRING * 8    'format: #####.##  t8
 Escrow          AS STRING * 8    'format: #####.##  t8
 PrevUse         AS STRING * 8    'format: #####.##  t8
 ImpactFee       AS STRING * 8    'format: #####.##  t8
'Good b309
 EstRead         AS STRING * 8    'format: ########  t8
 PrevRead        AS STRING * 8    'format: ########  t8
 CurrRead        AS STRING * 8    'format: ########  t8
 ServType        AS STRING * 2    'format: ##  t2 B333 & b334 all "1 or 0"
'Good b335
 NumbUnits       AS STRING * 3    'format: ### t3             all "1"
 Phone           AS STRING * 14   'format: #-###-###-####  t14
 MtrReadDate     AS STRING * 8    'format: ##/##/##  t8
 Comment         AS STRING * 10   'Desc says "Comment - Estimate meter."
 CurAmtDue       AS STRING * 8    'format: #####.##  t8
 RaceCode        AS STRING * 1    'format: "?"
 TypeCode        AS STRING * 1    'format: "?"
 DaysOver30      AS STRING * 8    'format: #####.##  t8
 DaysOver60      AS STRING * 8    'format: #####.##  t8
 DaysOver90      AS STRING * 8    'format: #####.##  t8
 LugOrElgin      AS STRING * 1    'format: "?" L or E
 RetChkChrg      AS STRING * 8    'format: #####.##  t8
 ReconFee        AS STRING * 8    'format: #####.##  t8
 SewCharge       AS STRING * 8    'format: #####.##  t8
 DateLastRead    AS STRING * 8    'format: ##/##/##  t8
 TimeLastRead    AS STRING * 8    'format: ##:##:##  t8
 LastReadDate    AS STRING * 8    'format: #####.##  t8
 LastReadTime    AS STRING * 8    'format: ##:##:##  t8
 BadChkCnt       AS STRING * 2    'format: ##        t2
 NoOfHolds       AS STRING * 1    'format: #
 CrLf            AS STRING * 2
END TYPE

TYPE LugOffSCustType
 CAcct           AS STRING * 6
 LastName        AS STRING * 32
 FrstName        AS STRING * 22
 Addr1           AS STRING * 22
 City            AS STRING * 16
 State           AS STRING * 2
 Zip             AS STRING * 10
 ServAddr        AS STRING * 22
 MrtNumber       AS STRING * 10
 PrevBal         AS STRING * 8    'format: #####.##  t9
 BalDue          AS STRING * 8    'format: #####.##  t8
 WatCharge       AS STRING * 8    'format: #####.##  t8
 PayThisMth      AS STRING * 8    'format: #####.##  t8
 CredThisMth     AS STRING * 8    'format: #####.##  t8
 LateCharge      AS STRING * 8    'format: #####.##  t8
 Payment         AS STRING * 8    'format: #####.##  t8
 Credit          AS STRING * 8    'format: #####.##  t8
 LPayAmt         AS STRING * 8    'format: #####.##  t8
 LPayDate        AS STRING * 8    'format: ##/##/##  t8
'good b229
 LCredAmt        AS STRING * 8    'format: #####.##  t8
 LCredDate       AS STRING * 8    'format: ##/##/##  t8
 YTDPayAmt       AS STRING * 8    'format: #####.##  t8
 YTDCharge       AS STRING * 8    'format: #####.##  t8
 YTDCredit       AS STRING * 8    'format: #####.##  t8
 TapCharge       AS STRING * 8    'format: #####.##  t8
 DepAmt          AS STRING * 8    'format: #####.##  t8
 Escrow          AS STRING * 8    'format: #####.##  t8
 PrevUse         AS STRING * 8    'format: #####.##  t8
 ImpactFee       AS STRING * 8    'format: #####.##  t8
'Good b309
' CurrRead        AS STRING * 8    'format: ########  t8
' PrevRead        AS STRING * 8    'format: ########  t8
' EstRead         AS STRING * 8    'format: ########  t8
 EstRead         AS STRING * 8    'format: ########  t8
 PrevRead        AS STRING * 8    'format: ########  t8
 CurrRead        AS STRING * 8    'format: ########  t8
 ServType        AS STRING * 2    'format: ##  t2 B333 & b334 all "1 or 0"
'Good b335
 NumbUnits       AS STRING * 3    'format: ### t3             all "1"
 Phone           AS STRING * 14   'format: #-###-###-####  t14
 ReservFee       AS STRING * 8    'format: #####.##  t8
 AppFee          AS STRING * 8    'format: #####.##  t8
 InspFee         AS STRING * 8    'format: #####.##  t8
 MtrReadDate     AS STRING * 8    'format: ##/##/##  t8
 LastReadDate    AS STRING * 8    'format: #####.##  t8
 LastReadTime    AS STRING * 8    'format: ##:##:##  t8
 BadChkCnt       AS STRING * 2    'format: ##        t2
 NoOfHolds       AS STRING * 1    'format: #
 CrLf            AS STRING * 2
END TYPE

