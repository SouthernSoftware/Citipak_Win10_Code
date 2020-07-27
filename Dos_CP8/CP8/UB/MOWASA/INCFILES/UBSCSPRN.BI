
TYPE BillOutServType
   ServText               AS STRING * 20
   ServAmt                AS STRING * 10
END TYPE

TYPE BillOutRecType   '
   AcctNo                 AS STRING * 8
   LocationNum            AS STRING * 9        'format ##-######
   CUSTNAME               AS STRING * 35
   ADDR1                  AS STRING * 35
   ADDR2                  AS STRING * 35
   SERVADDR               AS STRING * 35
   CITY                   AS STRING * 18
   STATE                  AS STRING * 2
   ZIPCODE                AS STRING * 10
   BillType               AS STRING * 1        'N=Normal F=Final
   DepAppAmt              AS STRING * 10       'if final applied deposit amt.
   PrevDue                AS STRING * 15       'total revenue amount
   CurrDue                AS STRING * 15       'total revenue amount
   TotalDue               AS STRING * 15       'total revenue amount
   MtrType                AS STRING * 1        'w=water s=sewer c=combined
   MtrUnit                AS STRING * 1
   CurrDate               AS STRING * 10
   PrevDate               AS STRING * 10
   ServDays               AS STRING * 4
   CurrRead               AS STRING * 10       'Last/Current meter readings
   PrevRead               AS STRING * 10       'Previous readings
   Consump                AS STRING * 10
   IServDays              AS STRING * 4
   ICurrRead              AS STRING * 10       'Last/Current meter readings
   IPrevRead              AS STRING * 10       'Previous readings
   IConsump               AS STRING * 10
   ServInfo(1 TO 15)      AS BillOutServType
   BillDate               AS STRING * 10
   PastDueDate            AS STRING * 10
   DraftDate              AS STRING * 10
'mod for Johnston CO.
   MsgLine1               AS STRING * 22     'was 25
   MsgLine2               AS STRING * 22     'was 25
   MsgLine3               AS STRING * 22     'was 25
   MsgLine4               AS STRING * 22     'was 25
   MtrNumb                AS STRING * 12     'Added Meter Number
   CrLf                   AS STRING * 2
END TYPE


TYPE MowBillOutRecType   'Mow
   AcctNo                 AS STRING * 8
   LocationNum            AS STRING * 9        'format ##-######
   CUSTNAME               AS STRING * 35
   ADDR1                  AS STRING * 35
   ADDR2                  AS STRING * 35
   SERVADDR               AS STRING * 35
   CITY                   AS STRING * 18
   STATE                  AS STRING * 2
   ZIPCODE                AS STRING * 10
   BillType               AS STRING * 1        'N=Normal F=Final
   DepAppAmt              AS STRING * 10       'if final applied deposit amt.
   PrevDue                AS STRING * 15       'total revenue amount
   CurrDue                AS STRING * 15       'total revenue amount
   TotalDue               AS STRING * 15       'total revenue amount
   MtrType                AS STRING * 1        'w=water s=sewer c=combined
   MtrUnit                AS STRING * 1
   CurrDate               AS STRING * 10
   PrevDate               AS STRING * 10
   ServDays               AS STRING * 4
   CurrRead               AS STRING * 10       'Last/Current meter readings
   PrevRead               AS STRING * 10       'Previous readings
   Consump                AS STRING * 10
   IServDays              AS STRING * 4
   ICurrRead              AS STRING * 10       'Last/Current meter readings
   IPrevRead              AS STRING * 10       'Previous readings
   IConsump               AS STRING * 10

   ServInfo(1 TO 15)      AS BillOutServType
   'this is the structure above

   BillDate               AS STRING * 10
   PastDueDate            AS STRING * 10
   DraftDate              AS STRING * 10
   MsgLine1               AS STRING * 25     'was 25
   MsgLine2               AS STRING * 25     'was 25
   MsgLine3               AS STRING * 25     'was 25
   MsgLine4               AS STRING * 25     'was 25
   CrLf                   AS STRING * 2
END TYPE


