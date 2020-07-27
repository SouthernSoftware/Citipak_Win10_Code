TYPE UBTransRecType
   TransDate              AS INTEGER      '
   TransType              AS INTEGER      '1=Charge : 2=Payment
   TransDesc              AS STRING * 25  'may change
   TransAmt               AS DOUBLE       'total revenue amount
   RevAmt(1 TO 10)        AS DOUBLE       'Revenue amounts
   CustRate(1 TO 10)      AS STRING * 4   'Records Rate Code at Time of Transaction
   CurRead(1 TO 7)        AS LONG         'Last/Current meter readings
   PrevRead(1 TO 7)       AS LONG         'Previous readings
   EstRead(1 TO 7)        AS STRING * 1   'Y/N Flags for meter est's
   CustLocation           AS LONG         'Pointer to Location RecNo
   CustStatus             AS STRING * 1   'Customer Status at Time of Transaction
   CustCode               AS STRING * 2   'Customer Code at Time of Transaction
   PostRoute              AS STRING * 4   'Postal Route at Time of Transaction
   BillCycle              AS STRING * 2   'Customer Billing Cycle at Time of Transaction
   PayTypeCode            AS STRING * 1   'Payment Type:  1=Cash, 2=Check
   OperatorNumber         AS INTEGER      '
   CustAcctNo             AS LONG         'Pointer to RecNo in ubcust.dat
   NextTrans              AS LONG
   ActiveFlag             AS INTEGER      'Valid transaction flag
   Unused AS STRING * 10
END TYPE


TYPE oUBTransRecType
   TransDate              AS INTEGER
   TransType              AS INTEGER     '1=Charge : 2=Payment
   TransDesc              AS STRING * 25
   TransAmt               AS DOUBLE
   RunBalance             AS DOUBLE
   RevAmt(1 TO 10)        AS DOUBLE
   CustRate(1 TO 10)      AS STRING * 2  'Records Rate Code at Time of Transaction
   CurRead(1 TO 7)        AS DOUBLE
   PastRead(1 TO 7)       AS DOUBLE
   Consumption(1 TO 7)    AS DOUBLE
   EstRead(1 TO 7)        AS STRING * 1
   CustLocation           AS STRING * 9  ' Records Customer Location at Time of Transaction
   CustStatus             AS STRING * 1  ' Records Customer Status at Time of Transaction
   CustCode               AS STRING * 3    ' Records Customer Code at Time of Transaction
   CustZone               AS STRING * 3    ' Records Customer Zone at Time of Transaction
   BillCycle              AS STRING * 2    ' Records Customer Billing Cycle at Time of Transaction
   PayTypeCode            AS INTEGER       ' Records Type of Payment Cash : Check
   OperatorNumber         AS INTEGER
   CustAcctNo             AS INTEGER     'Equals Record Number in ubcust.dat
   PrevTrans              AS LONG
   LastTrans              AS LONG
   Unused AS STRING * 161
END TYPE

