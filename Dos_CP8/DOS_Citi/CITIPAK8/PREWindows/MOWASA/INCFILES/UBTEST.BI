'020299


TYPE TestUBCustRecType
    Fill1 AS STRING * 1053

    'LastTrans     AS LONG
END TYPE

TYPE TestUBTransRecType
   'Fill1 AS STRING * 319
   'CustAcctNo             AS LONG         'Pointer to RecNo in ubcust.dat
   'Cust bytes offset = 320-323   4 byte long integer
   Fill1 AS STRING * 323
   'PrevTrans              AS LONG
   'Prev Bytes offset = 324-327   4 byte long integer
END TYPE


