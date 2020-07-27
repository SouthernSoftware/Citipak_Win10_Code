TYPE EPPTransRecType
   TransDate              AS INTEGER      '
  'TransType              AS INTEGER      '
  'TransAmt               AS DOUBLE       'total revenue amount
   CustAcctNo             AS LONG         'Pointer to customer rec number
   BillTrans              AS LONG         'Pointer to bill trans rec number
   PrevEPPTran            AS LONG         'Pointer to prev EPP Trans
   EPPAmt(1 TO 15)        AS DOUBLE       'Equal pay amounts
   BillAmt(1 TO 15)       AS SINGLE       'Actual Bill Amounts
END TYPE

'020199 Added Equal pay plan types
TYPE EPInfoType
    RevDis      AS DOUBLE
    RevBal      AS DOUBLE
END TYPE

TYPE UBCustEPPRecType
    EPInfo(1 TO 15) AS EPInfoType
    EPDisTot        AS DOUBLE
    EPBalTot        AS DOUBLE
    ChkByte         AS STRING * 1
END TYPE
'*********************************

