'This is new A/P Ledger record type

TYPE APLedgerRecType             '128 bytes
   VIN         AS INTEGER        '2 AS vennum$  our pin
   VendorCode  AS STRING * 10    'Do we need
   VRecNum     AS INTEGER        'vendors rec num
   InvDate     AS INTEGER        '8 AS INVDATE$
   GLDistDate  AS INTEGER
   DueDate     AS INTEGER
   TrCode      AS INTEGER        '2 AS PAYCODE$
   InvNum      AS STRING * 25    '25 AS invnum$
   PONum       AS STRING * 20    '10 AS ponum$
   MiscCode    AS STRING * 33    'Future needs
   DrAmt       AS DOUBLE         '8 AS amount$
   CrAmt       AS DOUBLE
   FrstDist    AS LONG           'pointer to first dist in apdist
   lastDist    AS LONG           'pointer to last dist in apdist
   NextTrans   AS LONG           'pointer to next trans for this vendor
END TYPE

TYPE APDistRecType               '42 bytes
   APLedgerRec AS LONG           'Points back to apledger rec
   DistAcctRec AS INTEGER        'pointer to account rec
   DistAcctNum AS STRING * 16    'account number string needed for re-link
   DistDrAmt   AS DOUBLE         'Debit Amt
   DistCrAmt   AS DOUBLE         'Credit Amt
   NextDist    AS LONG           'Next Distribution
END TYPE

