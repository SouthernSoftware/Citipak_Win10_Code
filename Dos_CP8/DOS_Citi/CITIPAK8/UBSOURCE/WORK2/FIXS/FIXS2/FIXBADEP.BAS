DEFINT A-Z
DECLARE FUNCTION FileSize& (FileName$)
  
  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubtrans.BI'

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  CLS
  PRINT
  PRINT
  PRINT
  PRINT
  LINE INPUT ; "Enter account number: "; Cust$
  WhatCust = VAL(Cust$)
  IF WhatCust = 0 THEN
    CLS
    END
  END IF

  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  IF WhatCust > NumOfRecs& THEN END

  GET UBFile, WhatCust, UBCustRec(1)
  IF UBCustRec(1).DepositAmt > 0 THEN
    Dep# = UBCustRec(1).DepositAmt
    GOSUB FindTrans
    UBCustRec(1).DepositAmt = 0
    PUT UBFile, WhatCust, UBCustRec(1)
    PRINT : PRINT : PRINT
    PRINT USING "Removed deposit: ####.## "; Dep#
    PRINT
    PRINT "Press any key. . ."

  ELSE
    PRINT : PRINT : PRINT
    PRINT "Customer has no deposit!"
    PRINT
    PRINT "Press any key. . ."
  END IF
  CLOSE UBFile
  a$ = INPUT$(1)
  RUN

FindTrans:
 REDIM Trans(1) AS UBTransRecType
 TransLen = LEN(Trans(1))
 UBTran = FREEFILE
 OPEN "UBtrans.dat" FOR RANDOM SHARED AS UBTran LEN = TransLen

 LastTran& = UBCustRec(1).LastTrans
 DO WHILE LastTran& > 0
   GET UBTran, LastTran&, Trans(1)

   IF (Trans(1).TransType = TranDepositPayment) OR (Trans(1).TransType = (TranDepositPayment + 100)) THEN

     Trans(1).TransDesc = "REFUNDED"
     Trans(1).TransType = -TranDepositPayment
     Trans(1).OperatorNumber = -100
     PUT UBTran, LastTran&, Trans(1)
     EXIT DO
   ELSE
     LastTran& = Trans(1).PrevTrans
   END IF
 LOOP
 CLOSE UBTran

RETURN

