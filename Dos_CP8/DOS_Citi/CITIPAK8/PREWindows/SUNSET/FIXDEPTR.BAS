DEFINT A-Z

  '$INCLUDE: 'UBTRANS.bi'
  '$INCLUDE: 'newcust.bi'
  CLS
  PRINT "Correcting Deposit Transaction. . ."
  
  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  GET UBTran, 29002, UBTranRec(1)
  UBTranRec(1).RevAmt(1) = 60
  PUT UBTran, 29002, UBTranRec(1)
  GET UBTran, 29210, UBTranRec(1)
  UBTranRec(1).RunBalance = 0
  PUT UBTran, 29210, UBTranRec(1)
  CLOSE UBTran


  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, 668, UBCustRec(1)
  UBCustRec(1).CurrBalance = -23.05
  PUT UBCust, 668, UBCustRec(1)

  CLOSE

  PRINT
  PRINT "Completed."
  PRINT
  PRINT "Correcting Invalid Adjustment Transactions & Account Balances"

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  GET UBCust, 683, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 683, UBCustRec(1)

  GET UBCust, 5, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 5, UBCustRec(1)

  GET UBCust, 5, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 5, UBCustRec(1)


CLOSE
END

FixRevs:
  FOR zz = 1 TO 15
    UBCustRec(1).CurrRevAmts(zz) = 0
    UBCustRec(1).PrevRevAmts(zz) = 0
  NEXT
  UBCustRec(1).CurrBalance = 0
  UBCustRec(1).PrevBalance = 0
RETURN

