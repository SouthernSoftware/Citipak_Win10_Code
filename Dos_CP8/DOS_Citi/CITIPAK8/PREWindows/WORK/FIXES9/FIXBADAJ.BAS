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
  PRINT "Correcting Invalid Adjustment to Account Balances"

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen


  GET UBCust, 5, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 5, UBCustRec(1)

  GET UBCust, 92, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 92, UBCustRec(1)

  GET UBCust, 172, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 172, UBCustRec(1)

  GET UBCust, 260, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 260, UBCustRec(1)

  GET UBCust, 261, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 261, UBCustRec(1)

  GET UBCust, 273, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 273, UBCustRec(1)

  GET UBCust, 341, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 341, UBCustRec(1)

  GET UBCust, 346, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 346, UBCustRec(1)

  GET UBCust, 373, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 373, UBCustRec(1)

  GET UBCust, 383, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 383, UBCustRec(1)

  GET UBCust, 390, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 390, UBCustRec(1)

  GET UBCust, 396, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 396, UBCustRec(1)

  GET UBCust, 418, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 418, UBCustRec(1)

  GET UBCust, 439, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 439, UBCustRec(1)

  GET UBCust, 464, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 464, UBCustRec(1)

  GET UBCust, 485, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 485, UBCustRec(1)

  GET UBCust, 491, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 491, UBCustRec(1)

  GET UBCust, 564, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 564, UBCustRec(1)

  GET UBCust, 573, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 573, UBCustRec(1)

  GET UBCust, 600, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 600, UBCustRec(1)

  GET UBCust, 607, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 607, UBCustRec(1)

  GET UBCust, 627, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 627, UBCustRec(1)

  GET UBCust, 633, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 633, UBCustRec(1)

  GET UBCust, 657, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 657, UBCustRec(1)

  GET UBCust, 661, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 661, UBCustRec(1)

  GET UBCust, 666, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 666, UBCustRec(1)

  GET UBCust, 669, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 669, UBCustRec(1)

  GET UBCust, 672, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 672, UBCustRec(1)

  GET UBCust, 674, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 674, UBCustRec(1)

  GET UBCust, 675, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 675, UBCustRec(1)

  GET UBCust, 680, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 680, UBCustRec(1)

  GET UBCust, 683, UBCustRec(1)
  GOSUB FixRevs
  PUT UBCust, 683, UBCustRec(1)

CLOSE
PRINT
PRINT
PRINT "Processing Complete"

KILL "FIXBADAJ.EXE"

END

FixRevs:
  FOR zz = 1 TO 15
    UBCustRec(1).CurrRevAmts(zz) = 0
    UBCustRec(1).PrevRevAmts(zz) = 0
  NEXT
  UBCustRec(1).CurrBalance = 0
  UBCustRec(1).PrevBalance = 0
RETURN

