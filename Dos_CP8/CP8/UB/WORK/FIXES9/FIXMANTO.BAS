DEFINT A-Z


  '$INCLUDE: 'newcust.bi'
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS #1 LEN = UBCustRecLen
  GET #1, 430, UBCustRec(1)
  UBCustRec(1).DepositAmt = 100
  PUT #1, 430, UBCustRec(1)

  CLOSE

