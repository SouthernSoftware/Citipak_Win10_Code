DECLARE FUNCTION uRound# (N#)
DEFINT A-Z

  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'

  CLS

  REDIM UBCust(1)    AS NewUBCustRecType
  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))
  CustLen = LEN(UBCust(1))

  UBFile1 = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBFile1 LEN = UBTranRecLen

  UBFile3 = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile3 LEN = CustLen

  LOCATE 5, 1
  PRINT "Adjusting Customer 332"
  GET UBFile3, 332, UBCust(1)
  GET UBFile1, UBCust(1).LastTrans, UBTranRec(1)
  UBTranRec(1).RunBalance = 0
  FOR zz = 1 TO 15
    UBCust(1).CurrRevAmts(zz) = 0
  NEXT
  UBCust(1).CurrBalance = 0
  UBCust(1).PrevBalance = 0
  PUT UBFile3, 332, UBCust(1)
  PUT UBFile1, UBCust(1).LastTrans, UBTranRec(1)

  CLOSE

  LOCATE 9, 1
  PRINT "Done."

FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

