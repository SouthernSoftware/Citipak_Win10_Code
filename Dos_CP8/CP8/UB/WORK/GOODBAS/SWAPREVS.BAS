DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubtrans.BI'

  ''$INCLUDE: 'ubrevswp.BI'

  CONST False = 0, True = NOT False

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  REDIM UBCustRec(1)  AS NewUBCustRecType
  UBCustLen = LEN(UBCustRec(1))

  CLS
  PRINT "Processing Customers. . ."
  
  UBFile = FREEFILE
  OPEN "ubcust.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustLen
  NumOfRec& = LOF(UBFile) \ UBCustLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 4, 1
    PRINT "Processing:"; Cnt&; " of"; NumOfRec&;
    GET UBFile, Cnt&, UBCustRec(1)
    SWAP UBCustRec(1).CurrRevAmts(3), UBCustRec(1).CurrRevAmts(14)
    SWAP UBCustRec(1).PrevRevAmts(3), UBCustRec(1).PrevRevAmts(14)
    SWAP UBCustRec(1).Serv(3), UBCustRec(1).Serv(14)
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT "Customer Processing Complete."
  PRINT
  PRINT
  PRINT
  PRINT "Processing Transactions. . ."
  UBTFile = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTFile LEN = UBTranRecLen
  TNumOfRecs& = LOF(UBTFile) / UBTranRecLen
  FOR Cnt& = 1 TO TNumOfRecs&
    LOCATE 12, 1: PRINT "Processing:"; Cnt&; "of"; TNumOfRecs&;
    GET UBTFile, Cnt&, UBTranRec(1)
    SWAP UBTranRec(1).RevAmt(3), UBTranRec(1).RevAmt(14)
    PUT UBTFile, Cnt&, UBTranRec(1)
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT "Transaction Processing Complete."
  PRINT
  PRINT
  PRINT "Done."
  CALL KillFile("swaprevs.exe")


