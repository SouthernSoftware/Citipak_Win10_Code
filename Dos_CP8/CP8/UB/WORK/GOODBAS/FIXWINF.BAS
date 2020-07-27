DEFINT A-Z


  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubTrans.BI'

  CLS
  PRINT "Processing. . ."
  PRINT
  PRINT
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  FOR RecNo& = 1 TO LOF(UBFile) \ UBCustRecLen
    GET UBFile, RecNo&, UBCustRec(1)
    UBCustRec(1).EstFlag = ""
    PUT UBFile, RecNo&, UBCustRec(1)
  NEXT
  CLOSE

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  FOR ThisTrans& = 1 TO LOF(UBTran) \ UBTranRecLen
    GET #UBTran, ThisTrans&, UBTranRec(1)
    FOR EstCnt = 1 TO 7
      UBTranRec(1).EstRead(EstCnt) = ""
    NEXT
    PUT #UBTran, ThisTrans&, UBTranRec(1)
  NEXT
  CLOSE

  PRINT "Done."
  CALL KillFile("fixwinf.exe")

