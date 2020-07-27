DEFINT A-Z
'DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'TaxCust.BI'
  CLS
  PRINT "Processing Customer Accounts. . . "
  DIM TaxCustRec(1)  AS TaxCustType
  DIM TaxTrans(1)  AS TaxTransactionType
  PRINT
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))

  TaxTran = FREEFILE
  OPEN "TAXTRANS.DAT" FOR RANDOM SHARED AS #TaxTran LEN = LEN(TaxTrans(1))

  BCnt = 4888
  FOR Cnt = 513 TO 569
    LOCATE , 1
    PRINT "Processing:"; Cnt; "of 569";
    GET TaxFile, Cnt, TaxCustRec(1)
    BCnt = BCnt + 1
    TaxCustRec(1).LastTrans = BCnt
    PUT TaxFile, Cnt, TaxCustRec(1)
    GET TaxTran, BCnt, TaxTrans(1)
    TaxTrans(1).LastTrans = 0
    PUT TaxTran, BCnt, TaxTrans(1)
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT
  PRINT "Processing Complete. . ."
  CALL KillFile("fixem.exe")

