DEFINT A-Z
'DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'TaxCust.BI'
  CLS
  PRINT "Processing Customer Accounts. . . "
  DIM TaxCustRec(1)  AS TaxCustType
  PRINT
  PRINT
  PRINT "Real estate. . ."
  PRINT
  TaxFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  FOR Cnt = 1 TO NumOfTaxRecs
    LOCATE , 1
    PRINT "Processing:"; Cnt; "of"; NumOfTaxRecs;
    GET TaxFile, Cnt, TaxCustRec(1)
    TaxCustRec(1).LastTrans = 0
    PUT TaxFile, Cnt, TaxCustRec(1)
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT "Personal property. . ."
  PRINT
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  FOR Cnt = 1 TO NumOfTaxRecs
    LOCATE , 1
    PRINT "Processing:"; Cnt; "of"; NumOfTaxRecs;
    GET TaxFile, Cnt, TaxCustRec(1)
    TaxCustRec(1).LastTrans = 0
    PUT TaxFile, Cnt, TaxCustRec(1)
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT
  PRINT "Processing Complete. . ."
  CALL KillFile("fixem.exe")

