DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'TaxCust.BI'
  CLS
  PRINT "Processing Customer Accounts. . . "
  PRINT "Personal property. . ."

  DIM TaxCustRec(1)  AS TaxCustType
  PRINT
  PRINT
  PRINT "Real estate. . ."
  PRINT
  TaxFile = FREEFILE
  OPEN "PPTXCUST.DAT" FOR RANDOM SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  FOR Cnt = 1 TO NumOfTaxRecs
    LOCATE , 1
    PRINT "Processing:"; Cnt; "of"; NumOfTaxRecs;
    GET TaxFile, Cnt, TaxCustRec(1)
    CityState$ = QPTrim$(TaxCustRec(1).City)

    IF INSTR(CityState$, ",") > 0 AND LEN(QPTrim$(TaxCustRec(1).State)) = 0 THEN
      FoundAt = INSTR(CityState$, ",")
      TaxCustRec(1).City = MID$(CityState$, 1, FoundAt - 1)
      TaxCustRec(1).State = RIGHT$(CityState$, 2)
      'PUT TaxFile, Cnt, TaxCustRec(1)
      FoundEm = FoundEm + 1
    END IF

    'TaxCustRec(1).LastTrans = 0
    '
  NEXT
  PRINT
  PRINT
  PRINT
  PRINT
  PRINT
  PRINT "Found:"; FoundEm
  PRINT "Processing Complete. . ."
  CALL KillFile("fixem.exe")

