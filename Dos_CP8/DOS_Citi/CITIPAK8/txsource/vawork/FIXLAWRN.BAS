DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'TaxCust.BI'
  CLS
  PRINT "Processing Customer Accounts. . . "
  PRINT
  PRINT "Personal property. . ."

  DIM TaxCustRec(1)  AS TaxCustType
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
      PUT TaxFile, Cnt, TaxCustRec(1)
      FoundEm = FoundEm + 1
    END IF
    CSSN$ = QPTrim$(TaxCustRec(1).SSSN)
    IF LEN(CSSN$) > 0 AND LEN(CSSN$) < 3 THEN
      TaxCustRec(1).SSSN = ""
      PUT TaxFile, Cnt, TaxCustRec(1)
      BadSsn = BadSsn + 1

    END IF
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT "Real Estate..."
  TaxFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  FOR Cnt = 1 TO NumOfTaxRecs
    LOCATE , 1
    PRINT "Processing:"; Cnt; "of"; NumOfTaxRecs;
    GET TaxFile, Cnt, TaxCustRec(1)
    CityState$ = QPTrim$(TaxCustRec(1).City)

    IF INSTR(CityState$, ",") > 0 THEN
      FoundAt = LEN(CityState$)
      TaxCustRec(1).City = LEFT$(CityState$, FoundAt - 1)
      PUT TaxFile, Cnt, TaxCustRec(1)
      RFoundEm = RFoundEm + 1
    END IF
    CSSN$ = QPTrim$(TaxCustRec(1).SSSN)
    IF LEN(CSSN$) > 0 AND LEN(CSSN$) < 3 THEN
      BadSsn = BadSsn + 1
      TaxCustRec(1).SSSN = ""
      PUT TaxFile, Cnt, TaxCustRec(1)
    END IF

  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT "Personal Property:"; FoundEm
  PRINT "      Real Estate:"; RFoundEm
  PRINT "    SSN Corrected:"; BadSsn
  PRINT "Processing Complete. . ."
  CALL KillFile("fixlawrn.exe")

