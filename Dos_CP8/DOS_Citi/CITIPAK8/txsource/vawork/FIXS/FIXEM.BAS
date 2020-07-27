DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

  '$INCLUDE: 'TaxCust.BI'
  CLS
  PRINT "Removing Blank Accounts. . . "
  DIM TaxCustRec(1)  AS TaxCustType

  TaxFile = FREEFILE
  OPEN "RETXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  FOR Cnt = 1 TO NumOfTaxRecs
    LOCATE 5, 1
    PRINT "Processing:"; Cnt; "of"; NumOfTaxRecs;
    GET TaxFile, Cnt, TaxCustRec(1)
    FirstName$ = QPTrim$(TaxCustRec(1).FName)
    LastName$ = QPTrim$(TaxCustRec(1).LName)
    IF LEN(FirstName$) = 0 AND LEN(LastName$) = 0 THEN
      TaxCustRec(1).Deleted = -1
      PUT TaxFile, Cnt, TaxCustRec(1)
    END IF
  NEXT
  CLOSE

  LOCATE 8, 1
  PRINT "Processing Complete. . ."
  CALL KillFile("fixem.exe")

