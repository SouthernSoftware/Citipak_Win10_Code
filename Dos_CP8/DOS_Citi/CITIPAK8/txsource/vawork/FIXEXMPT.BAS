DEFINT A-Z

  '$INCLUDE: 'bobTaxes.BI'
'  '$INCLUDE: 'Taxfiles.BI'
'  '$INCLUDE: 'PROPAbst.BI'

  DIM SHARED TaxCustRec(1) AS TaxCustType
  'DIM SHARED PropertyRec(1) AS PropertyRecType
  'DIM SHARED REALRec(1) AS PropertyRecType
  CLS
  PRINT "Setting Tax Exempt Flags. . ."

  TaxFile = FREEFILE
  OPEN "TaXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  OPEN "exempt.txt" FOR INPUT AS #20
  DO
    GotEm = 0
    LINE INPUT #20, Cnty$
    CntCty& = VAL(Cnty$)

   'Find Exiting Acct
    FOR Snt& = 1 TO NumOfTaxRecs
      GET TaxFile, Snt&, TaxCustRec(1)
      IF TaxCustRec(1).CountyAcct = CntCty& THEN
        TaxCustRec(1).TaxExempt = "Y"
        DidCnt = DidCnt + 1
        GotEm = -1
        PUT TaxFile, Snt&, TaxCustRec(1)
        EXIT FOR
      END IF
    NEXT
    IF GotEm = 0 THEN STOP
  LOOP UNTIL EOF(20)

  CLOSE
  PRINT
  PRINT
  PRINT "Changed:"; DidCnt
  PRINT
  PRINT "Done."


