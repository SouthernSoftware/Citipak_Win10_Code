  DEFINT A-Z

  CONST False = 0, True = NOT False

  '$INCLUDE: 'bobTaxes.BI'
  '$INCLUDE: 'Taxfiles.BI'


  DIM SHARED TaxCustRec(1) AS TaxCustType
  CLS
  PRINT "Clearing old SSN's. . ."
  PRINT
  PRINT
  PRINT "Processing. . . "
  Taxfile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #Taxfile LEN = LEN(TaxCustRec(1))

  FOR Snt& = 1 TO LOF(Taxfile) / LEN(TaxCustRec(1))
    GET Taxfile, Snt&, TaxCustRec(1)
    TaxCustRec(1).CSSN = ""
    TaxCustRec(1).OSSN = ""
    PUT Taxfile, Snt&, TaxCustRec(1)
  NEXT
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT "Done."

  CALL KillFile("FixSSN.exe")

