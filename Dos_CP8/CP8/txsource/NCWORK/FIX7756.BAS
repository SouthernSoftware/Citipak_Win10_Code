DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'
  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))
  CLS
  PRINT "Correcting Transaction."
  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  GET TaxTran, 7756, TaxTran(1)
  TaxTran(1).Revenue.Collection = 2
  TaxTran(1).Revenue.Interest = .72
  PUT TaxTran, 7756, TaxTran(1)
  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT "Done."
  CALL KILLFile("FIX7756.EXE")
