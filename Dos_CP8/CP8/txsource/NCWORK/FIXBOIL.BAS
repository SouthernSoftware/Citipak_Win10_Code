DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'
  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))
  CLS
  PRINT "Correcting Transactions."
  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  GET TaxTran, 7756, TaxTran(1)
  TaxTran(1).Revenue.Collection = 2
  TaxTran(1).Revenue.Interest = .72
  PUT TaxTran, 7756, TaxTran(1)
  GET TaxTran, 61507, TaxTran(1)
  TaxTran(1).Revenue.InterestPD = 2.62
  PUT TaxTran, 61507, TaxTran(1)

  GET TaxTran, 104757, TaxTran(1)
  TaxTran(1).Revenue.Interest = 0
  TaxTran(1).Revenue.InterestPD = 0
  TaxTran(1).Amount = 0
  TaxTran(1).Description = "VOID DOUBLED??"
  PUT TaxTran, 104757, TaxTran(1)

  GET TaxTran, 12691, TaxTran(1)
  TaxTran(1).Revenue.Interest = .98
  PUT TaxTran, 12691, TaxTran(1)

  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT "Done."
  CALL KILLFile("FIXBOIL.EXE")

