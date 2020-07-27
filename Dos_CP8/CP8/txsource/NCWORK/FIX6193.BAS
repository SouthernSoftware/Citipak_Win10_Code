DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'
  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen

  GET TaxTran, 6193, TaxTran(1)
  'TaxTran(1).LastTrans = 599
  TaxTran(1).Revenue.Collection = 0
  TaxTran(1).Revenue.CollectionPD = 0
  TaxTran(1).Description = "VOID - OPERATOR ERROR"
  TaxTran(1).Amount = 0
  PUT TaxTran, 6193, TaxTran(1)
  CLOSE
  CALL KILLFile("FIX6193.EXE")

