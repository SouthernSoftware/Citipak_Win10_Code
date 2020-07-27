DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'


  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  GET #TaxTran, 8561, TaxTran(1)
  TaxTran(1).Revenue.Principle5Pd = 4.14
  PUT #TaxTran, 8561, TaxTran(1)
  CLOSE

  kill "fix1473.exe"
