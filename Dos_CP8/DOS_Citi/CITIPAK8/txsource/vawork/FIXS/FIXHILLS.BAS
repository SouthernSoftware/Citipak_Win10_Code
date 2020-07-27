DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'


  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  GET #TaxTran, 154, TaxTran(1)
  TaxTran(1).Revenue.Principle1Pd = TaxTran(1).Revenue.Principle1Pd - 3.75
  PUT #TaxTran, 154, TaxTran(1)
  CLOSE

  KILL "fixHills.exe"

