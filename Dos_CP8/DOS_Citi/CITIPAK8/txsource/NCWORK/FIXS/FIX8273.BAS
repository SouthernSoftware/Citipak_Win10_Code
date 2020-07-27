DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'
  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen

  GET TaxTran, 8273, TaxTran(1)
  TaxTran(1).Revenue.Interest = 16.34
  TaxTran(1).Revenue.InterestPD = 14.18
  PUT TaxTran, 8273, TaxTran(1)
  CLOSE
  KILL "FIX8273.EXE"

