DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'

  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen

  GET TaxTran, 32, TaxTran(1)
  TaxTran(1).Revenue.Principle1PD = 0
  PUT TaxTran, 32, TaxTran(1)
  CLOSE
  CALL KILLFile("FIXTR32.EXE")

