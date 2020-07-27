DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'
  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen

  GET TaxTran, 2200, TaxTran(1)
  
  TaxTran(1).TaxYear = 1994
  PUT TaxTran, 2200, TaxTran(1)
  CLOSE
  CALL KILLFile("FIX2200.EXE")

