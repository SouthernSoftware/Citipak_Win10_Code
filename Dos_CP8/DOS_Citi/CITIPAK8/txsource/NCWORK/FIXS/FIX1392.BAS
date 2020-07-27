DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'
  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen

  GET TaxTran, 2027, TaxTran(1)
  TaxTran(1).LastTrans = 599
  PUT TaxTran, 2027, TaxTran(1)
  CLOSE
  KILL "FIX1392.EXE"

