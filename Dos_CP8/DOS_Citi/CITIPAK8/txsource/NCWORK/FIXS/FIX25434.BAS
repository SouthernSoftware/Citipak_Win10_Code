DEFINT A-Z
  CLS

  PRINT
  PRINT
  PRINT "Correcting Transaction."

  '$INCLUDE: 'TaxCust.BI'

  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  GET #TaxTran, 25434, TaxTran(1)
  TaxTran(1).Belongto = 20281
  PUT #TaxTran, 25434, TaxTran(1)

  CLOSE

  CALL KILLFile("fix25434.exe")

  PRINT
  PRINT
  PRINT "Done."

END

