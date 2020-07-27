DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'
  CLS
  PRINT "Correcting Transaction. . ."

  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen

  GET #TaxTran, 883, TaxTran(1)
  TaxTran(1).Revenue.Principle1 = 1.32
  PUT #TaxTran, 883, TaxTran(1)


donehere:

  CLOSE

  LOCATE 6, 1
  PRINT "Done."

  CALL KILLFile("fixVICT.exe")

  END

