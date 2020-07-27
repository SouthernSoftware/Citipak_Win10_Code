DEFINT A-Z

  CLS
  LOCATE 2, 1
  PRINT "Correcting Operator ERROR. . ."

  '$INCLUDE: 'TaxCust.BI'

  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen

  GET TaxTran, 103015, TaxTran(1)
  TaxTran(1).Revenue.Principle1PD = 54.05
  PUT TaxTran, 103015, TaxTran(1)

  GET TaxTran, 112082, TaxTran(1)
  TaxTran(1).Revenue.Principle1PD = 54.05
  TaxTran(1).Amount = 54.05
  PUT TaxTran, 112082, TaxTran(1)

  CLOSE
  CALL KILLFile("FIXHAM.EXE")
  LOCATE 9, 1
  PRINT "PROGRAM COMPLETE. FIXED: 2"
  PRINT
  PRINT "Press any key to exit. . .";
  DO
    BEEP
    A$ = INKEY$
  LOOP UNTIL LEN(A$) > 0





