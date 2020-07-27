DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'

  DIM TaxTran AS TaxTransactionType
  CLS
  LOCATE 3, 1
  PRINT "Correcting User Errors. . ."

  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS #1 LEN = LEN(TaxTran)
  GET #1, 8295, TaxTran
  TaxTran.Revenue.Interest# = 17.59
  PUT #1, 8295, TaxTran
  CLOSE
  LOCATE 7, 1
  PRINT "Correction Complete."
  CALL killfile("fixsunbh.exe")


