DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TAXRPTTY.BI'

  REDIM TaxTran(1 TO 2) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  GET TaxTran, 275, TaxTran(1)
  TaxTran(1).Revenue.interest = 54.38
  PUT TaxTran, 275, TaxTran(1)

  GET TaxTran, 2764, TaxTran(1)
  TaxTran(1).Revenue.interest = 11
  PUT TaxTran, 2764, TaxTran(1)
  CLOSE

  CALL KillFile("Fixgrndy.exe")

