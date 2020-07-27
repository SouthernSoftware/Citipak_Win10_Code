DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'
  CLS

  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  GET #TaxTran, 17269, TaxTran(1)
  GOSUB resetinfo
  PUT #TaxTran, 17269, TaxTran(1)

  GET #TaxTran, 17087, TaxTran(1)
  GOSUB resetinfo
  PUT #TaxTran, 17087, TaxTran(1)

  GET #TaxTran, 10580, TaxTran(1)
  GOSUB resetinfo
  PUT #TaxTran, 10580, TaxTran(1)

  GET #TaxTran, 17114, TaxTran(1)
  GOSUB resetinfo
  PUT #TaxTran, 17114, TaxTran(1)

  GET #TaxTran, 17325, TaxTran(1)
  GOSUB resetinfo
  PUT #TaxTran, 17325, TaxTran(1)

  CLOSE

  CALL KILLFile("fixBOIL.exe")
  PRINT "Done."
END

resetinfo:
  TaxTran(1).Description = "VOIDED"
  TaxTran(1).Revenue.Principle1 = 0
  TaxTran(1).Revenue.Principle2 = 0
  TaxTran(1).Revenue.Principle3 = 0
  TaxTran(1).Revenue.Principle4 = 0
  TaxTran(1).Revenue.Principle5 = 0
  TaxTran(1).Revenue.interest = 0
  TaxTran(1).Revenue.Penalty = 0
  TaxTran(1).Revenue.Collection = 0
  TaxTran(1).Revenue.Future1 = 0
  TaxTran(1).Revenue.Future2 = 0
  TaxTran(1).Revenue.Principle1Pd = 0
  TaxTran(1).Revenue.Principle2Pd = 0
  TaxTran(1).Revenue.Principle3Pd = 0
  TaxTran(1).Revenue.Principle4Pd = 0
  TaxTran(1).Revenue.Principle5Pd = 0
  TaxTran(1).Revenue.InterestPd = 0
  TaxTran(1).Revenue.PenaltyPd = 0
  TaxTran(1).Revenue.CollectionPd = 0
  TaxTran(1).Revenue.Future1Pd = 0
  TaxTran(1).Revenue.Future2Pd = 0
  TaxTran(1).Amount = 0
  TaxTran(1).CustomerRec = -1

RETURN

