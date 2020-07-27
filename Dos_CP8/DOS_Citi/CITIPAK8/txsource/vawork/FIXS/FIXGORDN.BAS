DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'


  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  GET #TaxTran, 13405, TaxTran(1)
  GOSUB ZeroTrans
  PUT #TaxTran, 13405, TaxTran(1)

  GET #TaxTran, 13406, TaxTran(1)
  GOSUB ZeroTrans
  PUT #TaxTran, 13406, TaxTran(1)

  GET #TaxTran, 13407, TaxTran(1)
  GOSUB ZeroTrans
  PUT #TaxTran, 13407, TaxTran(1)

  GET #TaxTran, 13408, TaxTran(1)
  GOSUB ZeroTrans
  PUT #TaxTran, 13408, TaxTran(1)
'done with payments

  GET #TaxTran, 12283, TaxTran(1)
  GOSUB FixBill
  PUT #TaxTran, 12283, TaxTran(1)

  GET #TaxTran, 2552, TaxTran(1)
  GOSUB FixBill
  PUT #TaxTran, 2552, TaxTran(1)

  GET #TaxTran, 2553, TaxTran(1)
  GOSUB FixBill
  PUT #TaxTran, 2553, TaxTran(1)

  GET #TaxTran, 2554, TaxTran(1)
  GOSUB FixBill
  PUT #TaxTran, 2554, TaxTran(1)

  CLOSE

  END

ZeroTrans:
  TaxTran(1).Amount = 0
  TaxTran(1).Revenue.Principle1 = 0
  TaxTran(1).Revenue.Principle2 = 0
  TaxTran(1).Revenue.Principle3 = 0
  TaxTran(1).Revenue.Principle4 = 0
  TaxTran(1).Revenue.Principle5 = 0
  TaxTran(1).Revenue.Interest = 0
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
  TaxTran(1).Description = "  VOID VOID VOID!"
RETURN

FixBill:
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
RETURN

