DEFINT A-Z
DECLARE SUB KILLFile (FileName$)

  '$INCLUDE: 'TaxCust.BI'

  CLS
  PRINT "Correcting. . ."
  

  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS #1 LEN = TaxTranRecLen

  NumOfRec& = LOF(1) / TaxTranRecLen
  Cnt& = 42749

  GET #1, Cnt&, TaxTran(1)
  TaxTran(1).CustomerRec = 3039
  'TaxTran(1).LastTrans = -1
  TaxTran(1).BelongTo = 40388
  TaxTran(1).Amount = 58.91
  TaxTran(1).Revenue.Principle1Pd = TaxTran(1).Amount
  PUT #1, Cnt&, TaxTran(1)
  PUT #1, Cnt&, TaxTran(1)

  Cnt& = 40388
  GET #1, Cnt&, TaxTran(1)
  TaxTran(1).Revenue.Principle1Pd = 58.91
  PUT #1, Cnt&, TaxTran(1)

'  TaxTran(1).Description = "2097"
  'TaxTran(1).Revenue.Principle1 = 0
  'TaxTran(1).Revenue.Principle2 = 0
  'TaxTran(1).Revenue.Principle3 = 0
  'TaxTran(1).Revenue.Principle4 = 0
  'TaxTran(1).Revenue.Principle5 = 0
  'TaxTran(1).Revenue.Interest = 0
  'TaxTran(1).Revenue.Penalty = 0
  'TaxTran(1).Revenue.Collection = 0
  'TaxTran(1).Revenue.Future1 = 0
  'TaxTran(1).Revenue.Future2 = 0
  'TaxTran(1).Revenue.Principle1Pd = 0
  'TaxTran(1).Revenue.Principle2Pd = 0
  'TaxTran(1).Revenue.Principle3Pd = 0
  'TaxTran(1).Revenue.Principle4Pd = 0
  'TaxTran(1).Revenue.Principle5Pd = 0
  'TaxTran(1).Revenue.InterestPd = 0
  'TaxTran(1).Revenue.PenaltyPd = 0
  'TaxTran(1).Revenue.CollectionPd = 0
  'TaxTran(1).Revenue.Future1Pd = 0
  'TaxTran(1).Revenue.Future2Pd = 0
  'TaxTran(1).Amount = 0
  'TaxTran(1).Description = "VOID VOID VOID"


  CLOSE
  PRINT
  PRINT
  PRINT "FIXED: 1"

  CALL KILLFile("fixit.exe")

