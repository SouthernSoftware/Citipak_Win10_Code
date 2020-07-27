DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'


  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  FOR cnt = 31344 TO 31418
    GET #TaxTran, cnt, TaxTran(1)
    TaxTran(1).Description = "VOID  void  void"
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
    PUT #TaxTran, cnt, TaxTran(1)
  NEXT
  CLOSE

  CALL KILLFile("fix13687.exe")

