DEFINT A-Z
DECLARE SUB KILLFile (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'

  CLS
  WhatDate = Date2Num("04-27-2000")

  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS #1 LEN = TaxTranRecLen

  NumOfRec& = LOF(1) / TaxTranRecLen
  FOR Cnt& = 1 TO NumOfRec&
    LOCATE 5, 1: PRINT "Processing:"; Cnt&;
    GET #1, Cnt&, TaxTran(1)
    IF TaxTran(1).TransDate = WhatDate AND TaxTran(1).TranType = 2 THEN
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
      TaxTran(1).Description = "VOIDED: " + DATE$
      Found = Found + 1
      PUT #1, Cnt&, TaxTran(1)
    END IF
  NEXT

  CLOSE
  PRINT
  PRINT
  PRINT "FIXED:"; Found

  CALL KILLFile("fixem.exe")

