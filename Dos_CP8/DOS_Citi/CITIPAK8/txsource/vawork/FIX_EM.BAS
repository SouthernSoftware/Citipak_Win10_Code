DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  CLS
  PRINT "Correcting Transaction. . ."

  DIM SHARED TaxTrans(1) AS TaxTransactionType

  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS #1 LEN = LEN(TaxTrans(1))

  GET #1, 7305, TaxTrans(1)
  TaxTrans(1).Revenue.Principle1Pd = 0
  TaxTrans(1).Revenue.Principle2Pd = 0
  TaxTrans(1).Revenue.Principle3Pd = 0
  TaxTrans(1).Revenue.Principle4Pd = 0
  TaxTrans(1).Revenue.Principle5Pd = 0
  TaxTrans(1).Amount = 0
  TaxTrans(1).Description = "VOID  VOID  VOID"

  PUT #1, 7305, TaxTrans(1)

  GET #1, 694, TaxTrans(1)
  TaxTrans(1).Revenue.Principle2Pd = 0
  TaxTrans(1).Revenue.Principle3Pd = 0
  TaxTrans(1).Revenue.Principle4Pd = 0
  TaxTrans(1).Revenue.Principle5Pd = 0

  PUT #1, 694, TaxTrans(1)

  'GET #1, 156, TaxTrans(1)
  'TaxTrans(1).Revenue.Interest = .49
  'PUT #1, 156, TaxTrans(1)
  '
  'GET #1, 2656, TaxTrans(1)
  'TaxTrans(1).Revenue.Interest = 0
  'TaxTrans(1).Amount = 0
  'TaxTrans(1).Description = "VOID  VOID  VOID"
  'PUT #1, 2656, TaxTrans(1)

  CLOSE
  PRINT
  PRINT
  PRINT
  PRINT "Correction Complete."
  CALL KillFile("fix_em.exe")

