DEFINT A-Z
DECLARE SUB KILLFile (FileName$)
DECLARE FUNCTION Date2Num (TheDate$)

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'

  CLS
  PRINT "Correcting Transactions. . ."

  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS #1 LEN = TaxTranRecLen


    GET #1, 28176, TaxTran(1)
    TaxTran(1).Revenue.InterestPd = 4.26
    PUT #1, 28176, TaxTran(1)

    GET #1, 35133, TaxTran(1)
    TaxTran(1).Amount = 0
    TaxTran(1).Revenue.InterestPd = 0
    TaxTran(1).description = "VOID VOID VOID"
    PUT #1, 35133, TaxTran(1)

    GET #1, 1243, TaxTran(1)
    TaxTran(1).Revenue.Principle1Pd = 158.67
    PUT #1, 1243, TaxTran(1)

    GET #1, 27290, TaxTran(1)
    TaxTran(1).Revenue.Principle1Pd = 0
    PUT #1, 27290, TaxTran(1)


    GET #1, 32217, TaxTran(1)
    TaxTran(1).CustomerRec = 0
    TaxTran(1).Revenue.Principle1Pd = 0
    TaxTran(1).Amount = 0
    TaxTran(1).Revenue.InterestPd = 0
    TaxTran(1).description = "VOID VOID VOID"
    PUT #1, 32217, TaxTran(1)

    GET #1, 32218, TaxTran(1)
    TaxTran(1).CustomerRec = 0
    TaxTran(1).Revenue.Principle1Pd = 0
    TaxTran(1).Amount = 0
    TaxTran(1).Revenue.InterestPd = 0
    TaxTran(1).description = "VOID VOID VOID"
    PUT #1, 32218, TaxTran(1)

    GET #1, 32219, TaxTran(1)
    TaxTran(1).CustomerRec = 0
    TaxTran(1).Revenue.Principle1Pd = 0
    TaxTran(1).Amount = 0
    TaxTran(1).Revenue.InterestPd = 0
    TaxTran(1).description = "VOID VOID VOID"
    PUT #1, 32219, TaxTran(1)

    GET #1, 31194, TaxTran(1)
    TaxTran(1).CustomerRec = 0
    TaxTran(1).Revenue.Principle1Pd = 0
    TaxTran(1).Amount = 0
    TaxTran(1).Revenue.InterestPd = 0
    TaxTran(1).description = "VOID VOID VOID"
    PUT #1, 31194, TaxTran(1)

    GET #1, 30424, TaxTran(1)
    TaxTran(1).CustomerRec = 0
    TaxTran(1).Revenue.Principle1Pd = 0
    TaxTran(1).Amount = 0
    TaxTran(1).Revenue.InterestPd = 0
    TaxTran(1).description = "VOID VOID VOID"
    PUT #1, 30424, TaxTran(1)

    GET #1, 30363, TaxTran(1)
    TaxTran(1).CustomerRec = 0
    TaxTran(1).Revenue.Principle1Pd = 0
    TaxTran(1).Amount = 0
    TaxTran(1).Revenue.InterestPd = 0
    TaxTran(1).description = "VOID VOID VOID"
    PUT #1, 30363, TaxTran(1)


DoneHere:
  CLOSE
  PRINT
  PRINT
  PRINT "dONE"

  CALL KILLFile("fixem.exe")

