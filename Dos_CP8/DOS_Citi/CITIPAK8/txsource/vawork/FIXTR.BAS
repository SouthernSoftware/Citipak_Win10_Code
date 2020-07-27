DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'

  DIM SHARED TaxTrans(1) AS TaxTransactionType

  NAME "TaxTRANS.DAT" AS "TaxTRANS.BAK"

  TaxTran = FREEFILE
  OPEN "TaxTRANS.BAK" FOR RANDOM SHARED AS TaxTran LEN = LEN(TaxTrans(1))
  NumOfTranRecs& = LOF(TaxTran) / LEN(TaxTrans(1))

  TTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TTran LEN = LEN(TaxTrans(1))

  FOR cnt = 1 TO 9957
    GET TaxTran, cnt, TaxTrans(1)
    PUT TTran, cnt, TaxTrans(1)
  NEXT

CLOSE

