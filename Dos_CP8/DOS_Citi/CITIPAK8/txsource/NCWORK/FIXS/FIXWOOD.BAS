DEFINT A-Z
DECLARE FUNCTION Date2Num% (TheDate$)

  '$INCLUDE: 'TaxCust.BI'
  GoodDate = Date2Num("12-03-2002")

  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen

  GET TaxTran, 3230, TaxTran(1)
  TaxTran(1).TransDate = GoodDate
  PUT TaxTran, 3230, TaxTran(1)
  GET TaxTran, 3231, TaxTran(1)
  TaxTran(1).TransDate = GoodDate
  PUT TaxTran, 3231, TaxTran(1)
  CLOSE
  KILL "FIXWOOD.EXE"

