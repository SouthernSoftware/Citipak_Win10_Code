DECLARE FUNCTION uRound# (N#)
DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'

  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen

  GET TaxTran, 1119, TaxTran(1)
  TaxTran(1).Revenue.Interest = 0
  TaxTran(1).Revenue.InterestPD = 0
  TaxTran(1).Amount = 0
  PUT TaxTran, 1119, TaxTran(1)
  GET TaxTran, 3170, TaxTran(1)
  TaxTran(1).Amount = 0
  PUT TaxTran, 3170, TaxTran(1)
  CLOSE
  CALL KILLFile("FIXSugar.exe")

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
'07-01-98
'corrected a bug which could cause certain numbers to round incorrectly
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5000001#) / 100
END FUNCTION

