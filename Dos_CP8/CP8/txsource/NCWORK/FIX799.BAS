DEFINT A-Z
DECLARE FUNCTION Round# (N#)

  '$INCLUDE: 'TaxCust.BI'

  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen

  GET TaxTran, 32, TaxTran(1)
  TaxTran(1).Revenue.Principle1PD = 0
  PUT TaxTran, 32, TaxTran(1)
  CLOSE
  CALL KILLFile("FIX799.EXE")

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
'07-01-98
'corrected a bug which could cause certain numbers to round incorrectly
FUNCTION Round# (N#)
  Round# = INT(N# * 100 + .5000001#) / 100
END FUNCTION

