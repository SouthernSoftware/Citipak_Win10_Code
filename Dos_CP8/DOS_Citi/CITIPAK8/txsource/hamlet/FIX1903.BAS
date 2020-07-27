DEFINT A-Z
DECLARE FUNCTION TRound# (N#)

  '$INCLUDE: 'TaxCust.BI'

  CLS

  REDIM TaxTran(1 TO 3) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))
  
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS #1 LEN = TaxTranRecLen
  GET #1, 12588, TaxTran(1)                'payment
  GET #1, TaxTran(1).BelongTo, TaxTran(2)  'bill
  TaxTran(2).Revenue.Principle1Pd = TRound#(TaxTran(2).Revenue.Principle1Pd - TaxTran(1).Revenue.Principle1Pd)
  TaxTran(2).Revenue.Principle2Pd = TRound#(TaxTran(2).Revenue.Principle2Pd - TaxTran(1).Revenue.Principle2Pd)
  TaxTran(2).Revenue.Principle3Pd = TRound#(TaxTran(2).Revenue.Principle3Pd - TaxTran(1).Revenue.Principle3Pd)
  TaxTran(2).Revenue.Principle4Pd = TRound#(TaxTran(2).Revenue.Principle4Pd - TaxTran(1).Revenue.Principle4Pd)
  TaxTran(2).Revenue.Principle5Pd = TRound#(TaxTran(2).Revenue.Principle5Pd - TaxTran(1).Revenue.Principle5Pd)
  TaxTran(2).Revenue.InterestPd = TRound#(TaxTran(2).Revenue.InterestPd - TaxTran(1).Revenue.InterestPd)
  TaxTran(2).Revenue.CollectionPD = TRound#(TaxTran(2).Revenue.CollectionPD - TaxTran(1).Revenue.CollectionPD)
  PUT #1, TaxTran(1).BelongTo, TaxTran(2)  'bill
  PUT #1, 12588, TaxTran(3)                'payment

  GET #1, 12589, TaxTran(1)                'payment
  GET #1, TaxTran(1).BelongTo, TaxTran(2)  'bill
  TaxTran(2).Revenue.Principle1Pd = TRound#(TaxTran(2).Revenue.Principle1Pd - TaxTran(1).Revenue.Principle1Pd)
  TaxTran(2).Revenue.Principle2Pd = TRound#(TaxTran(2).Revenue.Principle2Pd - TaxTran(1).Revenue.Principle2Pd)
  TaxTran(2).Revenue.Principle3Pd = TRound#(TaxTran(2).Revenue.Principle3Pd - TaxTran(1).Revenue.Principle3Pd)
  TaxTran(2).Revenue.Principle4Pd = TRound#(TaxTran(2).Revenue.Principle4Pd - TaxTran(1).Revenue.Principle4Pd)
  TaxTran(2).Revenue.Principle5Pd = TRound#(TaxTran(2).Revenue.Principle5Pd - TaxTran(1).Revenue.Principle5Pd)
  TaxTran(2).Revenue.InterestPd = TRound#(TaxTran(2).Revenue.InterestPd - TaxTran(1).Revenue.InterestPd)
  TaxTran(2).Revenue.CollectionPD = TRound#(TaxTran(2).Revenue.CollectionPD - TaxTran(1).Revenue.CollectionPD)
  PUT #1, TaxTran(1).BelongTo, TaxTran(2)  'bill
  PUT #1, 12589, TaxTran(3)                'payment


  'TaxTran (2)
  'PUT #1, 12589, TaxTran(1)


  CLOSE

'LOCATE 15, 1
'PRINT "Done."

'CALL KILLFILE("FIXADV2.EXE")
'  PRINT TaxTran(2).Revenue.Principle1Pd, TaxTran(1).Revenue.Principle1Pd
'  PRINT TaxTran(2).Revenue.Principle2Pd, TaxTran(1).Revenue.Principle2Pd
'  PRINT TaxTran(2).Revenue.Principle3Pd, TaxTran(1).Revenue.Principle3Pd
'  PRINT TaxTran(2).Revenue.Principle4Pd, TaxTran(1).Revenue.Principle4Pd
'  PRINT TaxTran(2).Revenue.Principle5Pd, TaxTran(1).Revenue.Principle5Pd
'  PRINT "I "; TaxTran(2).Revenue.InterestPd, TaxTran(1).Revenue.InterestPd
'  PRINT TaxTran(2).Revenue.CollectionPD, TaxTran(1).Revenue.CollectionPD
'  PRINT
'  PRINT

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
'07-01-98
'corrected a bug which could cause certain numbers to round incorrectly
FUNCTION TRound# (N#)
  TRound# = INT(N# * 100 + .5000001#) / 100
END FUNCTION

