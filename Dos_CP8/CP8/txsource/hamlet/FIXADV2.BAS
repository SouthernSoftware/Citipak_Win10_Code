DEFINT A-Z
DECLARE FUNCTION TRound# (N#)

  '$INCLUDE: 'TaxCust.BI'

  CLS

  REDIM TaxTran(1 TO 2) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))
  
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS #1 LEN = TaxTranRecLen
  NumTr& = LOF(1) \ TaxTranRecLen
  FOR Cnt& = 1 TO NumTr&
    GET #1, Cnt&, TaxTran(1)
    LOCATE 12, 1: PRINT "Processing:"; Cnt&; "of"; NumTr&;
    Diff# = TRound#(TaxTran(1).Revenue.Collection + TaxTran(1).Revenue.CollectionPd)
    IF Diff# <> 0 THEN
      TaxTran(1).Revenue.InterestPD = TRound#(TaxTran(1).Revenue.InterestPD + TaxTran(1).Revenue.CollectionPd)
      TaxTran(1).Revenue.CollectionPd = 0
      PUT #1, Cnt&, TaxTran(1)
      Fixed = Fixed + 1
    END IF
  NEXT
  CLOSE

LOCATE 15, 1
PRINT "Done."

CALL KILLFILE("FIXADV2.EXE")

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
'07-01-98
'corrected a bug which could cause certain numbers to round incorrectly
FUNCTION TRound# (N#)
  TRound# = INT(N# * 100 + .5000001#) / 100
END FUNCTION

