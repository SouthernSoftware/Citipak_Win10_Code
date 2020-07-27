DECLARE FUNCTION uRound# (N#)
DEFINT A-Z

  '$INCLUDE: 'TaxCust.BI'
CLS
  REDIM TaxTran(1) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  NumTR& = LOF(TaxTran) / TaxTranRecLen
  FOR Cnt& = 1 TO NumTR&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; "of"; NumTR&;
    GET #TaxTran, Cnt&, TaxTran(1)
    IF TaxTran(1).Revenue.Collection > 0 THEN
      Fixed& = Fixed& + 1
      TaxTran(1).Revenue.Penalty = uRound#(TaxTran(1).Revenue.Penalty + TaxTran(1).Revenue.Collection)
      TaxTran(1).Revenue.PenaltyPD = uRound#(TaxTran(1).Revenue.PenaltyPD + TaxTran(1).Revenue.CollectionPD)
      TaxTran(1).Revenue.Collection = 0
      TaxTran(1).Revenue.CollectionPD = 0
      PUT #TaxTran, Cnt&, TaxTran(1)
    END IF
  NEXT

  CLOSE

  LOCATE 10, 1
  PRINT "Corrected:"; Fixed&
  CALL killfile("fixvictr.exe")

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5000001) / 100
END FUNCTION

