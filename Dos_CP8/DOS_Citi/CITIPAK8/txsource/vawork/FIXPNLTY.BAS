DEFINT A-Z
'DECLARE FUNCTION QPTrim$ (TheText$)
DECLARE FUNCTION uRound# (N#)
DECLARE FUNCTION Date2Num% (TheDate$)


TrType = 5

  '$INCLUDE: 'TaxCust.BI'
CLS
  REDIM TaxTran(1 TO 2) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))
  TDate = Date2Num("01-20-2004")

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen
  NumTR& = LOF(TaxTran) / TaxTranRecLen
  FOR Cnt& = 1 TO NumTR&
    LOCATE 5, 1
    PRINT "Processing:"; Cnt&; "of"; NumTR&;
    REDIM TaxTran(1 TO 2) AS TaxTransactionType
    GET #TaxTran, Cnt&, TaxTran(1)

    IF (TaxTran(1).TransDate = TDate) AND (TaxTran(1).Trantype = 5) THEN
      BillTR& = TaxTran(1).BelongTo
      GET #TaxTran, BillTR&, TaxTran(2)
      TaxTran(2).Revenue.Penalty = uRound#(TaxTran(2).Revenue.Penalty - TaxTran(1).Revenue.Penalty)
      PUT #TaxTran, BillTR&, TaxTran(2)
      TaxTran(1).Revenue.Penalty = 0
      TaxTran(1).amount = 0
      TaxTran(1).Description = "VOID USER ERROR"
      PUT #TaxTran, Cnt&, TaxTran(1)
      Fixed& = Fixed& + 1
    END IF
  NEXT

  CLOSE

  LOCATE 10, 1
  PRINT "Corrected:"; Fixed&
  CALL killfile("fixPnlty.exe")

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5000001) / 100
END FUNCTION

