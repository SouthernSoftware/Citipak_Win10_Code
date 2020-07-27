DEFINT A-Z
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION TRound# (N#)

  '$INCLUDE: 'TaxCust.BI'

  CLS

  TDate = Date2Num("04-10-2001")

  REDIM TaxTran(1 TO 2) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))
  PRINT "Correcting Adv/Collections. . ."
  
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS #1 LEN = TaxTranRecLen
  NumTr& = LOF(1) \ TaxTranRecLen
  FOR Cnt& = 1 TO NumTr&
  GET #1, Cnt&, TaxTran(1)
  LOCATE 14, 1: PRINT "Processing:"; Cnt&; "of"; NumTr&;
  IF TaxTran(1).TransDate = TDate THEN
    SELECT CASE TaxTran(1).TranType
    CASE 6
      IF TaxTran(1).BelongTo > 0 THEN
        GET #1, TaxTran(1).BelongTo, TaxTran(2)
        TaxTran(2).Revenue.Collection = TRound(TaxTran(2).Revenue.Collection - TaxTran(1).Revenue.Collection)
        TaxTran(2).Revenue.Interest = TRound(TaxTran(2).Revenue.Interest + TaxTran(1).Revenue.Collection)
        PUT #1, TaxTran(1).BelongTo, TaxTran(2)
      END IF
      SWAP TaxTran(1).Revenue.Collection, TaxTran(1).Revenue.Interest
      PUT #1, Cnt&, TaxTran(1)
      Fixed = Fixed + 1
    END SELECT
  END IF
  NEXT
  CLOSE

  PRINT
  PRINT
  PRINT
  PRINT "Fixed:"; Fixed

  CALL KILLFILE("FIXHAM.EXE")
  END


'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
'07-01-98
'corrected a bug which could cause certain numbers to round incorrectly
FUNCTION TRound# (N#)
  TRound# = INT(N# * 100 + .5000001#) / 100
END FUNCTION

