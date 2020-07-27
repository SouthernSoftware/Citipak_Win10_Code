DEFINT A-Z
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION TRound# (N#)

  '$INCLUDE: 'TaxCust.BI'

  CLS
  PRINT "Processing. . ."
  TDate = Date2Num("03-25-2002")

  REDIM TaxTran(1 TO 2) AS TaxTransactionType
  TaxTranRecLen = LEN(TaxTran(1))

'  REDIM TaxCust(1) AS TaxCustType
'  CustLen = LEN(TaxCust(1))
'
'  OPEN "TaxCUST.DAT" FOR RANDOM SHARED AS #1 LEN = CustLen
'  NumCust& = LOF(1) \ CustLen
'  FOR Cnt& = 1 TO NumCust&
'    GET #1, Cnt&, TaxCust(1)


  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS #1 LEN = TaxTranRecLen
  NumTr& = LOF(1) \ TaxTranRecLen
  FOR Cnt& = 1 TO NumTr&
  GET #1, Cnt&, TaxTran(1)

  IF TaxTran(1).TransDate = TDate THEN
    IF TaxTran(1).TranType = 6 THEN
      IF TaxTran(1).BelongTo > 0 THEN
        GET #1, TaxTran(1).BelongTo, TaxTran(2)
        TaxTran(2).Revenue.interest = TaxTran(2).Revenue.interest + TaxTran(1).Amount
        TaxTran(2).Revenue.Collection = 0
        PUT #1, TaxTran(1).BelongTo, TaxTran(2)
'        'STOP
      END IF
      TaxTran(1).Revenue.interest = TaxTran(1).Amount
      TaxTran(1).Revenue.Collection = 0
      Fixed = Fixed + 1
      PUT #1, Cnt&, TaxTran(1)
    END IF
  END IF
  NEXT
  CLOSE

LOCATE 15, 1
PRINT "Fixed:"; Fixed

CALL KILLFILE("FIXADV.EXE")

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
'07-01-98
'corrected a bug which could cause certain numbers to round incorrectly
FUNCTION TRound# (N#)
  TRound# = INT(N# * 100 + .5000001#) / 100
END FUNCTION

