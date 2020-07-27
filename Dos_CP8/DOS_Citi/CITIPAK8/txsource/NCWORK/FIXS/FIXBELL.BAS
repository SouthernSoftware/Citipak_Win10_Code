DEFINT A-Z
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION uRound# (N#)

  '$INCLUDE: 'TaxCust.BI'

  BDate = Date2Num("05-17-2002")

  DIM TaxTran(1 TO 2) AS TaxTransactionType
  CLS
  LOCATE 3, 1
  PRINT "Correcting User Errors. . ."
  TRLen = LEN(TaxTran(1))
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS #1 LEN = TRLen
  NumTR& = LOF(1) / TRLen
  FOR Cnt& = 1 TO NumTR&
    GET #1, Cnt&, TaxTran(1)
    IF TaxTran(1).TransDate = BDate THEN
      IF TaxTran(1).TranType = 6 THEN
        'GotAdj = GotAdj + 1
        GET #1, TaxTran(1).BeLongTo, TaxTran(2)
        TaxTran(2).Revenue.Collection = uRound#(TaxTran(2).Revenue.Collection - TaxTran(1).Revenue.Collection)
        TaxTran(1).Revenue.Collection = 0
        TaxTran(1).Amount = 0
        PUT #1, TaxTran(1).BeLongTo, TaxTran(2)
        PUT #1, Cnt&, TaxTran(1)
      ELSEIF TaxTran(1).TranType = 7 THEN
        GET #1, TaxTran(1).BeLongTo, TaxTran(2)
        TaxTran(2).Revenue.CollectionPD = uRound#(TaxTran(2).Revenue.CollectionPD - TaxTran(1).Revenue.CollectionPD)
        TaxTran(1).Revenue.CollectionPD = 0
        TaxTran(1).Amount = 0
        PUT #1, TaxTran(1).BeLongTo, TaxTran(2)
        PUT #1, Cnt&, TaxTran(1)
      END IF
    END IF
  NEXT
  CLOSE
  NAME "taxtrans.dat" AS "taxtrans.gof"

  OPEN "TaxTRANS.gof" FOR RANDOM SHARED AS #1 LEN = TRLen
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS #2 LEN = TRLen
  NumTR& = LOF(1) / TRLen
  FOR Cnt& = 1 TO NumTR&
    GET #1, Cnt&, TaxTran(1)
    IF TaxTran(1).TransDate <> BDate THEN
      PUT #2, , TaxTran(1)
    END IF
  NEXT
  CLOSE

  LOCATE 7, 1

  PRINT "Correction Complete:"

  CALL killfile("fixbell.exe")

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION uRound# (N#)
  uRound# = INT(N# * 100 + .5) / 100
END FUNCTION

