  DECLARE FUNCTION Num2Date$ (DateNumber%)
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'Taxfiles.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TAXRPTTY.BI'
  
  CLS
  
  DIM SHARED TaxCustRec(1) AS TaxCustType
  DIM SHARED PropertyRec(1) AS PropertyRecType
  DIM SHARED PersRec(1) AS PersonalRecType
  DIM SHARED TaxTrans(1) AS TaxTransactionType

  TaxCustRecLen = LEN(TaxCustRec(1))
  TaxTranRecLen = LEN(TaxTrans(1))

  CLS
  INPUT "Enter Acct # "; RecNo&


  TaxFile = FREEFILE
  OPEN "reTXCUST.DAT" FOR RANDOM SHARED AS TaxFile LEN = TaxCustRecLen

  TaxTran = FREEFILE
  OPEN "TaxTRANS.DAT" FOR RANDOM SHARED AS TaxTran LEN = TaxTranRecLen





  GET TaxFile, RecNo&, TaxCustRec(1)
  


  PrevTranRec& = TaxCustRec(1).LastTrans
  CLS
  DO WHILE PrevTranRec& <> 0
  GET TaxTran, PrevTranRec&, TaxTrans(1)
     PRINT "   Record Number: "; PrevTranRec&
     PRINT "Transaction Type: "; TaxTrans(1).TranType
     PRINT "Transaction Date: "; Num2Date$(TaxTrans(1).TransDate)
     PRINT "Transaction Belg: "; TaxTrans(1).BelongTo
     PRINT "       Bill Type: "; TaxTrans(1).BillType
     PRINT "     Description: "; TaxTrans(1).Description
     PRINT "Customer Record : "; TaxTrans(1).CustomerRec
     PRINT "          Amount: "; TaxTrans(1).Amount
     PRINT

    IF TaxTrans(1).TranType = 2 THEN
     PRINT "Prin1 Paid: "; TaxTrans(1).Revenue.Principle1Pd
     PRINT "Prin2 Paid: "; TaxTrans(1).Revenue.Principle2Pd
     PRINT "Prin3 Paid: "; TaxTrans(1).Revenue.Principle3Pd
     PRINT "Prin4 Paid: "; TaxTrans(1).Revenue.Principle4Pd
     PRINT "Prin5 Paid: "; TaxTrans(1).Revenue.Principle5Pd
     PRINT "Intst Paid: "; TaxTrans(1).Revenue.InterestPd
     PRINT "Penty Paid: "; TaxTrans(1).Revenue.PenaltyPd

     TotalDetail# = TaxTrans(1).Revenue.Principle1Pd + TaxTrans(1).Revenue.Principle2Pd + TaxTrans(1).Revenue.Principle3Pd
     TotalDetail# = TotalDetail# + TaxTrans(1).Revenue.Principle4Pd + TaxTrans(1).Revenue.Principle5Pd + TaxTrans(1).Revenue.InterestPd + TaxTrans(1).Revenue.PenaltyPd
     TotalDetail# = INT((TotalDetail# * 100) + .5) / 100
    ' IF TotalDetail# = 0 THEN
    '  TaxTrans(1).Revenue.Principle1Pd = TaxTrans(1).Amount
    '  PUT TaxTran, PrevTranRec&, TaxTrans(1)
    ' END IF
    ' IF TotalDetail# >0 and TotalDetail# TaxTrans(1).Amount THEN
    '  TaxTrans(1).Revenue.Principle1Pd = TaxTrans(1).Amount
    ' END IF
    END IF
    TotalDetail# = 0

    IF TaxTrans(1).TranType = 1 THEN
     PRINT "Prin1 Paid: "; TaxTrans(1).Revenue.Principle1Pd
     PRINT "Prin2 Paid: "; TaxTrans(1).Revenue.Principle2Pd
     PRINT "Prin3 Paid: "; TaxTrans(1).Revenue.Principle3Pd
     PRINT "Prin4 Paid: "; TaxTrans(1).Revenue.Principle4Pd
     PRINT "Prin5 Paid: "; TaxTrans(1).Revenue.Principle5Pd
     PRINT "Intst Paid: "; TaxTrans(1).Revenue.InterestPd
     PRINT "Penty Paid: "; TaxTrans(1).Revenue.PenaltyPd
    END IF
    PrevTranRec& = TaxTrans(1).LastTrans
    SLEEP 5
  LOOP

  CLOSE

END

