      CMTRRec(1).TransRevAmt(1) = Round#(CMTRRec(1).TransRevAmt(1) + PayListRec(1).Principle1)
      CMTRRec(1).TransRevAmt(2) = Round#(CMTRRec(1).TransRevAmt(2) + PayListRec(1).Interest1)
      CMTRRec(1).TransRevAmt(3) = Round#(CMTRRec(1).TransRevAmt(3) + PayListRec(1).Collection)

      PUT #TranFile, PayListRec(1).BillRec, TaxTranRec(1)
      NextTransRec& = (LOF(TranFile) \ TaxTranRecLen) + 1
      PUT #TranFile, NextTransRec&, PayTranRec(1)
      TaxCustRec(1).LastTrans = NextTransRec&
      PUT #CustFile, TaxPaymentRec(1).CustAcct, TaxCustRec(1)
SkipThisRec:
      ThisListRec& = PayListRec(1).PrevListRec
    LOOP
    CMTRRec(1).TransDate = TaxPaymentRec(1).PayDate
    CMTRRec(1).TransAmount = TaxPaymentRec(1).AmtRecd
    CMTRRec(1).TransAmtOwed = TaxPaymentRec(1).AmtOwed
    CMTRRec(1).TransCash = TaxPaymentRec(1).CashAmt
    CMTRRec(1).TransCheck = TaxPaymentRec(1).ChkAmt
    CMTRRec(1).TransDesc = "Tax Payment"
    CMTRRec(1).TransSource = 33
    CMTRRec(1).TransName = TaxPaymentRec(1).CustName
    CMTRRec(1).TransAcctNum = TaxPaymentRec(1).CustAcct
    CMTRRec(1).TransDetNum = 0
    CMTRRec(1).TransOperNum = OperNum
    CMTRRec(1).TransPad = ""
    NextCMTransRec& = (LOF(CHandle) \ CMTrRecLen) + 1
    PUT #CHandle, NextCMTransRec&, CMTRRec(1)

