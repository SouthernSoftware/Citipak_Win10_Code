      PayTranRec(1).Revenue.Principle1Pd = PayListRec(1).Principle1

'TaxType% = 1 is Real, 2 is Personal
      'IF TaxPaymentRec(1).BillType = "R" THEN
        PayTranRec(1).Revenue.Principle2Pd = PayListRec(1).Principle2
        PayTranRec(1).Revenue.Principle3Pd = PayListRec(1).Principle3
        PayTranRec(1).Revenue.Principle4Pd = PayListRec(1).Principle4
        PayTranRec(1).Revenue.Principle5Pd = PayListRec(1).Principle5
        PayTranRec(1).Revenue.InterestPd = PayListRec(1).Interest
        PayTranRec(1).Revenue.PenaltyPd = PayListRec(1).Penalty

