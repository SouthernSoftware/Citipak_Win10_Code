

    IF PickedCnt > 0 THEN
      IF GotTaxFile THEN
        TInvAmt# = 0
        FOR zz = 1 TO DistCnt
          TInvAmt# = Round#(TInvAmt# + DistSum(zz).DistAmt)
        NEXT
        IF StaTaxFlag THEN
          IF STAcctRec > 0 THEN
            ActualCnt = ActualCnt + 1
            REDIM ActualDist(1 TO ActualCnt)  AS DistSumType
            ActualDist(TState).AcctTitle = GetAcctTitle$(STAcctRec)
            ActualDist(TState).DistAcctNum = InvTaxRec(1).InvTax(1).AcctNo
            DistSum(TState).AcctTitle = GetAcctTitle$(STAcctRec)
            DistSum(TState).DistAcctNum = InvTaxRec(1).InvTax(1).AcctNo
            TCalc# = InvTaxRec(1).InvTax(1).TaxAmt
            IF TCalc# > 0 THEN
              StaTax# = Round#(TInvAmt# * (TCalc# / 100))
              ActualDist(TState).DistAmt = StaTax#
            END IF
          END IF
        END IF
        IF CtyTaxFlag THEN
          CTAcctRec = FindAcct(QPTrim$(InvTaxRec(1).InvTax(2).AcctNo))
          IF CTAcctRec > 0 THEN
            ActualCnt = ActualCnt + 1
            REDIM ActualDist(1 TO ActualCnt)  AS DistSumType
            ActualDist(TCounty).AcctTitle = GetAcctTitle$(CTAcctRec)
            ActualDist(TCounty).DistAcctNum = InvTaxRec(1).InvTax(2).AcctNo
            DistSum(TCounty).AcctTitle = GetAcctTitle$(CTAcctRec)
            DistSum(TCounty).DistAcctNum = InvTaxRec(1).InvTax(2).AcctNo

            TCalc# = InvTaxRec(1).InvTax(2).TaxAmt
            IF TCalc# > 0 THEN
              CtyTax# = Round#(TInvAmt# * (TCalc# / 100))
              ActualDist(TCounty).DistAmt = CtyTax#
            END IF
          END IF
        END IF
      END IF
    END IF

    IF DistCnt > 2 THEN
      SortT DistSum(3), DistCnt - 2, 0, 44, 0, 16
    ELSE
      GOTO Skip2Here:
    END IF

    IF GotTaxFile AND PickedCnt > 0 THEN
      ActualCnt = TaxActualAdj
      FirstAcct$ = QPTrim$(DistSum(ActualCnt).DistAcctNum)
    ELSE
      ActualCnt = 1
    END IF

    REDIM ActualDist(1 TO ActualCnt)  AS DistSumType
    FirstAcct$ = QPTrim$(DistSum(ActualCnt).DistAcctNum)
    ActualDist(1).DistAcctNum = DistSum(ActualCnt).DistAcctNum
    ActualDist(1).AcctTitle = DistSum(ActualCnt).AcctTitle
    ActualDist(1).DistAmt = DistSum(ActualCnt).DistAmt


    FOR zz = ActualCnt + 1 TO DistCnt
      ThisAcct$ = QPTrim$(DistSum(zz).DistAcctNum)
      IF ThisAcct$ = FirstAcct$ THEN
        ActualDist(ActualCnt).DistAmt = Round#(ActualDist(ActualCnt).DistAmt + DistSum(zz).DistAmt)
      ELSE
        ActualCnt = ActualCnt + 1
        REDIM PRESERVE ActualDist(1 TO ActualCnt) AS DistSumType
        ActualDist(ActualCnt).DistAcctNum = DistSum(zz).DistAcctNum
        ActualDist(ActualCnt).AcctTitle = DistSum(zz).AcctTitle
        ActualDist(ActualCnt).DistAmt = DistSum(zz).DistAmt
        FirstAcct$ = QPTrim$(DistSum(zz).DistAcctNum)
      END IF
    NEXT

    FOR zz = 1 TO ActualCnt
      StartFld = AcctFlds(zz)
      LSET Form$(StartFld, 0) = ActualDist(zz).DistAcctNum
      LSET Form$(StartFld + 1, 0) = ActualDist(zz).AcctTitle
      LSET Form$(StartFld + 3, 0) = STR$(ActualDist(zz).DistAmt)
      TotalAmt# = Round#(TotalAmt# + ActualDist(zz).DistAmt)
      FOR SAVFld = StartFld TO StartFld + 3
        SaveField SAVFld, Form$(), Fld(), BadFld
      NEXT
    NEXT

Skip2Here:
    IF PickedCnt > 1 THEN
      LSET Form$(PONumFld, 0) = "Multi"
    ELSEIF PickedCnt = 1 THEN
      LSET Form$(PONumFld, 0) = QPTrim$(LEFT$(MTChoice(Picked(1)).V, 13))
    ELSE
      LSET Form$(PONumFld, 0) = ""
    END IF

    FOR zz = PickedCnt + 1 TO 6
      LSET Form$(POFld(zz), 0) = MKL$(0)
      SaveField POFld(zz), Form$(), Fld(), BadFld
    NEXT

    LSET Form$(7, 0) = QPTrim$(STR$(TotalAmt#))
    SaveField 7, Form$(), Fld(), BadFld

    IF ActualCnt < NumAcctFlds THEN
      FOR CDCnt = AcctFlds(ActualCnt + 1) TO AcctFlds(NumAcctFlds)                'clear all of distrubtions
        LSET Form$(CDCnt, 0) = ""               'fields
        LSET Form$(CDCnt + 1, 0) = ""
        LSET Form$(CDCnt + 2, 0) = ""
        LSET Form$(CDCnt + 3, 0) = ""
        FOR SAVFld = CDCnt TO CDCnt + 3
          SaveField SAVFld, Form$(), Fld(), BadFld
        NEXT
      NEXT
    END IF
  END IF

