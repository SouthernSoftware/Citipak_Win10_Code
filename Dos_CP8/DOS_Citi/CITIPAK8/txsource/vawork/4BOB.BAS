
  'Chilhowie Minimum Check Here
  IF INSTR(TownName$, "CHILHOWIE") THEN
    PersBalTest# = PersTaxDue# + MHTaxDue# + MCTaxDue# + FETaxDue# + MTTaxDue#
    CalcDiff# = 5 - PersBalTest#
    IF PersBalTest# > 0 AND PersBalTest# < 5 THEN
      TBillRec(1).ChillHowieFudge = CalcDiff#
      IF TBillRec(1).PPTRADiscnt > 0 AND PPTRAValue# <= 1000 THEN
        TBillRec(1).PPTRADiscnt = PersTaxDue# + CalcDiff#
      END IF
    END IF
  END IF

