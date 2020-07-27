'add some kind of an include section here for various state calc's
'look at employer's state in the controll file to determin which
'state tax tables to use or edit

'Alabama 1/8/97

  AL20PctDed# = Round(AnnualizedStaGross# * .2)
  FedTaxDed# = TransRec.FedTaxAmt * PayFreq

  'STOP

  SELECT CASE QPTrim$(EMP2Rec(1).EMPSTAS)

    CASE "S", ""
      IF EMP2Rec(1).EMPSTAA < 0 THEN EMP2Rec(1).EMPSTAA = 0

      '--new for al state calc
      IF AL20PctDed# >= 2000 THEN AL20PctDed# = 2000
      AnnualizedStaGross# = AnnualizedStaGross# - AL20PctDed# - FedTaxDed#
      '--end of al specific

      TaxableAmtSta# = Round#(AnnualizedStaGross# - (EMP2Rec(1).EMPSTAA * StateTax(1).TAX106))
      TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).TAX103)
      IF TaxableAmtSta# <= 0 THEN
        TransRec.StaTaxAmt = 0
      ELSEIF TaxableAmtSta# < StateTax(1).STS(3, 1) THEN
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 1) + (TaxableAmtSta# * (StateTax(1).STS(2, 1) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STS(3, 1) AND TaxableAmtSta# < StateTax(1).STS(3, 2) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 1))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 2) + (TaxableAmtSta# * (StateTax(1).STS(2, 2) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STS(3, 2) AND TaxableAmtSta# < StateTax(1).STS(3, 3) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 2))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 3) + (TaxableAmtSta# * (StateTax(1).STS(2, 3) * .01))) / PayFreq)
      ELSE
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 3))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 4) + (TaxableAmtSta# * (StateTax(1).STS(2, 4) * .01))) / PayFreq)
      END IF

    CASE "M"
      IF EMP2Rec(1).EMPSTAA < 0 THEN EMP2Rec(1).EMPSTAA = 0

      '--new for al state calc
      IF AL20PctDed# >= 4000 THEN AL20PctDed# = 4000
      AnnualizedStaGross# = AnnualizedStaGross# - AL20PctDed# - FedTaxDed#
      '--end of al specific
      
      TaxableAmtSta# = Round#(AnnualizedStaGross# - (EMP2Rec(1).EMPSTAA * StateTax(1).TAX206))
      TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).TAX203)
      IF TaxableAmtSta# <= 0 THEN
        TransRec.StaTaxAmt = 0
      ELSEIF TaxableAmtSta# < StateTax(1).STM(3, 1) THEN
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STM(1, 1) + (TaxableAmtSta# * (StateTax(1).STM(2, 1) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STM(3, 1) AND TaxableAmtSta# < StateTax(1).STM(3, 2) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STM(3, 1))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STM(1, 2) + (TaxableAmtSta# * (StateTax(1).STM(2, 2) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STM(3, 2) AND TaxableAmtSta# < StateTax(1).STM(3, 3) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STM(3, 2))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STM(1, 3) + (TaxableAmtSta# * (StateTax(1).STM(2, 3) * .01))) / PayFreq)
      ELSE
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STM(3, 3))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STM(1, 4) + (TaxableAmtSta# * (StateTax(1).STM(2, 4) * .01))) / PayFreq)
      END IF

    CASE "H"
      IF EMP2Rec(1).EMPSTAA < 0 THEN EMP2Rec(1).EMPSTAA = 0

      '--new for al state calc
      IF AL20PctDed# >= 2000 THEN AL20PctDed# = 2000
      AnnualizedStaGross# = AnnualizedStaGross# - AL20PctDed# - FedTaxDed#
      '--end of al specific
      
      TaxableAmtSta# = Round#(AnnualizedStaGross# - (EMP2Rec(1).EMPSTAA * StateTax(1).TAX306))
      TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).TAX303)
      IF TaxableAmtSta# <= 0 THEN
        TransRec.StaTaxAmt = 0
      ELSEIF TaxableAmtSta# < StateTax(1).STH(3, 1) THEN
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STH(1, 1) + (TaxableAmtSta# * (StateTax(1).STH(2, 1) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STH(3, 1) AND TaxableAmtSta# < StateTax(1).STH(3, 2) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STH(3, 1))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STH(1, 2) + (TaxableAmtSta# * (StateTax(1).STH(2, 2) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STS(3, 2) AND TaxableAmtSta# < StateTax(1).STS(3, 3) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 2))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 3) + (TaxableAmtSta# * (StateTax(1).STS(2, 3) * .01))) / PayFreq)
      ELSE
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 3))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 4) + (TaxableAmtSta# * (StateTax(1).STS(2, 4) * .01))) / PayFreq)
      END IF
  END SELECT

  IF EMP2Rec(1).EMPSTAAA > 0 THEN
    TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt + EMP2Rec(1).EMPSTAAA)
  END IF

