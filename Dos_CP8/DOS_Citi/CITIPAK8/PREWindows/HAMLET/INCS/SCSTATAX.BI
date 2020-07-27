

      IF EMP2Rec(1).EMPSTAA < 0 THEN EMP2Rec(1).EMPSTAA = 0

      TaxableAmtSta# = Round#(AnnualizedStaGross# - (EMP2Rec(1).EMPSTAA * StateTax(1).TAX106))

      IF EMP2Rec(1).EMPSTAA > 0 THEN
        SCStateExmpAmt# = Round#(AnnualizedStaGross# * .1)
        IF SCStateExmpAmt# > StateTax(1).TAX103 THEN
          SCStateExmpAmt# = StateTax(1).TAX103
        END IF
        TaxableAmtSta# = Round#(TaxableAmtSta# - SCStateExmpAmt#)
      END IF

      IF TaxableAmtSta# <= 0 THEN
        TransRec.StaTaxAmt = 0
      ELSEIF TaxableAmtSta# < StateTax(1).STS(3, 1) THEN  '1
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 1) + (TaxableAmtSta# * (StateTax(1).STS(2, 1) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STS(3, 1) AND TaxableAmtSta# < StateTax(1).STS(3, 2) THEN '2
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 1))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 2) + (TaxableAmtSta# * (StateTax(1).STS(2, 2) * .01))) / PayFreq)'3
      ELSEIF TaxableAmtSta# > StateTax(1).STS(3, 2) AND TaxableAmtSta# < StateTax(1).STS(3, 3) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 2))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 3) + (TaxableAmtSta# * (StateTax(1).STS(2, 3) * .01))) / PayFreq) '4

'******
      ELSEIF TaxableAmtSta# > StateTax(1).STS(3, 3) AND TaxableAmtSta# < StateTax(1).STS(3, 4) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 3))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 4) + (TaxableAmtSta# * (StateTax(1).STS(2, 4) * .01))) / PayFreq) '5

      ELSEIF TaxableAmtSta# > StateTax(1).STS(3, 4) AND TaxableAmtSta# < StateTax(1).STS(3, 5) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 4))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 5) + (TaxableAmtSta# * (StateTax(1).STS(2, 5) * .01))) / PayFreq) '6
'******
      ELSE
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 5))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 6) + (TaxableAmtSta# * (StateTax(1).STS(2, 6) * .01))) / PayFreq)
      END IF

  IF EMP2Rec(1).EMPSTAAA > 0 THEN
    TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt + EMP2Rec(1).EMPSTAAA)
  END IF

