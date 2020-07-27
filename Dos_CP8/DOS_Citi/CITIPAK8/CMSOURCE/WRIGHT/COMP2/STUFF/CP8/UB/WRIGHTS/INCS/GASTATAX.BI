'--Georgia state tax calculation

  TaxableAmtSta# = Round#(AnnualizedStaGross#)
  IF EMP2Rec(1).EMPSTAA < 0 THEN EMP2Rec(1).EMPSTAA = 0
  
  SELECT CASE QPTrim$(EMP2Rec(1).EMPSTAS)
   
    CASE "F"  'Table F - Married one income

      '--Get Standard Deduction
      TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).TAX101)

      '--Get personal and dependendant deductions if claiming at least one allowance
      IF EMP2Rec(1).EMPSTAA > 0 THEN
        '--Personal Allowance Amount
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).TAX102)
        '--Dependent Exemption
        TaxableAmtSta# = Round#(TaxableAmtSta# - (EMP2Rec(1).EMPSTAA * StateTax(1).TAX103))
      END IF

      IF TaxableAmtSta# <= 0 THEN
        TransRec.StaTaxAmt = 0
      ELSEIF TaxableAmtSta# <= StateTax(1).STS(3, 2) THEN
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 1) + (TaxableAmtSta# * (StateTax(1).STS(2, 1) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STS(3, 1) AND TaxableAmtSta# <= StateTax(1).STS(3, 2) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 1))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 1) + (TaxableAmtSta# * (StateTax(1).STS(2, 1) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STS(3, 2) AND TaxableAmtSta# <= StateTax(1).STS(3, 3) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 2))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 2) + (TaxableAmtSta# * (StateTax(1).STS(2, 2) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STS(3, 3) AND TaxableAmtSta# <= StateTax(1).STS(3, 4) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 3))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 3) + (TaxableAmtSta# * (StateTax(1).STS(2, 3) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STS(3, 4) AND TaxableAmtSta# <= StateTax(1).STS(3, 5) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 4))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 4) + (TaxableAmtSta# * (StateTax(1).STS(2, 4) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STS(3, 5) AND TaxableAmtSta# <= StateTax(1).STS(3, 6) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 5))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 5) + (TaxableAmtSta# * (StateTax(1).STS(2, 5) * .01))) / PayFreq)
      ELSE
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 6))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 6) + (TaxableAmtSta# * (StateTax(1).STS(2, 6) * .01))) / PayFreq)
      END IF
    
    
    CASE "G" 'Table G - Married Filing Joint (dual income)"

      '--Get Standard Deduction
      TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).TAX201)

      '--Get personal and dependendant deductions if claiming at least one allowance
      IF EMP2Rec(1).EMPSTAA > 0 THEN
        '--Personal Allowance Amount
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).TAX202)
        '--Dependent Exemption
        TaxableAmtSta# = Round#(TaxableAmtSta# - (EMP2Rec(1).EMPSTAA * StateTax(1).TAX203))
      END IF

      IF TaxableAmtSta# <= 0 THEN
        TransRec.StaTaxAmt = 0
      ELSEIF TaxableAmtSta# <= StateTax(1).STM(3, 2) THEN
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STM(1, 1) + (TaxableAmtSta# * (StateTax(1).STM(2, 1) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STM(3, 1) AND TaxableAmtSta# <= StateTax(1).STM(3, 2) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STM(3, 1))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STM(1, 1) + (TaxableAmtSta# * (StateTax(1).STM(2, 1) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STM(3, 2) AND TaxableAmtSta# <= StateTax(1).STM(3, 3) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STM(3, 2))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STM(1, 2) + (TaxableAmtSta# * (StateTax(1).STM(2, 2) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STM(3, 3) AND TaxableAmtSta# <= StateTax(1).STM(3, 4) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STM(3, 3))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STM(1, 3) + (TaxableAmtSta# * (StateTax(1).STM(2, 3) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STM(3, 4) AND TaxableAmtSta# <= StateTax(1).STM(3, 5) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STM(3, 4))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STM(1, 4) + (TaxableAmtSta# * (StateTax(1).STM(2, 4) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STM(3, 5) AND TaxableAmtSta# <= StateTax(1).STM(3, 6) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STM(3, 5))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STM(1, 5) + (TaxableAmtSta# * (StateTax(1).STM(2, 5) * .01))) / PayFreq)
      ELSE
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STM(3, 6))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STM(1, 6) + (TaxableAmtSta# * (StateTax(1).STM(2, 6) * .01))) / PayFreq)
      END IF

    
    CASE "H", "" 'Georgia table H - Single Individual

      '--Get Standard Deduction
      TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).TAX301)

      '--Get personal and dependendant deductions if claiming at least one allowance
      IF EMP2Rec(1).EMPSTAA > 0 THEN
        '--Personal Allowance Amount
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).TAX302)
        '--Dependent Exemption
        TaxableAmtSta# = Round#(TaxableAmtSta# - (EMP2Rec(1).EMPSTAA * StateTax(1).TAX303))
      END IF

      IF TaxableAmtSta# <= 0 THEN
        TransRec.StaTaxAmt = 0
      ELSEIF TaxableAmtSta# <= StateTax(1).STH(3, 1) THEN
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STH(1, 1) + (TaxableAmtSta# * (StateTax(1).STH(2, 1) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STH(3, 1) AND TaxableAmtSta# <= StateTax(1).STH(3, 2) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STH(3, 1))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STH(1, 1) + (TaxableAmtSta# * (StateTax(1).STH(2, 1) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STH(3, 2) AND TaxableAmtSta# <= StateTax(1).STH(3, 3) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STH(3, 2))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STH(1, 2) + (TaxableAmtSta# * (StateTax(1).STH(2, 2) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STH(3, 3) AND TaxableAmtSta# <= StateTax(1).STH(3, 4) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STH(3, 3))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STH(1, 3) + (TaxableAmtSta# * (StateTax(1).STH(2, 3) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STH(3, 4) AND TaxableAmtSta# <= StateTax(1).STH(3, 5) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STH(3, 4))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STH(1, 4) + (TaxableAmtSta# * (StateTax(1).STH(2, 4) * .01))) / PayFreq)
      ELSEIF TaxableAmtSta# > StateTax(1).STH(3, 5) AND TaxableAmtSta# <= StateTax(1).STH(3, 6) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STH(3, 5))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STH(1, 5) + (TaxableAmtSta# * (StateTax(1).STH(2, 5) * .01))) / PayFreq)
      ELSE
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STH(3, 6))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STH(1, 6) + (TaxableAmtSta# * (StateTax(1).STH(2, 6) * .01))) / PayFreq)
      END IF
  
  END SELECT

  '--Add additional set withholding amount
  IF EMP2Rec(1).EMPSTAAA > 0 THEN
    TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt + EMP2Rec(1).EMPSTAAA)
  END IF

