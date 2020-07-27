'add some kind of an include section here for various state calc's
'look at employer's state in the controll file to determin which
'state tax tables to use or edit

'Arkansas 1/21/97

  IF EMP2Rec(1).EMPSTAA < 0 THEN EMP2Rec(1).EMPSTAA = 0

  SELECT CASE QPTrim$(EMP2Rec(1).EMPSTAS)

    CASE "S", ""
      
      'AR10PctDed# = Round(AnnualizedStaGross# * StateTax(1).TAX105)

      'IF AR10PctDed# > StateTax(1).TAX103 THEN
        AR10PctDed# = StateTax(1).TAX103
      'END IF

      AnnualizedStaGross# = AnnualizedStaGross# - AR10PctDed#

      TaxableAmtSta# = Round#(AnnualizedStaGross#)

      IF TaxableAmtSta# <= 0 THEN
        TransRec.StaTaxAmt = 0
      ELSEIF TaxableAmtSta# < StateTax(1).STS(3, 1) THEN
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 1) + (TaxableAmtSta# * (StateTax(1).STS(2, 1) * .01))))

      ELSEIF TaxableAmtSta# > StateTax(1).STS(3, 1) AND TaxableAmtSta# < StateTax(1).STS(3, 2) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 1))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 2) + (TaxableAmtSta# * (StateTax(1).STS(2, 2) * .01))))

      ELSEIF TaxableAmtSta# > StateTax(1).STS(3, 2) AND TaxableAmtSta# < StateTax(1).STS(3, 3) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 2))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 3) + (TaxableAmtSta# * (StateTax(1).STS(2, 3) * .01))))

      ELSEIF TaxableAmtSta# > StateTax(1).STS(3, 3) AND TaxableAmtSta# < StateTax(1).STS(3, 4) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STH(3, 3))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 4) + (TaxableAmtSta# * (StateTax(1).STS(2, 4) * .01))))

      ELSEIF TaxableAmtSta# > StateTax(1).STS(3, 4) AND TaxableAmtSta# < StateTax(1).STS(3, 5) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 4))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 5) + (TaxableAmtSta# * (StateTax(1).STS(2, 5) * .01))))

      ELSE
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STS(3, 5))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STS(1, 6) + (TaxableAmtSta# * (StateTax(1).STS(2, 6) * .01))))
      END IF
'STOP
      'TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt - StateTax(1).TAX106)
      TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt - (EMP2Rec(1).EMPSTAA * StateTax(1).TAX106))
      TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt / PayFreq)

   CASE "M"

      'AR10PctDed# = Round(AnnualizedStaGross# / StateTax(1).TAX205)
      'IF AR10PctDed# > StateTax(1).TAX203 THEN
        AR10PctDed# = StateTax(1).TAX203
      'END IF

      AnnualizedStaGross# = AnnualizedStaGross# - AR10PctDed#
      TaxableAmtSta# = Round#(AnnualizedStaGross#)

      IF TaxableAmtSta# <= 0 THEN
        TransRec.StaTaxAmt = 0

      ELSEIF TaxableAmtSta# < StateTax(1).STM(3, 1) THEN
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STM(1, 1) + (TaxableAmtSta# * (StateTax(1).STM(2, 1) * .01))))

      ELSEIF TaxableAmtSta# > StateTax(1).STM(3, 1) AND TaxableAmtSta# < StateTax(1).STM(3, 2) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STM(3, 1))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STM(1, 2) + (TaxableAmtSta# * (StateTax(1).STM(2, 2) * .01))))

      ELSEIF TaxableAmtSta# > StateTax(1).STM(3, 2) AND TaxableAmtSta# < StateTax(1).STM(3, 3) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STM(3, 2))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STM(1, 3) + (TaxableAmtSta# * (StateTax(1).STM(2, 3) * .01))))

      ELSEIF TaxableAmtSta# > StateTax(1).STM(3, 3) AND TaxableAmtSta# < StateTax(1).STM(3, 4) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STM(3, 3))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STM(1, 4) + (TaxableAmtSta# * (StateTax(1).STM(2, 4) * .01))))

      ELSEIF TaxableAmtSta# > StateTax(1).STM(3, 4) AND TaxableAmtSta# < StateTax(1).STM(3, 5) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STM(3, 4))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STM(1, 5) + (TaxableAmtSta# * (StateTax(1).STM(2, 5) * .01))))

      ELSE
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STM(3, 5))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STM(1, 6) + (TaxableAmtSta# * (StateTax(1).STM(2, 6) * .01))))
      END IF

      'TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt - StateTax(1).TAX206)
      'TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt - (EMP2Rec(1).EMPSTAA * StateTax(1).TAX207))
      TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt - (EMP2Rec(1).EMPSTAA * StateTax(1).TAX206))
      TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt / PayFreq)

    CASE "H"

      'AR10PctDed# = Round(AnnualizedStaGross# * StateTax(1).TAX305)
      'IF AR10PctDed# > StateTax(1).TAX303 THEN
        AR10PctDed# = StateTax(1).TAX303
      'END IF

      AnnualizedStaGross# = AnnualizedStaGross# - AR10PctDed#
      TaxableAmtSta# = Round#(AnnualizedStaGross#)

      IF TaxableAmtSta# <= 0 THEN
        TransRec.StaTaxAmt = 0

      ELSEIF TaxableAmtSta# < StateTax(1).STH(3, 1) THEN
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STH(1, 1) + (TaxableAmtSta# * (StateTax(1).STH(2, 1) * .01))))

      ELSEIF TaxableAmtSta# > StateTax(1).STH(3, 1) AND TaxableAmtSta# < StateTax(1).STH(3, 2) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STH(3, 1))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STH(1, 2) + (TaxableAmtSta# * (StateTax(1).STH(2, 2) * .01))))

      ELSEIF TaxableAmtSta# > StateTax(1).STH(3, 2) AND TaxableAmtSta# < StateTax(1).STS(3, 3) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STH(3, 2))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STH(1, 3) + (TaxableAmtSta# * (StateTax(1).STH(2, 3) * .01))))

      ELSEIF TaxableAmtSta# > StateTax(1).STH(3, 3) AND TaxableAmtSta# < StateTax(1).STH(3, 4) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STH(3, 3))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STH(1, 4) + (TaxableAmtSta# * (StateTax(1).STH(2, 4) * .01))))

      ELSEIF TaxableAmtSta# > StateTax(1).STH(3, 4) AND TaxableAmtSta# < StateTax(1).STH(3, 5) THEN
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STH(3, 4))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STH(1, 5) + (TaxableAmtSta# * (StateTax(1).STH(2, 5) * .01))))

      ELSE
        TaxableAmtSta# = Round#(TaxableAmtSta# - StateTax(1).STH(3, 5))
        TransRec.StaTaxAmt = Round#(Round#(StateTax(1).STH(1, 6) + (TaxableAmtSta# * (StateTax(1).STH(2, 6) * .01))))
      END IF

      'TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt - StateTax(1).TAX306)
      'TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt - (EMP2Rec(1).EMPSTAA * StateTax(1).TAX307))
      TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt - (EMP2Rec(1).EMPSTAA * StateTax(1).TAX306))
      TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt / PayFreq)

  END SELECT

  'deduct tax credit from state tax amt.
  IF TransRec.StaTaxAmt < 0 THEN
    TransRec.StaTaxAmt = 0
  END IF


  IF EMP2Rec(1).EMPSTAAA > 0 THEN
    TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt + EMP2Rec(1).EMPSTAAA)
  END IF

  TaxableAmtSta# = TransRec.StaTaxAmt

'
'  END SELECT

