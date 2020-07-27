'add some kind of an include section here for various state calc's
'look at employer's state in the controll file to determin which
'state tax tables to use or edit
'  STOP

  IF EMP2Rec(1).EMPSTAA < 0 THEN EMP2Rec(1).EMPSTAA = 0

  SELECT CASE QPTrim$(EMP2Rec(1).EMPSTAS)
    CASE LEFT$(TaxText$(1), 1), ""
      IF AnnualizedStaGross# > StateTax(1).TAX101 THEN
        ExcessAmt# = AnnualizedStaGross# - StateTax(1).TAX101
        AnnualizedStaGross# = StateTax(1).TAX101
      END IF

      Exmp1Amt# = Round(EMP2Rec(1).EMPSTAA * StateTax(1).TAX102)
                    'number of allounces * std ded amt
      Exmp2Amt# = Round(AnnualizedStaGross# * (StateTax(1).TAX105 * .01))
                    'Annual gross wages * Std Ded Rate%

      SELECT CASE Exmp2Amt#
      CASE IS >= StateTax(1).TAX107
        'if exempt 2 amt is greater than max std amt
        Exmp2Amt# = StateTax(1).TAX107
      CASE IS < StateTax(1).TAX106
        'if exempt 2 amt is less than min std amt
        Exmp2Amt# = StateTax(1).TAX106
      END SELECT

      Exmp3Amt# = Round((AnnualizedStaGross# - (EMP2Rec(1).EMPSTAA * StateTax(1).TAX103) - StateTax(1).TAX104))

      Exmp3Amt# = Round(Exmp3Amt# * (StateTax(1).TAX105 * .01))

      IF Exmp3Amt# < 0 THEN Exmp3Amt# = 0

      TotalExmp# = Round(Exmp1Amt# + Exmp2Amt# + Exmp3Amt#)

      TaxableAmtSta# = Round#(AnnualizedStaGross# - TotalExmp#)

      IF TaxableAmtSta# < 0 THEN
        TaxableAmtSta# = 0
      END IF
STOP

      FOR TCnt = 1 TO 12
        IF TaxableAmtSta# < StateTax(1).STS(3, TCnt) THEN
          EXIT FOR
        END IF
      NEXT

      TblPoint = TCnt

      IF TblPoint = 1 THEN
        TransRec.StaTaxAmt = Round#(TaxableAmtSta# * (StateTax(1).STS(2, TblPoint) * .01) / PayFreq)
      ELSE
        TransRec.StaTaxAmt = Round#((StateTax(1).STS(1, TCnt) + (TaxableAmtSta# - StateTax(1).STS(3, TCnt - 1)) * (StateTax(1).STS(2, TCnt) * .01)) / PayFreq)
      END IF

      IF ExcessAmt# > 0 THEN
        FOR RCnt = 12 TO 1 STEP -1
          IF StateTax(1).STS(2, RCnt) > 0 THEN
            EXIT FOR
          END IF
        NEXT
        Multi# = Round#(StateTax(1).STS(2, RCnt) * .01)
        'TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt + (ExcessAmt# * Multi#))
        TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt + ((ExcessAmt# * Multi#) / PayFreq))
      END IF

    CASE LEFT$(TaxText$(2), 1)
     ' STOP
      IF AnnualizedStaGross# > StateTax(1).TAX201 THEN
        ExcessAmt# = AnnualizedStaGross# - StateTax(1).TAX201
        AnnualizedStaGross# = StateTax(1).TAX201
      END IF

      Exmp1Amt# = Round(EMP2Rec(1).EMPSTAA * StateTax(1).TAX202)
                    'number of allounces * std ded amt
      Exmp2Amt# = Round(AnnualizedStaGross# * (StateTax(1).TAX205 * .01))
                    'Annual gross wages * Std Ded Rate%

      SELECT CASE Exmp2Amt#
      CASE IS >= StateTax(1).TAX207
        'if exempt 2 amt is greater than max std amt
        Exmp2Amt# = StateTax(1).TAX207
      CASE IS < StateTax(1).TAX206
        'if exempt 2 amt is less than min std amt
        Exmp2Amt# = StateTax(1).TAX206
      END SELECT

      Exmp3Amt# = Round((AnnualizedStaGross# - (EMP2Rec(1).EMPSTAA * StateTax(1).TAX203) - StateTax(1).TAX204))

      Exmp3Amt# = Round(Exmp3Amt# * (StateTax(1).TAX205 * .01))

      IF Exmp3Amt# < 0 THEN Exmp3Amt# = 0

      TotalExmp# = Round(Exmp1Amt# + Exmp2Amt# + Exmp3Amt#)

      TaxableAmtSta# = Round#(AnnualizedStaGross# - TotalExmp#)

      IF TaxableAmtSta# < 0 THEN
        TaxableAmtSta# = 0
      END IF

      FOR TCnt = 1 TO 12
        IF TaxableAmtSta# < StateTax(1).STM(3, TCnt) THEN
          EXIT FOR
        END IF
      NEXT

      TblPoint = TCnt

      IF TblPoint = 1 THEN
        TransRec.StaTaxAmt = Round#(TaxableAmtSta# * (StateTax(1).STM(2, TblPoint) * .01) / PayFreq)
      ELSE
        TransRec.StaTaxAmt = Round#((StateTax(1).STM(1, TCnt) + (TaxableAmtSta# - StateTax(1).STM(3, TCnt - 1)) * (StateTax(1).STM(2, TCnt) * .01)) / PayFreq)
      END IF

      IF ExcessAmt# > 0 THEN
        FOR RCnt = 12 TO 1 STEP -1
          IF StateTax(1).STM(2, RCnt) > 0 THEN
            EXIT FOR
          END IF
        NEXT
        Multi# = Round#(StateTax(1).STM(2, RCnt) * .01)
        TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt + ((ExcessAmt# * Multi#) / PayFreq))
      END IF

    CASE LEFT$(TaxText$(3), 1)

      IF AnnualizedStaGross# > StateTax(1).TAX301 THEN
        ExcessAmt# = AnnualizedStaGross# - StateTax(1).TAX301
        AnnualizedStaGross# = StateTax(1).TAX301
      END IF

      Exmp1Amt# = Round(EMP2Rec(1).EMPSTAA * StateTax(1).TAX302)
                    'number of allounces * std ded amt
      Exmp2Amt# = Round(AnnualizedStaGross# * (StateTax(1).TAX305 * .01))
                    'Annual gross wages * Std Ded Rate%

      SELECT CASE Exmp2Amt#
      CASE IS >= StateTax(1).TAX307
        'if exempt 2 amt is greater than max std amt
        Exmp2Amt# = StateTax(1).TAX307
      CASE IS < StateTax(1).TAX306
        'if exempt 2 amt is less than min std amt
        Exmp2Amt# = StateTax(1).TAX306
      END SELECT

      Exmp3Amt# = Round((AnnualizedStaGross# - (EMP2Rec(1).EMPSTAA * StateTax(1).TAX303) - StateTax(1).TAX304))

      Exmp3Amt# = Round(Exmp3Amt# * (StateTax(1).TAX305 * .01))

      IF Exmp3Amt# < 0 THEN Exmp3Amt# = 0

      TotalExmp# = Round(Exmp1Amt# + Exmp2Amt# + Exmp3Amt#)

      TaxableAmtSta# = Round#(AnnualizedStaGross# - TotalExmp#)

      IF TaxableAmtSta# < 0 THEN
        TaxableAmtSta# = 0
      END IF

      FOR TCnt = 1 TO 12
        IF TaxableAmtSta# < StateTax(1).STH(3, TCnt) THEN
          EXIT FOR
        END IF
      NEXT

      TblPoint = TCnt

      IF TblPoint = 1 THEN
        TransRec.StaTaxAmt = Round#(TaxableAmtSta# * (StateTax(1).STH(2, TblPoint) * .01) / PayFreq)
      ELSE
        TransRec.StaTaxAmt = Round#((StateTax(1).STH(1, TCnt) + (TaxableAmtSta# - StateTax(1).STH(3, TCnt - 1)) * (StateTax(1).STH(2, TCnt) * .01)) / PayFreq)
      END IF

  '    STOP

      IF ExcessAmt# > 0 THEN
        FOR RCnt = 12 TO 1 STEP -1
          IF StateTax(1).STH(2, RCnt) > 0 THEN
            EXIT FOR
          END IF
        NEXT
        Multi# = Round#(StateTax(1).STH(2, RCnt) * .01)
        'TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt + (ExcessAmt# * Multi#))
        TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt + ((ExcessAmt# * Multi#) / PayFreq))
      END IF

  END SELECT


  IF EMP2Rec(1).EMPSTAAA > 0 AND PriorStaTax# = 0 THEN
    TransRec.StaTaxAmt = Round#(TransRec.StaTaxAmt + EMP2Rec(1).EMPSTAAA)
  END IF

