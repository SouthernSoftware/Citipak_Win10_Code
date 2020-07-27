FixCashChkFlds:

  SELECT CASE QPTrim$(Form$(TenderFld, 0))
  CASE Choice$(1, 0)            'CASH               this sets the cash or check
    Fld(TenderFld + 1).Protected = False        'amount fields protected or
    Fld(TenderFld + 2).Protected = True         'unprotected based of the
    LSET Form$(ChkAmtFld, 0) = "0"
  CASE Choice$(2, 0)            'CHECK              tender type field selection
    Fld(TenderFld + 1).Protected = True
    Fld(TenderFld + 2).Protected = False
    LSET Form$(CashAmtFld, 0) = "0"
  CASE ELSE     'BOTH
    Fld(TenderFld + 1).Protected = False
    Fld(TenderFld + 2).Protected = False
  END SELECT

  SaveField ChkAmtFld, Form$(), Fld(), BadField
  SaveField CashAmtFld, Form$(), Fld(), BadField

  CalcFields 0, AmtPadFlds(1), Form$(), Fld()
  CalcFields 0, CashAmtFld, Form$(), Fld()

  PrintArray 1, NumFlds - 1, Form$(), Fld()

  'GOSUB AutoDistribute:
RETURN

