
    IF UBCustRec(1).CASHONLY = "Y" THEN
      CashFlag = True
    END IF

SetTenderFlds:
  IF CashFlag THEN
    Choice$(1, 0) = "Cash"
    Choice$(2, 0) = ""
    Choice$(3, 0) = ""
  ELSE
    Choice$(1, 0) = "Cash"
    Choice$(2, 0) = "Check"
    Choice$(3, 0) = "Cash & Check"
  END IF
RETURN

