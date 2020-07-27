'Cleveland
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(5); LongNDate$
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    "; UBCustRec(1).CustName
  PRINT #UBRpt, "    "; UBCustRec(1).ADDR1
  PRINT #UBRpt, "    "; UBCustRec(1).ADDR2
  PRINT #UBRpt, "    "; QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    The water account located at "; QPTrim$(UBCustRec(1).ServAddr); ", is scheduled "
  PRINT #UBRpt, "    to be shut off at 8:00 a.m. on "; LongPDate$
  PRINT #UBRpt, "    if not paid in full before the above date.  If you have already"
  PRINT #UBRpt, "    sent your payment, Thank You, and please disregard this notice."
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    Reminder:"
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    Bills are due on, or around the 10th of each month. Any unpaid"
  PRINT #UBRpt, "    balance after this date, will be assessed a 10% late fee. With"
  PRINT #UBRpt, "    a minimum of $1.50 on any unpaid balance."
  PRINT #UBRpt, ""
  PRINT #UBRpt,
  PRINT #UBRpt, "    ACCOUNT #"; Acct$; TAB(40); "AMOUNT DUE"; USING "$$#####,#.##"; CustBal#
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "    Lewis D. Ewry"
  PRINT #UBRpt, "    Village Administrator"
  PRINT #UBRpt, "    (937)365-0104(h)"
  PRINT #UBRpt, "    634-2223(Office)"
  PRINT #UBRpt, "    634-3723(Water plant)"
  PRINT #UBRpt, CHR$(12)

