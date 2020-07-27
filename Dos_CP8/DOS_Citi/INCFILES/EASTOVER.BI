  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "Town of Eastover"
  PRINT #UBRpt, TAB(10); "500 Vanboklen Street"
  PRINT #UBRpt, TAB(10); "Eastover, SC 29044";
  PRINT #UBRpt, TAB(55); DATE$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); UBCustRec(1).CustName; TAB(55); "Acct # "; Acct$
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR1; TAB(55); "Pastdue: "; USING "$$#####.##"; CustBal#
  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(15); "Your water service will be discontinued for lack of "
  PRINT #UBRpt, TAB(10); "payment on "; LEFT$(PDate$, 2); "-"; MID$(PDate$, 4, 2); "-"; RIGHT$(PDate$, 2); ", at 12 o'clock unless arrangements"
  PRINT #UBRpt, TAB(10); "have been made with this office before this date. If service"
  PRINT #UBRpt, TAB(10); "is discontinued a $25.00 cut on fee will be due plus the "
  PRINT #UBRpt, TAB(10); "past due amount shown above."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(15); "The town office hours are from 9:00 A.M. until 5:00 P.M."
  PRINT #UBRpt, TAB(10); "Monday through Friday. Phone number is (804)-848-2414."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(15); "If you have already mailed your payment please disregard"
  PRINT #UBRpt, TAB(10); "this notice."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(40); "Water Department"
  PRINT #UBRpt, TAB(40); "Town of Eastover"
  PRINT #UBRpt,
  PRINT #UBRpt, CHR$(12);


