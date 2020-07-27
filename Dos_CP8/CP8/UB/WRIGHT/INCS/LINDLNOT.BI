PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(33); "Town of Linden"
PRINT #UBRpt, TAB(33); "P.O. Drawer 61"
PRINT #UBRpt, TAB(32); "Linden, NC  28356"
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); UBCustRec(1).CustName; TAB(58); USING "Acct: ##### "; AcctNo&
PRINT #UBRpt, TAB(10); UBCustRec(1).AddR1
PRINT #UBRpt, TAB(10); UBCustRec(1).AddR2
PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).City); " "; UBCustRec(1).State; "  "; Zip$; TAB(55); USING "Pastdue: $$#####.##"; CustBal#
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); LongNDate$
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(15); "Your water service has been discontinued for lack of"
PRINT #UBRpt, TAB(10); "payment.  A $25.00 cut on fee is due plus the past due"
PRINT #UBRpt, TAB(10); "amount shown above."
PRINT #UBRpt,
PRINT #UBRpt, TAB(15); "The Town Office hours are from 8:00 AM til 5:00 PM"
PRINT #UBRpt, TAB(10); "Monday through Friday.  Also, there is a depository box between"
PRINT #UBRpt, TAB(10); "the two front doors.  It will be necessary to have the payment"
PRINT #UBRpt, TAB(10); "in the office no later than 3:30 PM to have water services"
PRINT #UBRpt, TAB(10); "restored that day."
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); "Water Department"
PRINT #UBRpt, TAB(10); "TOWN OF LINDEN"
PRINT #UBRpt,
PRINT #UBRpt, CHR$(12);

