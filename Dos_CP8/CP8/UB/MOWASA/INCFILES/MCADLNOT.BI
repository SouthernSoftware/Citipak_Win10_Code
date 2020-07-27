'
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "                                 Town of Mcadenville"
  PRINT #UBRpt, "                                     P.O. Box 9"
  PRINT #UBRpt, "                                   125 Main Street"
  PRINT #UBRpt, "                                Mcadenville, NC 28101"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).CustName); TAB(51); "   Account: "; Acct$
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR1; TAB(51); "  Location: "; Location$
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR2; TAB(51); "Amount Due:"; FUsing$(STR$(TotalBal#), "$$,######.##")
  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; Zip$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "Dear Customer:"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     Your water service will be discontinued for lack of"
  PRINT #UBRpt, TAB(10); "payment effective "; LongPDate$; "."
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); "     It will be necessary to have the payment in the Town"
  PRINT #UBRpt, TAB(10); "Office by "; PDate$; " to avoid this scheduled CUT-OFF!"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     If payment is not received by the above date, A fee"
  PRINT #UBRpt, TAB(10); "of $100.00 in addition to the Past Due amount is required"
  PRINT #UBRpt, TAB(10); "before service will be restored!"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     Town Office hours are: 7:30 a.m. to 12:00 Noon and"
  PRINT #UBRpt, TAB(10); "1:00 p.m. to 3:00 p.m. Monday through Friday. For your "
  PRINT #UBRpt, TAB(10); "convenience, there is a mail slot at the door of the Town"
  PRINT #UBRpt, TAB(10); "Office."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     If you have any questions, please contact us at Town"
  PRINT #UBRpt, TAB(10); "Office. Phone: 824-3190"
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); "     If payment has been made prior to receiving this notice,"
  PRINT #UBRpt, TAB(10); "please accept our thanks and disregard this notice."
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); "Thank you,"
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); "     Town of Mcadenville"
  PRINT #UBRpt, CHR$(12)


