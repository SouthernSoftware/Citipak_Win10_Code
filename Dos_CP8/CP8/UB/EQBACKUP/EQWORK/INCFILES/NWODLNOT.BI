'
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "                                 TOWN OF NORWOOD"
  PRINT #UBRpt, "                              116 South Main Street"
  PRINT #UBRpt, "                                   P.O. BOX 697"
  PRINT #UBRpt, "                                NORWOOD NC, 28128"
  PRINT #UBRpt, "                                 (704) 474-3416"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).CustName); TAB(53); " Account# "; Acct$
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR1; TAB(53); "Location# "; Location$
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR2
  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; Zip$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); LongNDate$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "Dear Customer:"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     This letter is to notify you that your water and/or sewer"
  PRINT #UBRpt, TAB(10); "service will be disconnected. In order to avoid this service"
  PRINT #UBRpt, TAB(10); "interruption, a payment in the amount of:"; FUsing$(STR$(TotalBal#), "$$,######.##")
  PRINT #UBRpt, TAB(10); "must be made to the Norwood Town Office by 5:00 p.m. on or"
  PRINT #UBRpt, TAB(10); "before "; LongPDate$; "."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     If service is disconnected, your total balance must be"
  PRINT #UBRpt, TAB(10); "paid in full, plus a $20.00 reconnect fee in order for service"
  PRINT #UBRpt, TAB(10); "to be restored."
  PRINT #UBRpt, TAB(10); " "
  PRINT #UBRpt, TAB(10); "     If you have any questions, please contact the Town Hall"
  PRINT #UBRpt, TAB(10); "at 704-474-3416 Mon-Fri 9:00 a.m. to 5:00 p.m."
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); "     If payment has been made prior to receiving this notice,"
  PRINT #UBRpt, TAB(10); "please disregard this notice."
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); "Thank you,"
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); "Town of Norwood Utilities Department."
  PRINT #UBRpt, CHR$(12)


