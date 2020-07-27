'Crewe VA.  Late notice
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "                                   TOWN OF CREWE"
  PRINT #UBRpt, "                                125 E. CAROLINA AVE."
  PRINT #UBRpt, "                                  CREWE, VA  23930"
  PRINT #UBRpt, "                                   (804) 645-9453"
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
  PRINT #UBRpt, TAB(10); "FROM: WILFRED D. WILSON,  MAYOR"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "RE:   NOTICE OF WATER DISCONNECTION"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     This letter is to notify you that your water service will"
  PRINT #UBRpt, TAB(10); "be disconnected because of non-payment. In order to avoid this"
  PRINT #UBRpt, TAB(10); "service interruption, you must pay the amount of:"; FUsing$(STR$(TotalBal#), "$$,######.##")
  PRINT #UBRpt, TAB(10); "to the Town Office by 5:00 p.m. on, or before "; PDate$; "."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     If your bill is paid by a check and is returned from your"
  PRINT #UBRpt, TAB(10); "bank for any reason, your water service will be disconnected"
  PRINT #UBRpt, TAB(10); "WITHOUT further notice!"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     If service is disconnected, you will forfeit your original"
  PRINT #UBRpt, TAB(10); "deposit, and will be required to pay another deposit of $100.00"
  PRINT #UBRpt, TAB(10); "and, a $20.00 reconnection fee."
  PRINT #UBRpt, TAB(10); " "
  PRINT #UBRpt, TAB(10); "     If you have any questions you can contact us at the Town"
  PRINT #UBRpt, TAB(10); "Office Mon-Fri 8:30 a.m. to 5:00 p.m. or, call 804-645-9453"
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); "     If payment has been made prior to receiving this notice,"
  PRINT #UBRpt, TAB(10); "Please disregard this notice."
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); "Thank you,"
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); "Town of Crewe Utilities Department."
  PRINT #UBRpt, CHR$(12)


