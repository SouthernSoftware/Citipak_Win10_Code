'Brookneal VA
  PRINT #UBRpt,
  PRINT #UBRpt, "                                 Town of Brookneal"
  PRINT #UBRpt, "                                    P.O. Box 450"
  PRINT #UBRpt, "                                 Brookneal VA 24528"
  PRINT #UBRpt, "                                    803-376-3124"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).CustName)
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR1
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR2
  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; Zip$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); LongNDate$
  PRINT #UBRpt,
  PRINT #UBRpt, "~"; TAB(78); "~"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); " Account# "; Acct$
  PRINT #UBRpt, TAB(10); "Location# "; Location$;
  PRINT #UBRpt, TAB(50); "Amount: "; FUsing$(STR$(TotalBal#), "$$,#####.##");
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "Dear Customer:"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     Please be advised our records indicate that your water and"
  PRINT #UBRpt, TAB(10); "sewer bill is past due. Your bill should be PAID IN FULL no later"
  PRINT #UBRpt, TAB(10); "than 5:00 P.M. on "; LongPDate$; ". If no payment is"
  PRINT #UBRpt, TAB(10); "made by the CUT-OFF DATE, then your Water & Sewer service will be"
  PRINT #UBRpt, TAB(10); "DISCONNECTED WITHOUT FURTHER NOTICE!  If service is disconnected,"
  PRINT #UBRpt, TAB(10); "there is a $30.00 reconnect fee, in addition to PAST DUE AMOUNT,"
  PRINT #UBRpt, TAB(10); "to have your Water & Sewer service restored!  IF YOU DECIDE TO "
  PRINT #UBRpt, TAB(10); "TURN YOUR METER ON YOURSELF, THE METER WILL BE REMOVED AND AN "
  PRINT #UBRpt, TAB(10); "ADDITIONAL $30.00 FEE WILL ALSO BE ASSESSED."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     If you have any questions you can contact the Town Office"
  PRINT #UBRpt, TAB(10); "during normal working hours."
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); "     If payment has been made prior to receiving this notice,"
  PRINT #UBRpt, TAB(10); "Please disregard this notice."
  PRINT #UBRpt, "~"; TAB(78); "~"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "Town of Brookneal Utility Department."
  PRINT #UBRpt, CHR$(12)

