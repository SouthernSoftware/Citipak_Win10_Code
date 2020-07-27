'Brookneal VA
  Buck$ = QPTrim$(NoticeInfo.MsgLine1)
  PRINT #UBRpt,
  PRINT #UBRpt, "                                 Town of Brookneal"
  PRINT #UBRpt, "                                    P.O. Box 450"
  PRINT #UBRpt, "                                 Brookneal VA 24528"
  PRINT #UBRpt, "                                    434-376-3124"
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
  PRINT #UBRpt, TAB(10); "than 5:00 P.M. on "; LongPDate$; ", or you will be"
  PRINT #UBRpt, TAB(10); "charged a "; Buck$; " reconnect fee in addition to your past due amount."
  PRINT #UBRpt, TAB(10); "Your water and sewer service will be DISCONNECTED without further"
  PRINT #UBRpt, TAB(10); "notice. IF YOU DECIDE TO TURN ON YOUR METER YOURSELF, THE METER"
  PRINT #UBRpt, TAB(10); "WILL BE REMOVED AND AN ADDITIONAL "; Buck$; " FEE WILL ALSO BE ASSESSED!"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     If you have any questions you can contact the Town Office"
  PRINT #UBRpt, TAB(10); "during normal working hours."
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); "     If payment has been made prior to receiving this notice,"
  PRINT #UBRpt, TAB(10); "please disregard this notice."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "Thank You"
  PRINT #UBRpt, TAB(10); "Town of Brookneal Utility Department."
  PRINT #UBRpt, CHR$(12)

