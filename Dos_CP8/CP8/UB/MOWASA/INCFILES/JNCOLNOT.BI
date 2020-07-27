'Johnston CO.
  PRINT #UBRpt,
  PRINT #UBRpt, "                                  Johnston County"
  PRINT #UBRpt, "                                  Public Utilities"
  PRINT #UBRpt, "                                    P.O. Box 2234"
  PRINT #UBRpt, "                                 Smithfield NC 27577"
  PRINT #UBRpt, "                                    919-989-5075"
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
  PRINT #UBRpt, "                            F I N A L  -  N O T I C E"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "When we placed this Final Notice in the mail, payment of your"
  PRINT #UBRpt, TAB(10); "bill had not been received for your water service."
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); "If the Past due amount as shown above, is not paid in full by"
  PRINT #UBRpt, TAB(10); LongPDate$; " we will take action to disconnect"
  PRINT #UBRpt, TAB(10); "your water service. If the PAST DUE amount is not paid in full"
  PRINT #UBRpt, TAB(10); "when our representative is dispatched to disconnect service, the"
  PRINT #UBRpt, TAB(10); "TOTAL DUE, as shown above, will be required. When the service is "
  PRINT #UBRpt, TAB(10); "disconnected, the TOTAL DUE, AND A $25.00 RECONNECTION FEE MUST"
  PRINT #UBRpt, TAB(10); "BE PAID TO HAVE YOUR WATER SERVICE RESTORED. If you have any"
  PRINT #UBRpt, TAB(10); "questions about your water service, contact Public Utilities,"
  PRINT #UBRpt, TAB(10); "Monday thru Friday, between 8:00 am. and 5:00 pm."
  PRINT #UBRpt, TAB(10); ""
  PRINT #UBRpt, TAB(10); "     If payment has been made prior to receiving this notice,"
  PRINT #UBRpt, TAB(10); "Please disregard this notice."
  PRINT #UBRpt, "~"; TAB(78); "~"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "HOPEWELL-PISGAH WATER DISTRICT"
  PRINT #UBRpt, CHR$(12)

