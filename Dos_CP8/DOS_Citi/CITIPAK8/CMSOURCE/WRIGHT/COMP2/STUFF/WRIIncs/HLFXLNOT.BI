  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "                                 TOWN OF HALIFAX"
  PRINT #UBRpt,
  PRINT #UBRpt, "                                   P.O. BOX 627"
  PRINT #UBRpt,
  PRINT #UBRpt, "                                HALIFAX VA, 24558"
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
  PRINT #UBRpt, TAB(10); "TURN-OFF NOTICE"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "Dear Water/Sewer Customer:"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     As of the above date, our records indicate that your account"
  PRINT #UBRpt, TAB(10); "is PAST due.  Unless payment is made to this office within five"
  PRINT #UBRpt, TAB(10); "days of this notice, your service may be discontinued. If service"
  PRINT #UBRpt, TAB(10); "is discontinued, a restoration charge is required in addition to"
  PRINT #UBRpt, TAB(10); "the TOTAL bill due."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     Please pay the following total amount of:"; FUsing$(STR$(TotalBal# + (UBCustRec(1).CurrBalance * .1)), "$$,######.##")
  PRINT #UBRpt, TAB(10); "to prevent this scheduled turn-off."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "    If you have any questions please contact us at Town Hall,"
  PRINT #UBRpt, TAB(10); "70 Main Street, or call 476-2343."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     If payment has been made prior to receiving this notice,"
  PRINT #UBRpt, TAB(10); "please disregard this notice."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "Thank you,"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "TOWN OF HALIFAX WATER DEPARTMENT"
  PRINT #UBRpt, CHR$(12)


