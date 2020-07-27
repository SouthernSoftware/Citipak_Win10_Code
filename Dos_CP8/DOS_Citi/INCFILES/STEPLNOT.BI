  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "                              TOWN OF STEPHENS CITY"
  PRINT #UBRpt,
  PRINT #UBRpt, "                                   P.O. BOX 250"
  PRINT #UBRpt,
  PRINT #UBRpt, "                           STEPHENS CITY VA, 22655-0250"
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
  PRINT #UBRpt, TAB(10); "REMINDER NOTICE"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "Dear Water/Sewer Customer:"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     As of the above date, our records indicate that your account"
  PRINT #UBRpt, TAB(10); "is PAST due.  Payment must be made to this office by: "; PDate$
  PRINT #UBRpt, TAB(10); "or your service will be discontinued. If service is discontinued,"
  PRINT #UBRpt, TAB(10); "a restoration charge of $10.00 is required in addition to the"
  PRINT #UBRpt, TAB(10); "TOTAL bill due."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     Please pay the following total amount of:"; FUsing$(STR$(TotalBal#), "$$,######.##")
  PRINT #UBRpt, TAB(10); "to prevent this scheduled turn-off."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "    If you have any questions please contact us at Town Hall,"
  PRINT #UBRpt, TAB(10); "or call 540-869-3087."
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
  PRINT #UBRpt, TAB(10); "STEPHENS CITY TOWN OFFICE"
  PRINT #UBRpt, CHR$(12)


