'
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "                              Spencer Water Department"
  PRINT #UBRpt, "                                    P.O. Box 187"
  PRINT #UBRpt, "                                 Spencer TN 38585"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).CustName); TAB(51); " Account: "; Acct$
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR1; TAB(51); "Location: "; Location$
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR2
  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; Zip$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); LongNDate$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "TURN-OFF NOTICE"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "Dear Utilities Customer:"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     As of the above date, Our records indicate that your"
  PRINT #UBRpt, TAB(10); "account is PAST due.  Unless payment is made to this office"
  PRINT #UBRpt, TAB(10); "by 4:00 PM on the 20th of the month, Your service will be"
  PRINT #UBRpt, TAB(10); "discontinued.  If service is discontinued, a RECONNECT FEE"
  PRINT #UBRpt, TAB(10); "is required in addition to the BILLED AMOUNT DUE before"
  PRINT #UBRpt, TAB(10); "service will be restored."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     Please pay the amount of: "; FUsing$(STR$(TotalBal#), "$$,######.##"); " to prevent this"
  PRINT #UBRpt, TAB(10); "scheduled service interruption."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     If you have any questions please contact us at Spencer"
  PRINT #UBRpt, TAB(10); "Water Dept., Spring Street or, call 946-2351"
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
  PRINT #UBRpt, TAB(10); "     Spencer Water Department."
  PRINT #UBRpt, CHR$(12)


