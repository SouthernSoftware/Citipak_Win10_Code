'Ryan, OK.  Late notice
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "                                    Town of Ryan"
  PRINT #UBRpt, "                                    P.O. Box 489"
  PRINT #UBRpt, "                                   Ryan, OK 73565"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).CustName); TAB(53); " Account# "; Acct$
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR1; TAB(53); "Location# "; Location$
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR2
  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; Zip$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); LongNDate$
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "Dear Utilities Customer:"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     As of the above date, Our records indicate that your"
  PRINT #UBRpt, TAB(10); "account is PAST due.  Unless payment is made to this office"
  PRINT #UBRpt, TAB(10); "within TEN days of this notice, Your service may be "
  PRINT #UBRpt, TAB(10); "discontinued.  If service is discontinued, a RECONNECT FEE"
  PRINT #UBRpt, TAB(10); "is required in addition to the BILLED AMOUNT DUE before"
  PRINT #UBRpt, TAB(10); "service will be restored."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     Please pay the amount of: "; FUsing$(STR$(TotalBal#), "$$,######.##"); " to prevent this"
  PRINT #UBRpt, TAB(10); "scheduled service interruption."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     If you have any questions you can contact us at our"
  PRINT #UBRpt, TAB(10); "office Mon-Fri 8:00 am to 4:30 pm or, call 405-757-2277"
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
  PRINT #UBRpt, TAB(10); "Town of Ryan Utilities Department."
  PRINT #UBRpt, CHR$(12)


