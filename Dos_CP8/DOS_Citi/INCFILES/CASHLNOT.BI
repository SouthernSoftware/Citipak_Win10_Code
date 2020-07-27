'Polkton Late notice
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "                                  Town of Cashion"
  PRINT #UBRpt, "                                    P.O. Box 27"
  PRINT #UBRpt, "                                 Cashion OK 73016"
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
  PRINT #UBRpt, TAB(10); "account is PAST due. Unless payment is made to this office by"
  PRINT #UBRpt, TAB(10); "the last day of the month, Your service will be discontinued."
  PRINT #UBRpt, TAB(10); "If service is discontinued, a reconnect fee is required in"
  PRINT #UBRpt, TAB(10); "addition to the BILLED AMOUNT DUE before service will be"
  PRINT #UBRpt, TAB(10); "restored."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     Please pay the amount of: "; FUsing$(STR$(TotalBal#), "$$,######.##"); " to prevent this"
  PRINT #UBRpt, TAB(10); "scheduled service interruption."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     If you have any questions you can contact us at our"
  PRINT #UBRpt, TAB(10); "office Mon-Fri 8:30 am to 5:00 pm"
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
  PRINT #UBRpt, TAB(10); "Town of Cashion Utilities Department."
  PRINT #UBRpt, CHR$(12)


