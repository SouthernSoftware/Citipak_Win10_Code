'Polkton Late notice
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "                                  Town of Polkton"
  PRINT #UBRpt, "                                    P.O. Box 99"
  PRINT #UBRpt, "                                   113 W Polk St."
  PRINT #UBRpt, "                                 Polkton, NC 28135"
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
  PRINT #UBRpt, TAB(10); "Dear Water/Sewer Customer:"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     As of the above date, Our records indicate that your"
  PRINT #UBRpt, TAB(10); "account is PAST due.  Unless payment is made to this office"
  PRINT #UBRpt, TAB(10); "within SEVEN days of this notice, Your service may be "
  PRINT #UBRpt, TAB(10); "discontinued.  If service is discontinued, a RECONNECT FEE"
  PRINT #UBRpt, TAB(10); "is required in addition to the BILLED AMOUNT DUE before"
  PRINT #UBRpt, TAB(10); "service will be restored."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     Please pay the amount of: "; FUsing$(STR$(TotalBal#), "$$,######.##"); " to prevent this"
  PRINT #UBRpt, TAB(10); "scheduled service interruption."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     If you have any questions you can contact us at our"
  PRINT #UBRpt, TAB(10); "office 113 W Polk St. or, call 272-7463"
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
  PRINT #UBRpt, TAB(10); "Town of Polkton Water Department."
  PRINT #UBRpt, CHR$(12)


