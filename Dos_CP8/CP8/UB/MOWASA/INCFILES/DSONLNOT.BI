'Polkton Late notice
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "                                   Town of Dobson"
  PRINT #UBRpt, "                                    P.O. Box 351"
  PRINT #UBRpt, "                                  Dobson NC 27017"
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
  PRINT #UBRpt, TAB(10); "within SEVEN days of this notice, Your service may be "
  PRINT #UBRpt, TAB(10); "discontinued.  If service is discontinued, a $25.00 Reconnect"
  PRINT #UBRpt, TAB(10); "FEE is required in addition to the BILLED AMOUNT DUE before"
  PRINT #UBRpt, TAB(10); "service will be restored."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     Please pay the amount of: "; FUsing$(STR$(TotalBal#), "$$,######.##"); " to prevent this"
  PRINT #UBRpt, TAB(10); "scheduled service interruption."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "     If you have any questions you can contact us at our"
  PRINT #UBRpt, TAB(10); "office Mon-Fri 8:30 am to 5:00 pm or, call 336-386-8962"
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
  PRINT #UBRpt, TAB(10); "Town of Dobson Utilities Department."
  PRINT #UBRpt, CHR$(12)


