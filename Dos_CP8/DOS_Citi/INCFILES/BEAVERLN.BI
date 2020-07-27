PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(35); "Town of Beaver"
PRINT #UBRpt,
PRINT #UBRpt, TAB(36); "P.O. Box 698"
PRINT #UBRpt,
PRINT #UBRpt, TAB(32); "Beaver, OK  73932-0698"
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).CustName); TAB(55); DATE$
PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR1
PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR2
PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; Zip$; TAB(53); " Account# "; Acct$
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); "TURN-OFF NOTICE"
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); "Dear Utility Customer:"
PRINT #UBRpt,
PRINT #UBRpt, TAB(15); "As of the above date, our records indicate that your"
PRINT #UBRpt, TAB(10); "account is PAST due.  Unless payment is made to this"
PRINT #UBRpt, TAB(10); "office by "; LongPDate$; ", your service may be"
PRINT #UBRpt, TAB(10); "discontinued.  If service is discontinued, a reconnect"
PRINT #UBRpt, TAB(10); "fee of $25.00 will be added to the bill before service"
PRINT #UBRpt, TAB(10); "will be restored."
PRINT #UBRpt,
PRINT #UBRpt, TAB(15); "Please pay the following total amount of: "; QPTrim$(FUsing$(STR$(TotalBal#), "$$,######.##"))
PRINT #UBRpt, TAB(10); "to prevent this scheduled turn-off."
PRINT #UBRpt,
PRINT #UBRpt, TAB(15); "If payment has been made prior to receiving this"
PRINT #UBRpt, TAB(10); "notice, please accept our thanks and disregard this"
PRINT #UBRpt, TAB(10); "notice."
PRINT #UBRpt,
PRINT #UBRpt, TAB(15); "To prevent disconnection, you must contact us at the"
PRINT #UBRpt, TAB(10); "Town Hall, 30 West 2nd St., or call 625-3331."
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(10); "TOWN OF BEAVER"
PRINT #UBRpt, TAB(10); "OFFICE OF THE CLERK"
PRINT #UBRpt,
PRINT #UBRpt, CHR$(12)

