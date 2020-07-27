PRINT #UBRpt, ""
PRINT #UBRpt, "             Important Notice"
PRINT #UBRpt,
PRINT #UBRpt, "Our records indicate you have a past due account."
PRINT #UBRpt, "Your water service is scheduled to be terminated"
PRINT #UBRpt, "in Five (5) working days unless past due amount "
PRINT #UBRpt, "of"; FUsing$(STR$(TotalBal#), "$$,######.##"); " is paid at the Town Office"
PRINT #UBRpt, "by 5:00 pm, "; LongPDate$
PRINT #UBRpt, "If service is terminated a $15.00 reconnect fee"
PRINT #UBRpt, "will be charged."
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, TAB(26); UBCustRec(1).CustName
PRINT #UBRpt,
PRINT #UBRpt, TAB(26); UBCustRec(1).ADDR1
PRINT #UBRpt, TAB(26); UBCustRec(1).ADDR2
PRINT #UBRpt, TAB(26); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; Zip$
PRINT #UBRpt,
PRINT #UBRpt,
PRINT #UBRpt, ""
Next2Print = Next2Print + 1





