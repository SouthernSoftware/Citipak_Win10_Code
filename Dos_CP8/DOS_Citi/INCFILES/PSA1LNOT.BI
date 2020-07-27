PRINT #UBRpt, TAB(20); "GILES COUNTY PUBLIC SERVICE AUTHORITY"
PRINT #UBRpt, TAB(20); "            P.O. BOX 510"
PRINT #UBRpt, TAB(20); "       PEMBROKE, VIRGINIA 24136"
PRINT #UBRpt, TAB(20); "   PHONE: (540) 921-2499; 626-7046"
PRINT #UBRpt, TAB(20); ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, "           "; UBCustRec(1).CustName
PRINT #UBRpt, "           "; UBCustRec(1).Addr1
PRINT #UBRpt, "           "; UBCustRec(1).Addr2
PRINT #UBRpt, "           "; QPtrim$(UBCustRec(1).City); ", "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, "           "; LongNDate$
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, "              PLEASE BE ADVISED THIS NOTICE WILL SERVE AS A"
IF CustBal# > 99# THEN
  PRINT #UBRpt, "              FINAL";
ELSE
  PRINT #UBRpt, "             ";
END IF

PRINT #UBRpt, " PAST DUE NOTICE FOR YOUR GARBAGE ACCOUNT."
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, "                         Account: "; Acct$; "  Stat: "; UBCustRec(1).Status
PRINT #UBRpt, ""
PRINT #UBRpt, "                         Balance:"; USING "$$####,#.##"; CustBal#
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, "           Please remit payment on this past due account within"
PRINT #UBRpt, "           15 days from the date on this notice or your account"
IF CustBal# > 99# THEN
  PRINT #UBRpt, "           will ";
ELSE
  PRINT #UBRpt, "           may ";
END IF

PRINT #UBRpt, "be turned over to our attorney for collection. To"
PRINT #UBRpt, "           avoid further action, please remit payment promptly."
PRINT #UBRpt, ""
PRINT #UBRpt, "           If you have questions concerning your bill, please"
PRINT #UBRpt, "           call our office at 626-7046 or 921-2499. If your "
PRINT #UBRpt, "           payment has been sent or this account has been paid"
PRINT #UBRpt, "           in full, please disregard this notice."
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, "           Thank you,"
PRINT #UBRpt, "           Giles County PSA Utility Billing Department."
PRINT #UBRpt, ""

