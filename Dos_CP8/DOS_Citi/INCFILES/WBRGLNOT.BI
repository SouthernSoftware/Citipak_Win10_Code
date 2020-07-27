  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(20); "WILLIAMSBURG COUNTY WATER/SEWER AUTHORITY"
  PRINT #UBRpt, TAB(35); "P.O. BOX 1124"
  PRINT #UBRpt, TAB(32); "KINGSTREE, SC 29556"
  PRINT #UBRpt, TAB(34); "(843) 355-9321"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(5); LongNDate$
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    "; UBCustRec(1).CustName
  PRINT #UBRpt, "    "; UBCustRec(1).ADDR1
  PRINT #UBRpt, "    "; UBCustRec(1).ADDR2
  PRINT #UBRpt, "    "; QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "                                  FINAL NOTICE"
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "       According to our records, your account is now delinquent in the"
  PRINT #UBRpt, "    amount shown below. As stated on your monthly bill, your service is"
  PRINT #UBRpt, "    subject to CUT-OFF if not paid by 10:00 AM, on "; LongPDate$
  PRINT #UBRpt,
  PRINT #UBRpt, "    IF WE DO NOT HEAR FROM YOU, YOUR SERVICE WILL BE DISCONNECTED WITHOUT,"
  PRINT #UBRpt, "    FURTHER NOTICE, NO EXCEPTIONS. For your service to be restored, the"
  PRINT #UBRpt, "    past due amount must be paid along with a $25.00 reconnection fee."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "    If payment has been made prior to receiving this notice,"
  PRINT #UBRpt, "    Please disregard this notice."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "         ACCOUNT #"; Acct$; TAB(40); "AMOUNT DUE"; USING "$$#####,#.##"; CustBal#
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "    Thank you,"
  PRINT #UBRpt, "    Williamsburg County Water/Sewer Authority"
  PRINT #UBRpt, CHR$(12)

