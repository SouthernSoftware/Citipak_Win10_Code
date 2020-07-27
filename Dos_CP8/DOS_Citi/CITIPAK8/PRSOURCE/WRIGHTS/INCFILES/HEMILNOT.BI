'Cleveland
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "                               TOWN OF HEMINGWAY"
  PRINT #UBRpt, "                                  P.O. Box 968"
  PRINT #UBRpt, "                               HEMINGWAY SC 29554"
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(5); LongNDate$
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    "; UBCustRec(1).CustName; TAB(60); "ACCOUNT #"; Acct$
  PRINT #UBRpt, "    "; UBCustRec(1).ADDR1
  PRINT #UBRpt, "    "; UBCustRec(1).ADDR2
  PRINT #UBRpt, "    "; QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "                                REMINDER NOTICE"
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    Our records indicate that your account is past due. Your past due"
  PRINT #UBRpt, USING "    amount is $$#####,#.## If your bill is not paid by 5:00 p.m. on "; CustBal#
  PRINT #UBRpt, "    "; LongPDate$; ", your water service will be discontinued, and"
  PRINT #UBRpt, "    a fee of $10.00 will be added to your account regardless of whether"
  PRINT #UBRpt, "    service is discontinued or not."
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    The past due bill and $10.00 fee must be paid in full before service"
  PRINT #UBRpt, "    will be reconnected. No exceptions will be made."
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    A drop box is available for your convenience."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "    Thank you,"
  PRINT #UBRpt, "    Hemingway Town Hall"
  PRINT #UBRpt, CHR$(12)

