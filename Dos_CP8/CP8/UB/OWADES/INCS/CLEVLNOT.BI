'Cleveland
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "                               TOWN OF CLEVELAND"
  PRINT #UBRpt, "                                 NORTH CAROLINA"
  PRINT #UBRpt, "                             100 North Depot Street"
  PRINT #UBRpt, "                                  P.O. Box 429"
  PRINT #UBRpt, "                               Cleveland NC 27013"
  PRINT #UBRpt, "                                 (704) 278-4777"
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
  PRINT #UBRpt, "                                REMINDER NOTICE"
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    Your water/sewer payment was not received at Cleveland Town Hall by"
  PRINT #UBRpt, "    the due date on the invoice and a $2.00 late charge has been added"
  PRINT #UBRpt, "    to your account."
  PRINT #UBRpt,
  PRINT #UBRpt, "    If payment is not received by "; LongPDate$; " by 5:00 PM,"
  PRINT #UBRpt, "    service will be disconnected.  If service is disconnected, the current and"
  PRINT #UBRpt, "    past due amount must be paid along with a $25.00 reconnection fee."; ""
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "         ACCOUNT #"; Acct$; TAB(40); "AMOUNT DUE"; USING "$$#####,#.##"; CustBal#
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "    Thank you,"
  PRINT #UBRpt, "    Town of Cleveland"
  PRINT #UBRpt, "    Water and Sewer Department"
  PRINT #UBRpt, CHR$(12)

