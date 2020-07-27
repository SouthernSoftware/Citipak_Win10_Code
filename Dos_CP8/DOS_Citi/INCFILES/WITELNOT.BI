
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "                               Town of White Lake"
  PRINT #UBRpt,
  PRINT #UBRpt, "                          1879 White Lake Dr. PMB 7250"
  PRINT #UBRpt, "                                                "
  PRINT #UBRpt, "                              White Lake, NC 28337"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(5); LongNDate$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "    "; UBCustRec(1).CustName
  PRINT #UBRpt, "    "; UBCustRec(1).ADDR1
  PRINT #UBRpt, "    "; UBCustRec(1).ADDR2
  PRINT #UBRpt, "    "; QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode; TAB(55); "Loc: "; Location$
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "                                DISCONNECT NOTICE"
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    Dear Customer:"
  PRINT #UBRpt, ""
  PRINT #UBRpt, "       As of the above date, our records indicate that your account"
  PRINT #UBRpt, "    is PAST due.  Unless payment is made to this office within fourteen"
  PRINT #UBRpt, "    days of this notice, your service will be discontinued. (Pursuant"
  PRINT #UBRpt, "    to Town Code Sec. 17-19), If service is discontinued, a restoration"
  PRINT #UBRpt, "    charge in the amount of $25.00 is required in addition to your"
  PRINT #UBRpt, "    account balance, BEFORE service will be restored. You may appear"
  PRINT #UBRpt, "    at the White Lake town office on any business day, and request an"
  PRINT #UBRpt, "    informal hearing with the Town Clerk, for the purpose of showing"
  PRINT #UBRpt, "    error or working out a satisfactory extended payment arrangement."
  PRINT #UBRpt,
  PRINT #UBRpt, "       Please pay the amount listed below to prevent this scheduled"
  PRINT #UBRpt, "    service interruption."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "            ACCOUNT #"; Acct$; TAB(40); "TOTAL DUE:"; USING "$$#####,#.##"; CustBal#
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "       If payment has been made prior to receiving this notice, Please"
  PRINT #UBRpt, "    accept our thanks and disregard this notice."
  PRINT #UBRpt,
  PRINT #UBRpt, "       If you have any questions please contact us at 1879 White Lake Dr."
  PRINT #UBRpt, "    or, call 910-862-4800."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "    Town of White Lake"
  PRINT #UBRpt, "    Water & Sewer Dept."
  PRINT #UBRpt, CHR$(12)

