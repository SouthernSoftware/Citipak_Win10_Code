'Saluda
  PRINT #UBRpt, "                                Town of Andrews"
  PRINT #UBRpt,
  PRINT #UBRpt, "                                 P.O. Box 1210"
  PRINT #UBRpt, "                                                "
  PRINT #UBRpt, "                               Andrews, NC 28901"
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
  PRINT #UBRpt, "    is PAST due.  Unless payment is made to this office within five"
  PRINT #UBRpt, "    days of this notice, your service may be discontinued.  If service"
  PRINT #UBRpt, "    is discontinued, a restoration charge is required in addition to"
  PRINT #UBRpt, "    your account balance, BEFORE service will be restored."
  PRINT #UBRpt,
  PRINT #UBRpt, "       Please pay the amount listed below to prevent this scheduled"
  PRINT #UBRpt, "    service interruption."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(26); "   WATER/SEWER:"; USING "$$#####,#.##"; CustBal#
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(26); "     TOTAL DUE:"; USING "$$#####,#.##"; CustBal#
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "       If payment has been made prior to receiving this notice, Please"
  PRINT #UBRpt, "    accept our thanks and disregard this notice."
  PRINT #UBRpt,
  PRINT #UBRpt, "       If you have any questions please contact us at the Town Hall,"
  PRINT #UBRpt, "    Main Street, or call 828-321-5111."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "     Town of Andrews"
  PRINT #UBRpt, "    Water Department"
  PRINT #UBRpt, CHR$(12)

