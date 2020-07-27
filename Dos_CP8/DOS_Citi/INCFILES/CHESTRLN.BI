  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(33); "Town of Chester"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(33); "  P.O. Box 67"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(32); "Chester, GA  31012"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); UBCustRec(1).CustName; TAB(55); DATE$
  PRINT #UBRpt, TAB(10); UBCustRec(1).Addr1
  PRINT #UBRpt, TAB(10); UBCustRec(1).Addr2
  PRINT #UBRpt, TAB(10); UBCustRec(1).City; " "; UBCustRec(1).State; "  "; UBCustRec(1).ZipCode; TAB(53); "Loc. No: "; UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "PAST DUE NOTICE"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "Dear Water/Sewer Customer:"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(15); "As of the above date, our records indicate that your"
  PRINT #UBRpt, TAB(10); "account is PAST DUE.  Unless payment is made to this"
  PRINT #UBRpt, TAB(10); "office within ten days of this notice, your service could"
  PRINT #UBRpt, TAB(10); "be disconnected.  If service is discontinued, a re-connection"
  PRINT #UBRpt, TAB(10); "fee of $25.00 is required in addition to the bill outstanding"
  PRINT #UBRpt, TAB(10); "BEFORE service will be restored."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(15); "Please pay the following PAST DUE amount of "; USING "#####.##"; CustBal#
  PRINT #UBRpt, TAB(10); "to prevent this scheduled turn-off:"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(15); "Water/Sewer payments are DUE by the 15th of every month."
  PRINT #UBRpt, TAB(10); "Should payment NOT be received by then, LATE CHARGES of 10% are"
  PRINT #UBRpt, TAB(10); "assessed to any balance previous or current and added to your bill"
  PRINT #UBRpt, TAB(10); "monthly. Due to untimely payments, this could become costly. Which"
  PRINT #UBRpt, TAB(10); "requires you to pay unnecessary fees."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(15); "If payment has been made prior to receiving this notice,"
  PRINT #UBRpt, TAB(10); "please accept our thanks and disregard this notice."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(15); "If you have any questions please contact us at City"
  PRINT #UBRpt, TAB(10); "Hall, (358-4550) between the hours of 9:00 am through"
  PRINT #UBRpt, TAB(10); "2:00 pm."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "TOWN OF CHESTER"
  PRINT #UBRpt, TAB(10); "WATER/SEWER DEPARTMENT"
  PRINT #UBRpt, CHR$(12);

