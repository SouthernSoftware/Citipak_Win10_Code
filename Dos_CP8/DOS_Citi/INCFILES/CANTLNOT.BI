
  PRINT #UBRpt,
  PRINT #UBRpt, "                                TOWN OF CANTON"
  PRINT #UBRpt, "                               WATER DEPARTMENT"
  PRINT #UBRpt, "                                 CANTON, N.C."
  PRINT #UBRpt, "                                   648-2363"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(9); LongNDate$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(9); QPTrim$(UBCustRec(1).CustName)
  PRINT #UBRpt, TAB(9); UBCustRec(1).ADDR1
  IF LEN(QPTrim$(UBCustRec(1).ADDR2)) > 0 THEN
    PRINT #UBRpt, TAB(9); UBCustRec(1).ADDR2
  END IF
  PRINT #UBRpt, TAB(9); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(9); "  Your account shows a past due balance. The past due amount"
  PRINT #UBRpt, TAB(9); "MUST be PAID by closing on "; LongPDate$; ". If the"
  PRINT #UBRpt, TAB(9); "PAST DUE amount is not paid, your service will be CUT OFF."
  PRINT #UBRpt, TAB(9); "If service is cut off, A $25.00 fee will be charged, before"
  PRINT #UBRpt, TAB(9); "your water is reconnected."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(9); "If you feel this is incorrect, please contact the Town Hall"
  PRINT #UBRpt, TAB(9); "to schedule a hearing before the due date in order to avoid"
  PRINT #UBRpt, TAB(9); "any unnecessary interruption of your service."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(30); USING " Account:##########"; AcctNo&
  PRINT #UBRpt, TAB(30); "Location: "; Location$
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(29); USING "TOTAL DUE:$$#####.##"; CustBal#
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(9); " Thank you,"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(9); " Town of Canton"
  PRINT #UBRpt, TAB(9); " Water Department"
  PRINT #UBRpt, CHR$(12)

