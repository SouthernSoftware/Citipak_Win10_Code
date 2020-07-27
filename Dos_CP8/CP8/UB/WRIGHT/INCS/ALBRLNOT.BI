
  'PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "                                TOWN OF ALBERTA"
  PRINT #UBRpt, "                                  P.O. BOX 157"
  PRINT #UBRpt, "                               ALBERTA, VA  23821"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(9); LongNDate$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  'PRINT #UBRpt, TAB(9); QPTrim$(UBCustRec(1).CustName); TAB(57); USING " ACCOUNT:##########"; AcctNo&
  'PRINT #UBRpt, TAB(9); UBCustRec(1).ADDR1; TAB(57); "LOCATION: "; Location$
  'PRINT #UBRpt, TAB(9); UBCustRec(1).ADDR2; TAB(57); USING "PAST DUE:$$#####.##"; CustBal#
  'PRINT #UBRpt, TAB(9); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode

  PRINT #UBRpt, TAB(9); QPTrim$(UBCustRec(1).CustName)
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(9); UBCustRec(1).ADDR1
  IF LEN(QPTrim$(UBCustRec(1).ADDR2)) > 0 THEN
    PRINT #UBRpt, TAB(9); UBCustRec(1).ADDR2
  END IF
  PRINT #UBRpt, TAB(9); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode

  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(9); "  Your water account shows a past due balance. The past due"
  PRINT #UBRpt, TAB(9); "amount MUST be PAID by closing on "; LongPDate$; ". If the"
  PRINT #UBRpt, TAB(9); "PAST DUE amount is not paid, your water service will be CUT OFF."
  PRINT #UBRpt, TAB(9); "If service is cut off, A $25.00 fee will be charged, before"
  PRINT #UBRpt, TAB(9); "your water is reconnected."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(9); "Please pay now to avoid additional charges."
  PRINT #UBRpt,
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
  PRINT #UBRpt, TAB(9); " Town of Alberta"
  PRINT #UBRpt, TAB(9); " Water Department"
  PRINT #UBRpt, CHR$(12)

