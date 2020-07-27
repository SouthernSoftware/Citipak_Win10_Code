  PRINT #UBRpt,
  PRINT #UBRpt, "                                  BURNETTOWN"
  PRINT #UBRpt, "                                 WATER SYSTEM"
  PRINT #UBRpt, "                               BATH, S.C. 29816"
  PRINT #UBRpt, "                                (803) 593-2676"
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
  PRINT #UBRpt, TAB(9); "If service is cut off, A reconnect fee of $25.00 during business"
  PRINT #UBRpt, TAB(9); "hours or, $40.00 after business hours will be charged, before your"
  PRINT #UBRpt, TAB(9); "water is reconnected."
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
  PRINT #UBRpt, TAB(9); " BURNETTOWN"
  PRINT #UBRpt, TAB(9); " Water System"
  PRINT #UBRpt, CHR$(12)

