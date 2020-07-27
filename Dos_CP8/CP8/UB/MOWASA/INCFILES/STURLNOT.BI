'
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "                                  Town of Stuart"
  PRINT #UBRpt, "                                   P.O. Box 422"
  PRINT #UBRpt, "                                 Stuart, VA 24171"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); LongNDate$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); UBCustRec(1).CustName
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR1
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR2
  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "Dear Customer:"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "   Our records of your account show that you are past due in your"
  PRINT #UBRpt, TAB(10); "payment to the Town of Stuart in the amount of"; USING "$$#####,#.##"; CustBal#;
  PRINT #UBRpt, ". This"
  PRINT #UBRpt, TAB(10); "letter is to serve notice that unless payment is received to this"
  PRINT #UBRpt, TAB(10); "office by "; LongPDate$; " service will be discontinued."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "   If you have any question you may contact us at the Town Office."
  PRINT #UBRpt, TAB(10); "Office hours are 9:00 a.m. to 1:00 p.m. and 2:00 p.m. to 5:00 p.m."
  PRINT #UBRpt, TAB(10); "Monday through Friday, or call us by phone at (540) 694-3811."
  PRINT #UBRpt, TAB(10); "If payment has already been made prior to receiving this notice,"
  PRINT #UBRpt, TAB(10); "Thank you and please disregard this notice."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "Susan C. Slate"
  PRINT #UBRpt, TAB(10); "Town Clerk/Treasurer"
  PRINT #UBRpt, CHR$(12)

