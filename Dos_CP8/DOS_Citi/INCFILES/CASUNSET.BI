
  PRINT #UBRpt, TAB(29); "TOWN OF SUNSET BEACH"
  PRINT #UBRpt, TAB(26); "FINANCE DEPT - COLLECTION DIV"
  PRINT #UBRpt, TAB(30); "700 SUNSET BLVD. N."
  PRINT #UBRpt, TAB(28); "SUNSET BEACH, NC 28468"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(60); "DATE: "; DATE$
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(3); LEFT$(RTRIM$(UBCustRec(1).CustName), 30);
  PRINT #UBRpt, TAB(40); "    ACCT # "; AcctNo&
  PRINT #UBRpt, TAB(3); LEFT$(RTRIM$(UBCustRec(1).Addr1), 30)
  PRINT #UBRpt, TAB(3); QPTrim$(UBCustRec(1).Addr2)
  PRINT #UBRpt, TAB(3); QPTrim$(UBCustRec(1).City); " "; QPTrim$(UBCustRec(1).State); " "; QPTrim$(UBCustRec(1).ZipCode);
  PRINT #UBRpt, TAB(40); "LEGAL ADDR "; QPTrim$(UBCustRec(1).ServAddr)
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(3); "THIS IS AN INVOICE FOR ASSESSMENTS DUE TO THE TOWN OF SUNSET BEACH."
  PRINT #UBRpt, TAB(3); "THE STATUS OF THIS ACCOUNT IS SHOWN BELOW."
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(5); "Balance Due: "; USING "$$#####,#.##"; CustBal#
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(5); "Assessment payments may be paid in full or in "; QPTrim$(Message$); " equal annual"
  PRINT #UBRpt, TAB(5); "installments, based on the confirmation of the assessment roll."
  PRINT #UBRpt, TAB(5); "Interest is accrued monthly (using the simple interest method)"
  PRINT #UBRpt, TAB(5); "on any outstanding principle balance."
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(3); NoticeInfo.MsgLine1
  PRINT #UBRpt, TAB(3); NoticeInfo.MsgLine2
  PRINT #UBRpt, TAB(3); NoticeInfo.MsgLine3
  PRINT #UBRpt, TAB(3); NoticeInfo.MsgLine4
  PRINT #UBRpt, CHR$(12)








