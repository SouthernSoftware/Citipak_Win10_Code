  Line1$ = QPTrim$(NoticeInfo.MsgLine1)
  Line2$ = QPTrim$(NoticeInfo.MsgLine2)
  Line3$ = QPTrim$(NoticeInfo.MsgLine3)

  PRINT #UBRpt, TAB(44); "Remington Water Department"
  PRINT #UBRpt, TAB(44); "203 East Main Street"
  PRINT #UBRpt, TAB(44); "Remington, VA  22734"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(44); LongNDate$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(5); UBCustRec(1).CustName
  PRINT #UBRpt, TAB(5); UBCustRec(1).Addr1
  PRINT #UBRpt, TAB(5); LEFT$(UBCustRec(1).City, 15); " "; UBCustRec(1).State; "  "; UBCustRec(1).ZipCode
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(44); "RE: Account "; Acct$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(5); "Attention: "; UBCustRec(1).CustName
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(5); "Our records show that payment of your bill is past due in the amount"
  PRINT #UBRpt, TAB(5); "of "; USING "$$####.##"; CustBal#;
  PRINT #UBRpt, " through the "; Line1$; " of this month."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(5); "Payment in full must be received by 4:30 p.m. on "; PDate$;
  PRINT #UBRpt, " or service"
  PRINT #UBRpt, TAB(5); "will be discontinued on "; Line3$;
  PRINT #UBRpt, ".  If service is discontinued, there"
  PRINT #UBRpt, TAB(5); "will be a reconnection charge of "; Line2$;
  PRINT #UBRpt, " to restore your service."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(5); "If for any reason you feel that our records may be in error, please"
  PRINT #UBRpt, TAB(5); "contact us immediately.  Please help us to resolve this matter soon"
  PRINT #UBRpt, TAB(5); "so that it will not be necessary to discontinue your service."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(44); "Sincerely,"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(44); "Remington Water Department"
  PRINT #UBRpt, TAB(44); "(540) 439-3220"
  PRINT #UBRpt, CHR$(12);
RETURN

