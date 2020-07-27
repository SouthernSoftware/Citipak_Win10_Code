'Johnston CO.
  NDate2$ = QPTrim$(NoticeInfo.MsgLine1)
  PRINT #UBRpt, "                                  Johnston County"
  PRINT #UBRpt, "                                  Public Utilities"
  PRINT #UBRpt, "                                  P.O. Box 2234"
  PRINT #UBRpt, "                                  204 S. Second St."
  PRINT #UBRpt, "                                  Smithfield NC 27577"
  PRINT #UBRpt, "                                  919-989-5075"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(6); QPTrim$(UBCustRec(1).CustName)
  PRINT #UBRpt, TAB(6); UBCustRec(1).ADDR1
  PRINT #UBRpt, TAB(6); UBCustRec(1).ADDR2
  PRINT #UBRpt, TAB(6); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; Zip$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(6); LongNDate$
  PRINT #UBRpt,
  PRINT #UBRpt, "~"; TAB(78); "~"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(6); " Account# "; Acct$
  PRINT #UBRpt, TAB(6); "Location# "; Location$;
  PRINT #UBRpt, TAB(50); "Amount: "; FUsing$(STR$(TotalBal#), "$$,#####.##");
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(22); "N O T I C E   T O   D I S C O N N E C T"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(6); "Dear Water Customer:"
  PRINT #UBRpt, TAB(6); ""
  PRINT #UBRpt, TAB(6); "  We have not received payment of your water bill this month."
  PRINT #UBRpt, TAB(6); "Your account is scheduled for disconnection on "; NDate2$; ". A service"
  PRINT #UBRpt, TAB(6); "fee of $25.00 will be added to accounts not paid BEFORE this date."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(6); "Payment must be made at this office on or before the close of business"
  PRINT #UBRpt, TAB(6); "on "; PDate$; " to prevent disconnection of your water service."
  PRINT #UBRpt, TAB(6); "Our office hours are Monday - Friday, 8:00 am. - 5:00 pm."
  PRINT #UBRpt, TAB(6); "A night drop box is available for your convenience."
  PRINT #UBRpt, TAB(6); ""
  PRINT #UBRpt, TAB(6); "If your service is disconnected, the total amount due, plus the $25.00"
  PRINT #UBRpt, TAB(6); "service fee must be paid before your service will be restored."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(6); "  If payment has been made prior to receiving this notice, please"
  PRINT #UBRpt, TAB(6); "disregard this notice."
  PRINT #UBRpt, "~"; TAB(78); "~"
  PRINT #UBRpt, TAB(6); ""
  PRINT #UBRpt, TAB(6); "Thank You."
  PRINT #UBRpt, CHR$(12)


'Johnston CO.
'  PRINT #UBRpt,
'  PRINT #UBRpt, "                                  Johnston County"
'  PRINT #UBRpt, "                                  Public Utilities"
'  PRINT #UBRpt, "                                  P.O. Box 2234"
'  PRINT #UBRpt, "                                  Smithfield NC 27577"
'  PRINT #UBRpt, "                                  919-989-5075"
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).CustName)
'  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR1
'  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR2
'  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; Zip$
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt, TAB(10); LongNDate$
'  PRINT #UBRpt,
'  PRINT #UBRpt, "~"; TAB(78); "~"
'  PRINT #UBRpt,
'  PRINT #UBRpt, TAB(10); " Account# "; Acct$
'  PRINT #UBRpt, TAB(10); "Location# "; Location$;
'  PRINT #UBRpt, TAB(50); "Amount: "; FUsing$(STR$(TotalBal#), "$$,#####.##");
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt, "                            F I N A L  -  N O T I C E"
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt, TAB(10); "Dear Water Customer:"
'  PRINT #UBRpt, TAB(10); ""
'  PRINT #UBRpt, TAB(10); "  We have not received payment of your water bill this month."
'  PRINT #UBRpt, TAB(10); "To prevent your water from being cut off, we must receive"
'  PRINT #UBRpt, TAB(10); "payment before the scheduled cut off date of "; PDate$; ". If"
'  PRINT #UBRpt, TAB(10); "your service is disconnected, the TOTAL AMOUNT DUE, plus a"
'  PRINT #UBRpt, TAB(10); "$25.00 RECONNECT FEE MUST BE PAID BEFORE YOUR SERVICE WILL BE"
'  PRINT #UBRpt, TAB(10); "RESTORED. If you have any questions about your water service,"
'  PRINT #UBRpt, TAB(10); "you can contact Public Utilities, Monday thru Friday, between"
'  PRINT #UBRpt, TAB(10); "8:00 am. and 5:00 pm. before the cut off date to resolve your"
'  PRINT #UBRpt, TAB(10); "billing problems."
'  PRINT #UBRpt, TAB(10); ""
'  PRINT #UBRpt, TAB(10); "  If payment has been made prior to receiving this notice,"
'  PRINT #UBRpt, TAB(10); "Please disregard this notice."
'  PRINT #UBRpt, TAB(10); ""
'  PRINT #UBRpt, TAB(10); "Thank You."
'  PRINT #UBRpt, "~"; TAB(78); "~"
'  PRINT #UBRpt, CHR$(12)

