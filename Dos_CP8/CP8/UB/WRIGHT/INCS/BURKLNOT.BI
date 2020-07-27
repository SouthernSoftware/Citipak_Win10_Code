  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(33); "Town of Burkeville"
  PRINT #UBRpt, TAB(36); "P.O. Box 277"
  PRINT #UBRpt, TAB(32); "Burkeville, VA 23922"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); LongNDate$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).CustName)
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR1; TAB(53);
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR2
  PRINT #UBRpt, TAB(10); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; Zip$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "RE: Account # "; Location$; TAB(45); "AMOUNT DUE "; FUsing$(STR$(TotalBal#), "$$,######.##")
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(10); "Dear Customer:"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(15); "Our records indicate your account with the Water Department"
  PRINT #UBRpt, TAB(10); "of the Town of Burkeville is now due.  To avoid cut-off, please "
  PRINT #UBRpt, TAB(10); "pay the above amount at the Town Office by "; PDate$
  PRINT #UBRpt,
  PRINT #UBRpt, CHR$(12);
  

