
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(33); "Town of Burkeville"
  PRINT #UBRpt, TAB(35); "P.O. Box 277"
  PRINT #UBRpt, TAB(32); "Burkeville, VA 23922"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(8); LongNDate$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(8); UBCustRec(1).CustName
  PRINT #UBRpt, TAB(8); UBCustRec(1).ADDR1
  PRINT #UBRpt, TAB(8); UBCustRec(1).ADDR2
  PRINT #UBRpt, TAB(8); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(8); "RE: Account # "; Acct$; TAB(45); "AMOUNT DUE "; USING "$#####,.##"; CustBal#
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(8); "Dear Customer:"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(13); "Our records indicate your account with the Water Department"
  PRINT #UBRpt, TAB(8); "of the Town of Burkeville is now due.  To avoid cut-off, please"
  PRINT #UBRpt, TAB(8); "pay the above amount at the Town Office by "; LongPDate$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(8); "Thank you,"
  PRINT #UBRpt, TAB(8); "Town of Burkeville"
  PRINT #UBRpt, TAB(8); "Water Department"
  PRINT #UBRpt, CHR$(12);

