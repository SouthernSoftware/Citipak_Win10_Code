
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(33); "Town of Dillwyn"
  PRINT #UBRpt, TAB(35); "P.O. Box 249"
  PRINT #UBRpt, TAB(32); "Dillwyn, VA 23936"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(7); LongNDate$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(7); UBCustRec(1).CustName
  PRINT #UBRpt, TAB(7); UBCustRec(1).ADDR1
  PRINT #UBRpt, TAB(7); UBCustRec(1).ADDR2
  PRINT #UBRpt, TAB(7); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode; TAB(55); "Loc: "; Location$

  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(7); "RE: Account #"; AcctNo&; TAB(45); "AMOUNT DUE "; USING "$#####,.##"; TotalBal#
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(7); "Dear Customer:"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(12); "Our records indicate your account with the Utility Department"
  PRINT #UBRpt, TAB(7); "of the Town of Dillwyn is now due.  To avoid cut-off, please "
  PRINT #UBRpt, TAB(7); "pay the above amount at the Town Office by "; LongPDate$
  PRINT #UBRpt,
  PRINT #UBRpt, CHR$(12);

