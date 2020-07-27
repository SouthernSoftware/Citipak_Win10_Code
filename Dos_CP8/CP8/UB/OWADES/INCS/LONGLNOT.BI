'This is longviews
  FOR ZCnt = 1 TO 13
    PRINT #UBRpt, ""
  NEXT
  PRINT #UBRpt, TAB(10); UBCustRec(1).CustName
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR1
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR2
  PRINT #UBRpt, TAB(10); UBCustRec(1).City; "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
  PRINT #UBRpt, TAB(65); NDate$
  FOR ZCnt = 19 TO 24
   PRINT #UBRpt, ""
  NEXT
  PRINT #UBRpt, TAB(25); FUsing$(STR$(CustBal#), ",######.##")
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(37); PDate$
  PRINT #UBRpt, CHR$(12);

