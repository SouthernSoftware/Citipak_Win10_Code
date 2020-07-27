'North Lenoir
  PRINT #UBRpt, "    North Lenoir Water Corp."
  PRINT #UBRpt, "    P.O. Box 1564"
  PRINT #UBRpt, "    Kinston, NC 28503"
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    Forwarding & address correction requested."
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(5); LongNDate$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    "; UBCustRec(1).CustName
  PRINT #UBRpt, "    "; UBCustRec(1).ADDR1
  PRINT #UBRpt, "    "; UBCustRec(1).ADDR2
  PRINT #UBRpt, "    "; QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    ACCOUNT: "; Acct$; TAB(40); "LOC: "; UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    YOUR ACCOUNT IS DUE FOR TWO MONTHS.  YOU MUST PAY THE TOTAL AMOUNT"
  PRINT #UBRpt, "    DUE, AND A LATE FEE OF $15.00.  IF NOT PAID BY 10:00 AM ON THE DATE"
  PRINT #UBRpt, "    BELOW, YOU WILL OWE AN ADDITIONAL CHARGE OF $20.00 AND YOUR SERVICE"
  PRINT #UBRpt, "    WILL BE DISCONTINUED."
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(32); "        DUE DATE:  "; PDate$
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(32); "BEFORE 10:00 PAY:"; USING "#####.##"; TotalBal#
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(32); " AFTER 10:00 PAY:"; USING "#####.##"; Round#(TotalBal# + 20)
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "    FOR YOUR CONVENIENCE, YOU MAY USE THE PAYMENT DROP LOCATED AT THE"
  PRINT #UBRpt, "    FRONT OF OUR OFFICE."
  PRINT #UBRpt, CHR$(12)

