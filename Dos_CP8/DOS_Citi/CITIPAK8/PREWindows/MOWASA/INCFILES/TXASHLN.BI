  CLOSE RealFile
  RealFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS RealFile LEN = RealRecLen

  CustName$ = RTRIM$(TaxCust(1).LName) + "  " + RTRIM$(TaxCust(1).FNAME)
  CustName$ = LTRIM$(CustName$)


  'Must Calc Late Fee Here
  

  PRINT #RptFile, "~"
  PRINT #RptFile, TAB(75); BillNumber$
  PRINT #RptFile, TAB(75); WhatYear
  FOR LC = 4 TO 10
   PRINT #RptFile, " "
  NEXT LC
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(5); "Acct # "; CustAcct&
  PRINT #RptFile, TAB(5); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxCust(1).Addr1, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxCust(1).Addr2, 25)
  PRINT #RptFile, TAB(5); qptrim$(CitySt$); " "; TaxCust(1).Zip
  FOR LC = 18 TO 20
   PRINT #RptFile, " "
  NEXT LC
   PRINT #RptFile, TAB(39); "TAX DUE"; TAB(52); "PENALTY"; TAB(64); "INTEREST"; TAB(74); "TOTAL DUE"
   PRINT #RptFile, " "
  'Line 23 Starts Here
  PRINT #RptFile, TAB(2); Desc1$;
   PRINT #RptFile, TAB(37); USING "#####,#.##"; TaxDue#;
   PRINT #RptFile, TAB(51); USING "#####,#.##"; PenaltyDue#;
   PRINT #RptFile, TAB(62); USING "#####,#.##"; InterestDue#;
   PRINT #RptFile, TAB(74); USING "######,#.##"; Balance# + PenaltyDue# + InterestDue#
  PRINT #RptFile, TAB(2); Desc2$
  PRINT #RptFile, TAB(2); Desc3$

 'Lines 26 to 35 are blank
   FOR LCnt! = 26 TO 35: PRINT #RptFile, "": NEXT LCnt!
 'Line 36 for Totals
   PRINT #RptFile, TAB(50); "Unpaid Tax ";
     PRINT #RptFile, TAB(74); USING "$$####,#.##"; TaxDue#
    PRINT #RptFile, TAB(50); "Unpaid Penalty";
     PRINT #RptFile, TAB(74); USING "$$####,#.##"; PenaltyDue#
    PRINT #RptFile, TAB(50); "Interest Thru "; InterestDate$;
     PRINT #RptFile, TAB(74); USING "$$####,#.##"; InterestDue#
    PRINT #RptFile, TAB(50); "Total Balance Due ";
     PRINT #RptFile, TAB(74); USING "$$####,#.##"; Balance# + PenaltyDue# + InterestDue#
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile, "~"
  

