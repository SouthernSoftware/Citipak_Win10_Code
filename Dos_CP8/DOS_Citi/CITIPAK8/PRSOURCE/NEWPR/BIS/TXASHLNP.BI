
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
   PRINT #RptFile, TAB(39); "ORIG TAX"; TAB(52); "PAID/ABATE"; TAB(74); "BALANCE DUE"
   PRINT #RptFile, " "
  'Line 23 Starts Here

  PRINT #RptFile, TAB(2); "Personal Property (QV)";
   PRINT #RptFile, TAB(33); USING ".##"; Rate1#;
   PRINT #RptFile, TAB(37); USING "#####,#.##"; TaxTrans(1).Revenue.Principle1;
   PRINT #RptFile, TAB(51); USING "#####,#.##"; TaxTrans(1).Revenue.Principle1Pd;
   PRINT #RptFile, TAB(74); USING "######,#.##"; TaxTrans(1).Revenue.Principle1 - TaxTrans(1).Revenue.Principle1Pd
  PRINT #RptFile, TAB(2); "Machinery/Tools";
   PRINT #RptFile, TAB(33); USING ".##"; Rate4#;
   PRINT #RptFile, TAB(37); USING "#####,#.##"; TaxTrans(1).Revenue.Principle2;
   PRINT #RptFile, TAB(51); USING "#####,#.##"; TaxTrans(1).Revenue.Principle2Pd;
   PRINT #RptFile, TAB(74); USING "######,#.##"; TaxTrans(1).Revenue.Principle2 - TaxTrans(1).Revenue.Principle2Pd
  PRINT #RptFile, TAB(2); "Merchant Capital";
   PRINT #RptFile, TAB(33); USING ".##"; Rate5#;
   PRINT #RptFile, TAB(37); USING "#####,#.##"; TaxTrans(1).Revenue.Principle3;
   PRINT #RptFile, TAB(51); USING "#####,#.##"; TaxTrans(1).Revenue.Principle3Pd;
   PRINT #RptFile, TAB(74); USING "######,#.##"; TaxTrans(1).Revenue.Principle3 - TaxTrans(1).Revenue.Principle3Pd
  PRINT #RptFile, TAB(2); "Farm Equipment";
   PRINT #RptFile, TAB(33); USING ".##"; Rate2#;
   PRINT #RptFile, TAB(37); USING "#####,#.##"; TaxTrans(1).Revenue.Principle4;
   PRINT #RptFile, TAB(51); USING "#####,#.##"; TaxTrans(1).Revenue.Principle4Pd;
   PRINT #RptFile, TAB(74); USING "######,#.##"; TaxTrans(1).Revenue.Principle4 - TaxTrans(1).Revenue.Principle4Pd
 
  PRINT #RptFile, TAB(2); "Mobile Home";
   PRINT #RptFile, TAB(33); USING ".##"; Rate3#;
   PRINT #RptFile, TAB(37); USING "#####,#.##"; TaxTrans(1).Revenue.Principle5;
   PRINT #RptFile, TAB(51); USING "#####,#.##"; TaxTrans(1).Revenue.Principle5Pd;
   PRINT #RptFile, TAB(74); USING "######,#.##"; TaxTrans(1).Revenue.Principle5 - TaxTrans(1).Revenue.Principle5Pd
  PRINT #RptFile, TAB(2); "Late Payment Penalty ";
   PRINT #RptFile, TAB(37); USING "#####,#.##"; Balance# * .1;
   PRINT #RptFile, TAB(74); USING "######,#.##"; Balance# * .1

   PenaltyDue# = Balance# * .1

 'Lines 29 to 35 are blank
   FOR LCnt! = 29 TO 35: PRINT #RptFile, "": NEXT LCnt!
 'Line 36 for Totals
   PRINT #RptFile, TAB(50); "Unpaid Tax ";
     PRINT #RptFile, TAB(74); USING "$$####,#.##"; Balance#
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
  

