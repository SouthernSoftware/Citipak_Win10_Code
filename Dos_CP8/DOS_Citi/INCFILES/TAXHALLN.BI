  'STANDARD REAL ESTATE BILL FORMAT AS SOLD BY SOUTHERN SOFTWARE

   CustName$ = RTRIM$(TaxCust(1).FNAME) + " " + LTRIM$(TaxCust(1).LName)
   CustName$ = LTRIM$(CustName$)

   'Must Calc Late Fee Here
   PRINT #RptFile, "~"
   PRINT #RptFile, TAB(64); "TAX YEAR:"; WhatYear
   PRINT #RptFile, TAB(73); LEFT$(QPTrim$(BillNumber$), 6)
   PRINT #RptFile,
   PRINT #RptFile, " "
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile, " "
   PRINT #RptFile, ""
   PRINT #RptFile, ""
   PRINT #RptFile, TAB(5); "Acct # "; CustAcct&
   PRINT #RptFile, TAB(5); QPTrim$(CustName$)
   PRINT #RptFile, TAB(5); LEFT$(TaxCust(1).Addr1, 25)
   PRINT #RptFile, TAB(5); LEFT$(TaxCust(1).Addr2, 25)
   PRINT #RptFile, TAB(5); QPTrim$(CitySt$); " "; TaxCust(1).Zip
   PRINT #RptFile, " "
   PRINT #RptFile, " "
   PRINT #RptFile, TAB(38); "TAX DUE"; TAB(51); "PENALTY"; TAB(62); "INTEREST"; TAB(75); "TOTAL"
   PRINT #RptFile, " "

   PRINT #RptFile, TAB(2); Desc1$;
   PRINT #RptFile, TAB(38); USING "######.##"; TaxDue#;
   PRINT #RptFile, TAB(51); USING "#####.##"; PenaltyDue#;
   PRINT #RptFile, TAB(62); USING "#####.##"; InterestDue#;
   PRINT #RptFile, TAB(72); USING "######.##"; Balance#
   PRINT #RptFile, TAB(2); Desc2$
   PRINT #RptFile, TAB(2); Desc3$
   PRINT #RptFile, TAB(2); Mess1$
   PRINT #RptFile, TAB(2); Mess2$
   PRINT #RptFile, TAB(2); Mess3$
   PRINT #RptFile, TAB(2); Mess4$
   PRINT #RptFile, TAB(2); Mess5$
   PRINT #RptFile, TAB(2); "Interest Calculated Thru "; InterestDate$
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile, TAB(50); "Unpaid Tax ";
   PRINT #RptFile, TAB(70); USING "$$####,#.##"; TaxDue#
   PRINT #RptFile, TAB(50); "Unpaid Penalty";
   PRINT #RptFile, TAB(70); USING "$$####,#.##"; PenaltyDue#
   PRINT #RptFile, TAB(50); "Unpaid Interest";
   PRINT #RptFile, TAB(70); USING "$$####,#.##"; InterestDue#
   PRINT #RptFile, TAB(50); "Total Balance Due ";
   PRINT #RptFile, TAB(70); USING "$$####,#.##"; Balance#
   PRINT #RptFile, " "
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile, "~"
  

