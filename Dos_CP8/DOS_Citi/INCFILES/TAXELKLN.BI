  'STANDARD REAL ESTATE BILL FORMAT AS SOLD BY SOUTHERN SOFTWARE

   CustName$ = RTRIM$(TaxCust(1).FNAME) + " " + LTRIM$(TaxCust(1).LName)
   CustName$ = LTRIM$(CustName$)

   'Must Calc Late Fee Here
   PRINT #RptFile, "~"
   PRINT #RptFile, TAB(t1); TownName$
   PRINT #RptFile, TAB(S1); Add1$
   PRINT #RptFile, TAB(S2); ADD2$
   PRINT #RptFile, TAB(S3); ADD3$
   PRINT #RptFile, 'TAB(73); LEFT$(QPTrim$(BillNumber$), 6)
   PRINT #RptFile,
   PRINT #RptFile, " "
   PRINT #RptFile, 'TAB(5); "TAX YEAR:"; WhatYear
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile, TAB(5); "Acct # "; CustAcct&
   PRINT #RptFile, TAB(5); QPTrim$(CustName$)
   PRINT #RptFile, TAB(5); LEFT$(TaxCust(1).Addr1, 25)
   PRINT #RptFile, TAB(5); LEFT$(TaxCust(1).Addr2, 25)
   PRINT #RptFile, TAB(5); QPTrim$(CitySt$); " "; TaxCust(1).Zip
   PRINT #RptFile, " "
   PRINT #RptFile, " "
   PRINT #RptFile, TAB(30); "BILLYEAR"; TAB(40); "TAX DUE"; TAB(52); "PENALTY"; TAB(62); "INTEREST"; TAB(76); "TOTAL"
   PRINT #RptFile, " "

   PRINT #RptFile, TAB(2); Desc1$;
   PRINT #RptFile, TAB(34); QPTrim$(TaxYear$);
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
   PRINT #RptFile, TAB(40); "Unpaid Tax "; TAB(60); USING "$$####,#.##"; TaxDue#
   PRINT #RptFile, TAB(40); "Unpaid Penalty"; TAB(60); USING "$$####,#.##"; PenaltyDue#
   PRINT #RptFile, TAB(40); "Unpaid Interest"; TAB(60); USING "$$####,#.##"; InterestDue#
   PRINT #RptFile, TAB(40); "Total Balance Due "; TAB(60); USING "$$####,#.##"; Balance#
   PRINT #RptFile, " "
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile, CHR$(12)
  

