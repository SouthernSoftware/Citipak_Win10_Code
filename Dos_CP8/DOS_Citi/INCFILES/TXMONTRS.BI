'  CustName$ = RTRIM$(TaxCust(1).FNAME) + " " + LTRIM$(TaxCust(1).LName)
'  CustName$ = LTRIM$(CustName$)
  PRINT #RptFile,
  PRINT #RptFile, TAB(t1); TownName$
  PRINT #RptFile, TAB(S1); Add1$
  PRINT #RptFile, TAB(S2); ADD2$
  PRINT #RptFile, TAB(S3); ADD3$
  PRINT #RptFile, ""
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); LNDate$
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); "Account: "; CustAcct&
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); CustName$
  PRINT #RptFile, TAB(5); LEFT$(TaxCust(1).Addr1, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxCust(1).Addr2, 25)
  PRINT #RptFile, TAB(5); RTRIM$(TaxCust(1).City) + " "; TaxCust(1).State + " " + LEFT$(TaxCust(1).Zip, 5) + "-" + MID$(TaxCust(1).Zip, 6, 4)
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); "According to our records, your real estate tax bill is past due at"
  PRINT #RptFile, TAB(5); "this time.  Penalty and interest charges have been added. Interest"
  PRINT #RptFile, TAB(5); "charges will continue to be added the 5th of each month."
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); "If you have any questions, Please contact us."
  PRINT #RptFile,
  PRINT #RptFile,
  
  PRINT #RptFile, '"     Bill Number: "; ParseBillNum$(TaxTrans(1).Description)
  PRINT #RptFile, TAB(23); "TAX"; TAB(37); "Penalities"; TAB(57); "Total Due"
  PRINT #RptFile, TAB(15); STRING$(55, "-")
  PRINT #RptFile, USING "######,#.##"; TAB(15); TaxDue#; TAB(36); Round#(InterestDue# + PenaltyDue#); TAB(55); Balance#
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); "Your truly,"
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); "Brenda T. Reamy"
  PRINT #RptFile, CHR$(12);

