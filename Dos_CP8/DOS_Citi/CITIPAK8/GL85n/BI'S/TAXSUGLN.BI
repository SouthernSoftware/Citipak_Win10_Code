  GOSUB GetRealInfo
  GOSUB GetPersInfo

  BillNum$ = ParseBillNum$(TaxTransRec.Description)
  BillNum$ = LEFT$(BillNum$, 9)

  CalcVal# = Round#(RealCalcVal# + PersCalcVal#)

  PRINT #RptFile, "~"
  PRINT #RptFile, ' TAB(36); "Lien"
  PRINT #RptFile, ' TAB(36); "Advertising"
  PRINT #RptFile, ' TAB(36); "Next Month in"
  PRINT #RptFile, ' TAB(36); "Mount Olive Tribune"
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, WhatYear;
  PRINT #RptFile, " "; BillNum$;
  PRINT #RptFile, TAB(17); USING "######"; CustAcct&;
  PRINT #RptFile, " "; PINTemp(1);
  PRINT #RptFile, TAB(36); USING "####"; WhatYear;
  PRINT #RptFile, TAB(41); BillNum$;
  PRINT #RptFile, TAB(51); USING "######"; CustAcct&
  PRINT #RptFile, TAB(12); LEFT$(CustName$, 19)
  PRINT #RptFile, TAB(12); TaxTransRec.Description
  PRINT #RptFile, TAB(12); "PAST DUE NOTICE"
 'v line 12
  PRINT #RptFile,
  'PRINT #RptFile,
  PRINT #RptFile, TAB(2); USING "########,#"; CalcVal#
  PRINT #RptFile,
  PRINT #RptFile, TAB(36); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(2); USING "########,#"; CalcVal#;
  PRINT #RptFile, "    0.65";
  PRINT #RptFile, TAB(20); USING "#######,#.##"; Balance#;
  PRINT #RptFile, TAB(36); LEFT$(TaxCustRec.Addr1, 25)
  PRINT #RptFile, TAB(36); LEFT$(TaxCustRec.Addr2, 25)
  PRINT #RptFile, TAB(36); RTRIM$(TaxCustRec.City) + " "; TaxCustRec.State + " " + LEFT$(TaxCustRec.Zip, 5) + "-" + MID$(TaxCustRec.Zip, 6, 4)
  PRINT #RptFile,
  PRINT #RptFile, TAB(36); "LATE NOTICE"
  PRINT #RptFile, TAB(20); USING "#######,#.##"; Balance#
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, "~"

