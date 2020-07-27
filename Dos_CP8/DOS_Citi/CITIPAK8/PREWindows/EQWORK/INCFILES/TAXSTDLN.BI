  PRINT #RptFile, "~"
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, WhatYear;
   PRINT #RptFile, TAB(15); USING "######"; CustAcct&;
   PRINT #RptFile, TAB(36); USING "####"; WhatYear;
   PRINT #RptFile, TAB(51); USING "######"; CustAcct&
  PRINT #RptFile, TAB(12); LEFT$(CustName$, 21)
  PRINT #RptFile, TAB(12); TaxTransRec.Description
  PRINT #RptFile, TAB(12); "PAST DUE NOTICE"
 'v line 12
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(36); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(20); USING "#######,#.##"; Balance#;
   PRINT #RptFile, TAB(36); LEFT$(TaxCustRec.Addr1, 25)
  PRINT #RptFile, TAB(36); LEFT$(TaxCustRec.Addr2, 25)
  PRINT #RptFile, TAB(36); RTRIM$(TaxCustRec.City) + " "; TaxCustRec.State + " " + LEFT$(TaxCustRec.Zip, 5) + "-" + MID$(TaxCustRec.Zip, 6, 4)
  PRINT #RptFile,
  PRINT #RptFile,
   PRINT #RptFile, TAB(20); USING "#######,#.##"; Balance#
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, "~"

