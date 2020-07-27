
  CLOSE TaxFile
  TaxFile = FREEFILE
  OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, TaxBillRec(1).CustRec, TaxCust(1)

  CustName$ = RTRIM$(TaxCust(1).FNAME) + " " + LTRIM$(TaxCust(1).LName)
  CustName$ = LTRIM$(CustName$)

  'Must Calc Late Fee Here

  PRINT #RptFile, "~"
  PRINT #RptFile, TAB(50); WhatYear; TAB(75); USING "########"; TaxBillRec(1).BillNumber
  PRINT #RptFile,
  PRINT #RptFile, " "
  PRINT #RptFile, " "
  PRINT #RptFile, " "
  PRINT #RptFile, TAB(28); USING "#.##"; RealTaxRate#;
  PRINT #RptFile, TAB(36); USING "########,#"; (TaxBillRec(1).RealValue - TaxBillRec(1).PersValue);
  PRINT #RptFile, TAB(48); USING "#######,#"; TaxBillRec(1).PersValue;
  PRINT #RptFile, TAB(61); USING "#######,#"; TaxBillRec(1).RealValue;
  PRINT #RptFile, TAB(71); USING "#####,#.##"; TaxBillRec(1).RealTaxDue;
  PRINT #RptFile, "    5%"
  PRINT #RptFile, " "
  PRINT #RptFile, TAB(68); DueDate$
  PRINT #RptFile, TAB(3); QPTrim$(TaxBillRec(1).RDesc1)
  PRINT #RptFile, TAB(3); QPTrim$(TaxBillRec(1).RDesc2)
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(10); "ACCT # "; TaxBillRec(1).CustRec
  PRINT #RptFile, TAB(10); LEFT$(CustName$, 45)
  PRINT #RptFile, TAB(10); LEFT$(TaxBillRec(1).CustAdd1, 35)
  PRINT #RptFile, TAB(10); LEFT$(TaxBillRec(1).CustAdd2, 35)
  PRINT #RptFile, TAB(10); QPTrim$(TaxBillRec(1).CustAdd3); " "; TaxBillRec(1).CustZip

  'PRINT #RptFile, USING "$$#####,#.##"; TaxBillRec(1).TotalBillDue
  'PRINT #RptFile, TAB(48); "Tax Due Date: "; DueDate$
  PRINT #RptFile,
  PRINT #RptFile, "BN"; USING "#####"; PrnCnt
  PRINT #RptFile, "~"
  

