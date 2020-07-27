
  CLOSE TaxFile
  TaxFile = FREEFILE
  OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, TaxBillRec(1).CustRec, TaxCust(1)

  CustName$ = RTRIM$(TaxCust(1).FNAME) + " " + LTRIM$(TaxCust(1).LName)
  CustName$ = LTRIM$(CustName$)
  
  'Must Calc Late Fee Here
  
  PRINT #RptFile, "~"
  PRINT #RptFile, TAB(75); USING "#####"; TaxBillRec(1).BillNumber
  PRINT #RptFile, TAB(75); WhatYear
  PRINT #RptFile, " "
  PRINT #RptFile, " "
  PRINT #RptFile, TAB(5); TownName$
  PRINT #RptFile, TAB(5); Add1$
  PRINT #RptFile, TAB(5); Add2$
  PRINT #RptFile, TAB(5); Add3$
  PRINT #RptFile, " "
  PRINT #RptFile, " "
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(5); "Acct # "; TaxBillRec(1).CustRec
  PRINT #RptFile, TAB(5); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxBillRec(1).CustAdd1, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxBillRec(1).CustAdd2, 25)
  PRINT #RptFile, TAB(5); QPTrim$(TaxBillRec(1).CustAdd3); " "; TaxBillRec(1).CustZip
  PRINT #RptFile, " "
  PRINT #RptFile, " "
  PRINT #RptFile, " "
  PRINT #RptFile, TAB(43); "LAND"; TAB(51); "BUILDING"; TAB(65); "TOTAL"; TAB(72); "TOTAL DUE"
  PRINT #RptFile, " "
 'Line 23 Starts Here
  PRINT #RptFile, TAB(2); QPTrim$(TaxBillRec(1).RDesc1);
  PRINT #RptFile, TAB(33); USING "#.##"; RealTaxRate#;
  PRINT #RptFile, TAB(37); USING "########,#"; (TaxBillRec(1).RealValue - TaxBillRec(1).PersValue);
  PRINT #RptFile, TAB(50); USING "#######,#"; TaxBillRec(1).PersValue;
  PRINT #RptFile, TAB(61); USING "#######,#"; (TaxBillRec(1).RealValue);
  PRINT #RptFile, TAB(70); USING "######,#.##"; (TaxBillRec(1).RealTaxDue)
  PRINT #RptFile, TAB(2); QPTrim$(TaxBillRec(1).RDesc2)
 'Lines 25 to 36 are blank
  FOR LCnt! = 25 TO 36
    PRINT #RptFile, ""
  NEXT
 'Line 37 for Totals
   PRINT #RptFile,
   PRINT #RptFile, TAB(3); PINTemp(1)
   PRINT #RptFile,
   PRINT #RptFile, TAB(54); "   Tax Due Date: "; DueDate$
   PRINT #RptFile, TAB(60); USING "Total Due ######,#.##"; TaxBillRec(1).TotalBillDue
   PRINT #RptFile, "BN"; USING "#####"; PrnCnt
   PRINT #RptFile, "~"
  

