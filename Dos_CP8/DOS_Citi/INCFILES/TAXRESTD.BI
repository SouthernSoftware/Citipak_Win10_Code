 'STANDARD REAL ESTATE BILL FORMAT AS SOLD BY SOUTHERN SOFTWARE
  CLOSE TaxFile
  TaxFile = FREEFILE
  OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, TaxBillRec(1).CustRec, TaxCust(1)

  CustName$ = RTRIM$(TaxCust(1).FNAME) + " " + LTRIM$(TaxCust(1).LName)
  CustName$ = LTRIM$(CustName$)


  'Must Calc Late Fee Here
  

  PRINT #RptFile, "~"
  PRINT #RptFile, TAB(64); "TAX YEAR: "; WhatYear
  PRINT #RptFile, TAB(75); USING "#####"; TaxBillRec(1).BillNumber
  PRINT #RptFile, " "
  PRINT #RptFile, " "
  PRINT #RptFile, TAB(5); TownName$
  PRINT #RptFile, TAB(5); Add1$
  PRINT #RptFile, TAB(5); Add2$
  PRINT #RptFile, TAB(5); Add3$
  PRINT #RptFile, " "
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(5); " PIN: "; PINTemp(1)
  PRINT #RptFile, TAB(5); USING "ACCT: ####"; TaxBillRec(1).CustRec
  PRINT #RptFile, TAB(5); LEFT$(CustName$, 45)
  PRINT #RptFile, TAB(5); LEFT$(TaxBillRec(1).CustAdd1, 35)
  PRINT #RptFile, TAB(5); LEFT$(TaxBillRec(1).CustAdd2, 35)
  PRINT #RptFile, TAB(5); qptrim$(TaxBillRec(1).CustAdd3); " "; TaxBillRec(1).CustZip
  
  FOR LC = 18 TO 20
   PRINT #RptFile, " "
  NEXT LC
   PRINT #RptFile, TAB(41); "LAND"; TAB(51); "BUILDING"; TAB(64); "TOTAL"; TAB(72); "TOTAL DUE"
   PRINT #RptFile, " "
  'Line 23 Starts Here
  PRINT #RptFile, qptrim$(TaxBillRec(1).RDesc1);
   PRINT #RptFile, TAB(31); USING "#.##"; RealTaxRate#;
   PRINT #RptFile, TAB(37); USING "########,#"; (TaxBillRec(1).RealValue - TaxBillRec(1).PersValue);
   PRINT #RptFile, TAB(50); USING "#######,#"; TaxBillRec(1).PersValue;
   PRINT #RptFile, TAB(61); USING "#######,#"; TaxBillRec(1).RealValue;
   PRINT #RptFile, TAB(71); USING "#####,#.##"; TaxBillRec(1).RealTaxDue
  PRINT #RptFile, qptrim$(TaxBillRec(1).RDesc2)

  'Lines 25 to 36 are blank
   FOR LCnt! = 25 TO 36: PRINT #RptFile, "": NEXT LCnt!
 'Line 37 for Totals
   PRINT #RptFile, ""
   PRINT #RptFile, TAB(48); "Total Tax Due ... ";
    PRINT #RptFile, USING "$$#####,#.##"; TaxBillRec(1).TotalBillDue
   PRINT #RptFile, TAB(48); "Tax Due Date: "; DueDate$
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile, "BN"; USING "#####"; PrnCnt
   PRINT #RptFile, "~"
  

