 'STANDARD REAL ESTATE BILL FORMAT AS SOLD BY SOUTHERN SOFTWARE REMINGTON VA
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
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(5); "ACCT # "; TaxBillRec(1).CustRec
  PRINT #RptFile, TAB(5); LEFT$(CustName$, 45)
  PRINT #RptFile, TAB(5); LEFT$(TaxBillRec(1).CustAdd1, 35)
  PRINT #RptFile, TAB(5); LEFT$(TaxBillRec(1).CustAdd2, 35)
  PRINT #RptFile, TAB(5); QPTrim$(TaxBillRec(1).CustAdd3); " "; TaxBillRec(1).CustZip
  FOR LC = 18 TO 20
    PRINT #RptFile, " "
  NEXT
  PRINT #RptFile, TAB(41); "LAND"; TAB(51); "BUILDING"; TAB(64); "TOTAL"; TAB(72); "TOTAL DUE"
  PRINT #RptFile, " "
 'Line 23 Starts Here
  PRINT #RptFile, QPTrim$(TaxBillRec(1).RDesc1);
  PRINT #RptFile, TAB(31); USING "#.####"; RealTaxRate#;
  PRINT #RptFile, TAB(38); USING "########,#"; (TaxBillRec(1).RealValue - TaxBillRec(1).PersValue);
  PRINT #RptFile, TAB(50); USING "#######,#"; TaxBillRec(1).PersValue;
  PRINT #RptFile, TAB(61); USING "#######,#"; TaxBillRec(1).RealValue;
  PRINT #RptFile, TAB(71); USING "#####,#.##"; TaxBillRec(1).RealTaxDue
  PRINT #RptFile, QPTrim$(TaxBillRec(1).RDesc2)
  
  PRINT #RptFile, "PIN # "; QPTrim$(TaxBillRec(1).REALPIN)

  IF TaxBillRec(1).CarShore > 0 THEN
    PRINT #RptFile, TAB(2); USING "Storm Water: #####.##"; TaxBillRec(1).CarShore
  ELSE
    PRINT #RptFile,
  END IF

  PRINT #RptFile, QPTrim$(TaxBillRec(1).RDesc3)

  'Lines 28 to 36 are blank
  FOR LC = 28 TO 36
    PRINT #RptFile, ""
  NEXT LC
 'Line 37 for Totals
   PRINT #RptFile, ""
   PIYear$ = QPTrim$(STR$(WhatYear))
   PRINT #RptFile, TAB(2); "10% Penalty assessed 06/05/"; PIYear$; TAB(48); "Total Tax Due ... ";
   PRINT #RptFile, USING "$$#####,#.##"; TaxBillRec(1).TotalBillDue
   PRINT #RptFile, TAB(2); "    and 12/05/"; PIYear$; TAB(48); "Tax Due Date: "; DueDate$
   PRINT #RptFile, TAB(2); "10% Interest Accrues Annually"
   PRINT #RptFile, TAB(2); "    After 07/01/"; PIYear$
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile, "~  BN"; USING "#####"; PrnCnt
  

