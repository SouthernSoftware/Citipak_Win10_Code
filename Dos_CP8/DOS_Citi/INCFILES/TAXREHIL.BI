 'STANDARD REAL ESTATE BILL FORMAT AS SOLD BY SOUTHERN SOFTWARE HILLSVILLE VA
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
  NEXT LC
  PRINT #RptFile, TAB(41); "LAND"; TAB(51); "BUILDING"; TAB(64); "TOTAL"; TAB(72); "TOTAL DUE"
  PRINT #RptFile, " "
  'Line 23 Starts Here
  PRINT #RptFile, QPTrim$(TaxBillRec(1).RDesc1);
  PRINT #RptFile, TAB(31); USING "#.##"; RealTaxRate#;
  PRINT #RptFile, TAB(37); USING "########,#"; (TaxBillRec(1).RealValue - TaxBillRec(1).PersValue);
  PRINT #RptFile, TAB(50); USING "#######,#"; TaxBillRec(1).PersValue;
  PRINT #RptFile, TAB(61); USING "#######,#"; TaxBillRec(1).RealValue;
  PRINT #RptFile, TAB(71); USING "#####,#.##"; TaxBillRec(1).RealTaxDue
  PRINT #RptFile, QPTrim$(TaxBillRec(1).RDesc2)
  PRINT #RptFile, QPTrim$(TaxBillRec(1).RDesc3)

  'Lines 25 to 36 are blank
   FOR LCnt! = 26 TO 36: PRINT #RptFile, "": NEXT LCnt!
 
   PRINT #RptFile,
   PRINT #RptFile, TAB(48); "Total Tax Due ... "; USING "$$#####,#.##"; TaxBillRec(1).TotalBillDue
   PRINT #RptFile, TAB(48); "Tax Due Date: "; DueDate$
   PRINT #RptFile, "5% Penalty after 12/05/"; QPTrim$(STR$(WhatYear)); "."
   PRINT #RptFile, "10% Interest Accrues After 01/01/"; QPTrim$(STR$(WhatYear + 1)); "."
   PRINT #RptFile,
   PRINT #RptFile, "BN"; USING "#####"; PrnCnt
   PRINT #RptFile, "~"
  
'   PRINT #RptFile, "5% Penalty if paid after 12/05/"; QPTrim$(STR$(WhatYear)); "."
'   PRINT #RptFile, "Additional 7% after 01/20/"; QPTrim$(STR$(WhatYear + 1)); "."; TAB(48); "Total Tax Due ... "; USING "$$#####,#.##"; TaxBillRec(1).TotalBillDue
'   PRINT #RptFile, "Additional 15% after 04/15/"; QPTrim$(STR$(WhatYear + 1)); "."; TAB(48); "Tax Due Date: "; DueDate$
'   PRINT #RptFile,
'   PRINT #RptFile, "15% Annum Interest beginning 01/03/"; QPTrim$(STR$(WhatYear + 1)); "."
'   PRINT #RptFile,
'   PRINT #RptFile, "BN"; USING "#####"; PrnCnt
'   PRINT #RptFile, "~"

