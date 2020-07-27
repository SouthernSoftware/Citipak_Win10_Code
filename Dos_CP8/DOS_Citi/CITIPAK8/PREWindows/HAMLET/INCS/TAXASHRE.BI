'  CLOSE RealFile
'  RealFile = FREEFILE
'  OPEN TaxPropFile FOR RANDOM SHARED AS RealFile LEN = RealRecLen

  CLOSE TaxFile
  TaxFile = FREEFILE
  OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, TaxBillRec(1).CustRec, TaxCust(1)

  CustName$ = RTRIM$(TaxCust(1).LName) + "  " + RTRIM$(TaxCust(1).FNAME)
  CustName$ = LTRIM$(CustName$)


  'Must Calc Late Fee Here
  

  PRINT #RptFile, "~"
  PRINT #RptFile, TAB(75); USING "#####"; TaxBillRec(1).BillNumber
  PRINT #RptFile, TAB(75); WhatYear
  FOR LC = 4 TO 10
   PRINT #RptFile, " "
  NEXT LC
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(5); "Acct # "; TaxBillRec(1).CustRec
  PRINT #RptFile, TAB(5); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxBillRec(1).CustAdd1, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxBillRec(1).CustAdd2, 25)
  PRINT #RptFile, TAB(5); qptrim$(TaxBillRec(1).CustAdd3); " "; TaxBillRec(1).CustZip
  FOR LC = 18 TO 20
   PRINT #RptFile, " "
  NEXT LC
   PRINT #RptFile, TAB(43); "LAND"; TAB(52); "BUILDING"; TAB(66); "TOTAL"; TAB(75); "TOTAL DUE"
   PRINT #RptFile, " "
  'Line 23 Starts Here
  PRINT #RptFile, TAB(2); qptrim$(TaxBillRec(1).RDesc1);
   PRINT #RptFile, TAB(33); USING ".##"; RealTaxRate#;
   PRINT #RptFile, TAB(37); USING "########,#"; (TaxBillRec(1).RealValue - TaxBillRec(1).PersValue);
   PRINT #RptFile, TAB(51); USING "#######,#"; TaxBillRec(1).PersValue;
   PRINT #RptFile, TAB(62); USING "#######,#"; (TaxBillRec(1).RealValue);
   PRINT #RptFile, TAB(74); USING "######,#.##"; (TaxBillRec(1).RealTaxDue)
  PRINT #RptFile, TAB(2); qptrim$(TaxBillRec(1).RDesc2)

  'Lines 25 to 36 are blank
   FOR LCnt! = 25 TO 36: PRINT #RptFile, "": NEXT LCnt!
 'Line 37 for Totals
   PRINT #RptFile, ""
    PRINT #RptFile, TAB(49); "Total Tax Due "; DueDate$;
    PRINT #RptFile, TAB(74); USING "$$####,#.##"; TaxBillRec(1).TotalBillDue
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile, "BN"; USING "#####"; PrnCnt
   PRINT #RptFile, "~"
  

