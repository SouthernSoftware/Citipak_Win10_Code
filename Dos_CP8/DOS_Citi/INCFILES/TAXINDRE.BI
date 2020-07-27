 'Independence VA Plain Paper
  CLOSE TaxFile
  TaxFile = FREEFILE
  OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, TaxBillRec(1).CustRec, TaxCust(1)

  CustName$ = RTRIM$(TaxCust(1).FNAME) + " " + LTRIM$(TaxCust(1).LName)
  CustName$ = LTRIM$(CustName$)

  'Must Calc Late Fee Here

  PRINT #RptFile,
  PRINT #RptFile, TAB(30); "TOWN OF INDEPENDENCE"
  PRINT #RptFile, TAB(30); "     PO BOX 99      "
  PRINT #RptFile, TAB(29); "INDEPENDENCE, VA 24348"
  PRINT #RptFile, TAB(30); "REAL ESTATE TAX BILL"
  FOR LC = 6 TO 8
   PRINT #RptFile, " "
  NEXT LC
  PRINT #RptFile, TAB(12); "ACCT # "; TaxBillRec(1).CustRec;
    PRINT #RptFile, TAB(65); "BILL #"; USING "######"; TaxBillRec(1).BillNumber
  PRINT #RptFile, TAB(12); LEFT$(CustName$, 25);
   PRINT #RptFile, TAB(63); "TAX YEAR "; WhatYear
  PRINT #RptFile, TAB(12); LEFT$(TaxBillRec(1).CustAdd1, 25);
   PRINT #RptFile, TAB(63); "TAX RATE "; USING "#.###"; RealTaxRate#
  PRINT #RptFile, TAB(12); LEFT$(TaxBillRec(1).CustAdd2, 25)
  PRINT #RptFile, TAB(12); QPTRIM$(TaxBillRec(1).CustAdd3); " "; LEFT$(TaxBillRec(1).CustZip, 5) + "-" + MID$(TaxBillRec(1).CustZip, 6, 4)
  FOR LC = 14 TO 17
   PRINT #RptFile, " "
  NEXT LC
   PRINT #RptFile, TAB(40); "[-------- VALUATIONS --------]"
   PRINT #RptFile, TAB(2); "PROPERTY DESCRIPTION"; TAB(33); "RATE"; TAB(43); "LAND"; TAB(51); "BUILDING"; TAB(64); "TOTAL"; TAB(72); "TOTAL DUE"
   PRINT #RptFile, " "
  'Line 23 Starts Here
  PRINT #RptFile, TAB(2); QPTRIM$(TaxBillRec(1).RDesc1);
   PRINT #RptFile, TAB(33); USING ".##"; RealTaxRate#;
   PRINT #RptFile, TAB(37); USING "########,#"; (TaxBillRec(1).RealValue - TaxBillRec(1).PersValue);
   PRINT #RptFile, TAB(50); USING "#######,#"; TaxBillRec(1).PersValue;
   PRINT #RptFile, TAB(62); USING "######,#"; (TaxBillRec(1).RealValue);
   PRINT #RptFile, TAB(71); USING "#####,#.##"; (TaxBillRec(1).RealTaxDue)
  PRINT #RptFile, TAB(2); QPTRIM$(TaxBillRec(1).RDesc2)

  'Lines 25 to 36 are blank
   FOR LCnt! = 25 TO 36: PRINT #RptFile, "": NEXT LCnt!
 'Line 37 for Totals
   PRINT #RptFile, ""
    PRINT #RptFile, TAB(39); "Total Tax Due by "; DueDate$;
    PRINT #RptFile, TAB(70); USING "$$####,#.##"; TaxBillRec(1).TotalBillDue
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile, "BN"; USING "#####"; PrnCnt
   PRINT #RptFile, CHR$(12);
  

