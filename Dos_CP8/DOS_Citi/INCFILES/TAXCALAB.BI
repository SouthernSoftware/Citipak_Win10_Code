
  TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF
  NetTaxVal# = round#(TaxBill(1).RealValue + TaxBill(1).PersValue)

  PRINT #RptFile, "~"
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(50); TaxBill(1).TaxYear;
  'PRINT #RptFile, USING "#.##"; TAB(32); TaxRate#;
  'PRINT #RptFile, TAB(64); USING "######"; TaxBill(1).CustRec;

  PRINT #RptFile, TAB(58); USING "######"; TaxBill(1).BillNumber;
  PRINT #RptFile, TAB(69); LEFT$(QPTrim$(TaxBill(1).RealPIN), 12)

  PRINT #RptFile,
  PRINT #RptFile, TAB(50); USING "######"; TaxBill(1).CustRec
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(50); USING "########"; TaxBill(1).RealValue;
  PRINT #RptFile, TAB(59); USING "######"; TaxBill(1).PersValue;
  PRINT #RptFile, TAB(69); USING "#########,"; NetTaxVal#
  PRINT #RptFile, 'TAB(73); USING "#####.##"; TaxBill(1).TotalBillDue
  PRINT #RptFile, TAB(69); USING "########"; TaxBill(1).ExptValue
  PRINT #RptFile,
  PRINT #RptFile, TAB(6); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(6); LEFT$(TaxBill(1).CustAdd1, 25); TAB(68); USING "#.##"; TaxRate#
  PRINT #RptFile, TAB(6); LEFT$(TaxBill(1).CustAdd2, 25)
  PRINT #RptFile, TAB(6); QPTrim$(TaxBill(1).CustAdd3); " "; QPTrim$(TaxBill(1).CustZip); TAB(69); USING "#####.##"; TaxBill(1).LateTaxDue
  PRINT #RptFile,
  PRINT #RptFile, TAB(44); LEFT$(TaxBill(1).RDesc1, 21); TAB(69); USING "#####.##"; TaxBill(1).TotalBillDue
  PRINT #RptFile,
  PRINT #RptFile, "~ BN"; USING "#####"; PrnCnt


