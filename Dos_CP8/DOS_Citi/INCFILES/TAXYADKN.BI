  TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF
  
  PRINT #RptFile, CHR$(27); CHR$(77); "~"
  PRINT #RptFile, TAB(33); TaxBill(1).TaxYear;
  PRINT #RptFile, USING "#.##"; TAB(45); TaxRate#;
  PRINT #RptFile, TAB(76); USING "######"; TaxBill(1).CustRec;
  PRINT #RptFile, TAB(90); USING "######"; TaxBill(1).BillNumber
  PRINT #RptFile,
  PRINT #RptFile, 'TAB(68); USING "######"; TaxBill(1).CustRec
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  NetTaxVal# = round#(TaxBill(1).RealValue + TaxBill(1).PersValue)
  PRINT #RptFile, USING "#######,#"; TaxBill(1).PersValue;
  PRINT #RptFile, TAB(14); USING "#######,#"; TaxBill(1).RealValue;
  PRINT #RptFile, TAB(27); USING "#######,#"; NetTaxVal#;
  NetTaxVal# = round#(NetTaxVal# - TaxBill(1).ExptValue)
  PRINT #RptFile, TAB(41); USING "######.##"; (TaxBill(1).TotalBillDue - TaxBill(1).LateTaxDue);
  PRINT #RptFile, TAB(53); USING "######,#"; TaxBill(1).ExptValue;
  PRINT #RptFile, TAB(81); USING "####.##"; TaxBill(1).LateTaxDue;
  PRINT #RptFile, TAB(92); USING "####.##"; TaxBill(1).TotalBillDue
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, 'TAB(5); "IF TAXES ARE ESCROWED SEND BILL TO"
  PRINT #RptFile, 'TAB(5); "MORTGAGE COMPANY."
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxBill(1).CustAdd1, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxBill(1).CustAdd2, 25)
  PRINT #RptFile, TAB(5); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, "~"


