  TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF
  
  PRINT #RptFile, "~"
  PRINT #RptFile, TAB(27); TaxBill(1).TaxYear;
   PRINT #RptFile, USING "#.##"; TAB(36); TaxRate#;
   PRINT #RptFile, TAB(66); USING "######"; TaxBill(1).CustRec;
   PRINT #RptFile, TAB(78); USING "#####"; TaxBill(1).BillNumber
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
      NetTaxVal# = round#(TaxBill(1).RealValue + TaxBill(1).PersValue)
  PRINT #RptFile, USING "########,#"; TaxBill(1).PersValue; TAB(12); TaxBill(1).RealValue; TAB(23); NetTaxVal#;
      NetTaxVal# = round#(NetTaxVal# - TaxBill(1).ExptValue)
   PRINT #RptFile, TAB(34); USING "######.##"; (TaxBill(1).TotalBillDue - TaxBill(1).LateTaxDue);
   PRINT #RptFile, TAB(44); USING "######,#"; TaxBill(1).ExptValue;
   PRINT #RptFile, TAB(70); USING "#####.##"; TaxBill(1).LateTaxDue;
   PRINT #RptFile, TAB(78); USING "#####.##"; TaxBill(1).TotalBillDue
  PRINT #RptFile,
  PRINT #RptFile, " BN"; USING "#####"; PrnCnt
  PRINT #RptFile, TAB(2); LEFT$(TaxBill(1).RDesc1, 21)
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxBill(1).CustAdd1, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxBill(1).CustAdd2, 25)
  PRINT #RptFile, TAB(5); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, "~"


