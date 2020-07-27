  TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF

  PRINT #RptFile, "~"; TAB(50); USING "####"; PrnCnt
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TaxBill(1).TaxYear;
  PRINT #RptFile, TAB(7); USING "######"; TaxBill(1).BillNumber;
  PRINT #RptFile, TAB(16); USING "#####"; TaxBill(1).CustRec;
  PRINT #RptFile, TAB(23); PINTemp(1); TAB(37); USING "####"; TaxBill(1).TaxYear;

  PRINT #RptFile, TAB(42); USING "########"; TaxBill(1).CustRec;
  'PRINT #RptFile, TAB(49); USING "######,#.##"; TaxBill(1).TotalBillDue + TaxBill(1).PriorYrBalance
  PRINT #RptFile, TAB(51); USING "######"; TaxBill(1).CustRec
  PRINT #RptFile,
  PRINT #RptFile, TAB(11); LEFT$(CustName$, 21)
  PRINT #RptFile, TAB(11); LEFT$(TaxBill(1).RDesc1, 21)
  PRINT #RptFile, TAB(11); LEFT$(TaxBill(1).RDesc2, 21)
  'v line 12
  PRINT #RptFile,
  
  PRINT #RptFile, USING "########,#"; TaxBill(1).RealValue; TAB(15); TaxBill(1).PersValue;
  NetTaxVal# = Round#(TaxBill(1).RealValue + TaxBill(1).PersValue)
  NetTaxVal# = Round#(NetTaxVal# - TaxBill(1).ExptValue)
'  STOP
  PRINT #RptFile, TAB(25); USING "######,#"; TaxBill(1).ExptValue;
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(36); LEFT$(CustName$, 25)
  PRINT #RptFile, USING "########,#"; NetTaxVal#;
  PRINT #RptFile, USING "#.##"; TAB(16); TaxRate#;
  PRINT #RptFile, TAB(21); USING "#######,#.##"; Round#(TaxBill(1).TotalBillDue - TaxBill(1).LateTaxDue);
  PRINT #RptFile, TAB(36); LEFT$(QPTrim$(TaxBill(1).CustAdd1), 25)
  PRINT #RptFile, TAB(36); LEFT$(QPTrim$(TaxBill(1).CustAdd2), 25)
  PRINT #RptFile, TAB(36); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  PRINT #RptFile, TAB(21); USING "#######,#.##"; TaxBill(1).LateTaxDue
  PRINT #RptFile,
  PRINT #RptFile, TAB(2); USING "#####,#.##"; TaxBill(1).PriorYrBalance;
  PRINT #RptFile, TAB(21); USING "#######,#.##"; TaxBill(1).TotalBillDue + TaxBill(1).PriorYrBalance
'; TAB(47); TaxBill(1).TotalBillDue + TaxBill(1).PriorYrBalance
'  PRINT #RptFile, TAB(21); USING "#######,#.##"; TaxBill(1).TotalBillDue; TAB(47); TaxBill(1).TotalBillDue
  PRINT #RptFile,
'  PRINT #RptFile,
  PRINT #RptFile, "~"

  'PRINT #RptFile, TAB(25); USING "######,#"; TaxBill(1).ExptValue;
  'PRINT #RptFile, TAB(36); LEFT$(CustName$, 25)
  'PRINT #RptFile, TAB(36); LEFT$(QPTrim$(TaxBill(1).CustAdd1), 25)
  'PRINT #RptFile, TAB(36); LEFT$(QPTrim$(TaxBill(1).CustAdd2), 25)
  'PRINT #RptFile, USING "########,#"; NetTaxVal#;
  'PRINT #RptFile, USING "#.##"; TAB(16); TaxRate#;
  'PRINT #RptFile, TAB(21); USING "#######,#.##"; Round#(TaxBill(1).TotalBillDue - TaxBill(1).LateTaxDue);
  'PRINT #RptFile, TAB(36); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip

