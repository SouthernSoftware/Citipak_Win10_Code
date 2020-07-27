
  TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF

  PRINT #RptFile, "~"
  PRINT #RptFile, TAB(24); TaxBill(1).TaxYear;
  PRINT #RptFile, USING "#.##"; TAB(32); TaxRate#;
  PRINT #RptFile, TAB(64); USING "######"; TaxBill(1).CustRec;
  PRINT #RptFile, TAB(73); USING "######"; TaxBill(1).BillNumber
  PRINT #RptFile,
  PRINT #RptFile, TAB(67); QPTrim$(TaxBill(1).RealPin)
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  NetTaxVal# = round#(TaxBill(1).RealValue + TaxBill(1).PersValue)
  PRINT #RptFile, USING "########"; TaxBill(1).PersValue;
  PRINT #RptFile, TAB(10); USING "########"; TaxBill(1).RealValue;
  
  NetTaxVal# = round#(NetTaxVal# - TaxBill(1).ExptValue)

  PRINT #RptFile, TAB(20); USING "######"; TaxBill(1).ExptValue;
  PRINT #RptFile, TAB(30); USING "########"; NetTaxVal#;

  PRINT #RptFile, TAB(40); USING "#####.##"; (TaxBill(1).TotalBillDue - TaxBill(1).LateTaxDue);

  'PRINT #RptFile, TAB(60); USING "#####.##"; TaxBill(1).TotalBillDue;

  'IF TaxBill(1).LateTaxDue > 0 THEN
    PRINT #RptFile, TAB(50); USING "####.##"; TaxBill(1).LateTaxDue;
  'END IF

  PRINT #RptFile, TAB(72); USING "#####.##"; TaxBill(1).TotalBillDue
  PRINT #RptFile,
  PRINT #RptFile, 'TAB(2); RTRIM$(TaxBill(1).RealPIN); "   "; LEFT$(TaxBill(1).RDesc1, 21)
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(4); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(4); LEFT$(TaxBill(1).CustAdd1, 25)
  PRINT #RptFile, TAB(4); LEFT$(TaxBill(1).CustAdd2, 25)
  PRINT #RptFile, TAB(4); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  PRINT #RptFile, 'TAB(55); "Janice S. Thomas"
  PRINT #RptFile,
'for wagram'm bill only
'  PRINT #RptFile,
'******************
  PRINT #RptFile, "~ BN"; USING "#####"; PrnCnt
  'PRINT #RptFile, "bottom!!"


