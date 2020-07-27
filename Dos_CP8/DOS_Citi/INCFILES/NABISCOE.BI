  TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF

  PRINT #RptFile, CHR$(27); CHR$(48); "~"; TAB(50); USING "####"; PrnCnt

  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TaxBill(1).TaxYear;
  PRINT #RptFile, TAB(10); USING "######"; TaxBill(1).BillNumber;
  PRINT #RptFile, TAB(19); USING "#####"; TaxBill(1).CustRec;
  PRINT #RptFile, TAB(34); USING "####"; TaxBill(1).TaxYear;
  PRINT #RptFile, TAB(43); USING "##.##"; TaxRate#;
  PRINT #RptFile, TAB(48); USING "######.##"; TaxBill(1).TotalBillDue + TaxBill(1).PriorYrBalance
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(2); LEFT$(TaxBill(1).RDesc1, 21)
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, PINTemp(1); TAB(20); USING "########"; TaxBill(1).ExptValue
  PRINT #RptFile,
  PRINT #RptFile, TAB(48); USING "#####,#.##"; TaxBill(1).TotalBillDue + TaxBill(1).PriorYrBalance
  PRINT #RptFile,
  NetTaxVal# = Round#(TaxBill(1).RealValue + TaxBill(1).PersValue)
  PRINT #RptFile, USING "########"; TaxBill(1).RealValue; TAB(11); TaxBill(1).PersValue; TAB(22); NetTaxVal#
  NetTaxVal# = Round#(NetTaxVal# - TaxBill(1).ExptValue)
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  
  PRINT #RptFile, TAB(24); TaxBill(1).TaxYear;
  PRINT #RptFile, TAB(30); USING "######"; TaxBill(1).BillNumber;
  PRINT #RptFile, TAB(39); USING "#####"; TaxBill(1).CustRec;
  PRINT #RptFile, TAB(48); USING "#####,#.##"; TaxBill(1).TotalBillDue + TaxBill(1).PriorYrBalance
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(30); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(30); LEFT$(QPTrim$(TaxBill(1).CustAdd1), 25)
  PRINT #RptFile, TAB(30); LEFT$(QPTrim$(TaxBill(1).CustAdd2), 25)
  PRINT #RptFile, TAB(30); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, "~"

