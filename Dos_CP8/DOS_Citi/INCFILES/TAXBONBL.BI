    TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF

  
  NetTaxVal# = Round#(TaxBill(1).RealValue + TaxBill(1).PersValue)
  NetTaxVal# = Round#(NetTaxVal# - TaxBill(1).ExptValue)


  PRINT #RptFile, "~"; TAB(78); "~"
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(50); TaxBill(1).TaxYear; TAB(59); USING "######"; TaxBill(1).BillNumber;
  PRINT #RptFile, TAB(69); TaxBill(1).RealPin
  PRINT #RptFile,
  PRINT #RptFile, TAB(50); USING "######"; TaxBill(1).CustRec
  'PRINT #RptFile, TAB(63); TaxBill(1).TownShip
  PRINT #RptFile,
  PRINT #RptFile, TAB(51); QPTrim$(LEFT$(TaxBill(1).RDesc1, 21))
  PRINT #RptFile,
  PRINT #RptFile, TAB(51); USING "#######,#"; TaxBill(1).RealValue; TAB(61); TaxBill(1).PersValue; TAB(71); NetTaxVal#
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); LEFT$(CustName$, 25);
  PRINT #RptFile, TAB(66); USING "#.##"; TaxRate#;
  PRINT #RptFile, TAB(71); USING "#####.##"; Round#(TaxBill(1).TotalBillDue - TaxBill(1).LateTaxDue)
  PRINT #RptFile, TAB(5); LEFT$(TaxBill(1).CustAdd1, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxBill(1).CustAdd2, 25);
  PRINT #RptFile, TAB(71); USING "####.##"; TaxBill(1).LateTaxDue
  PRINT #RptFile, TAB(5); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  PRINT #RptFile, TAB(51); "09-01-"; RIGHT$(DATE$, 4); TAB(71); USING "#####.##"; TaxBill(1).TotalBillDue
  'PRINT #RptFile, TAB(71); USING "######.##";;
  'PRINT #RptFile, TAB(44); USING "######,#"; TaxBill(1).ExptValue;
  'PRINT #RptFile,
  'PRINT #RptFile, " BN"; USING "#####"; PrnCnt
  PRINT #RptFile,
  'PRINT #RptFile,
  PRINT #RptFile, "~"; TAB(78); "~"


