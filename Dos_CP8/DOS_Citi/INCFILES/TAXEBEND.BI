  TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF
  
  '
  PRINT #RptFile, " BN"; USING "#####"; PrnCnt;
  PRINT #RptFile, ' TAB(74); "TOP"
  PRINT #RptFile,
  PRINT #RptFile, TAB(58); TaxBill(1).TaxYear;

  PRINT #RptFile, TAB(66); USING "######"; TaxBill(1).CustRec;
  PRINT #RptFile, TAB(74); USING "######"; TaxBill(1).BillNumber
  PRINT #RptFile,
  PRINT #RptFile,
  NetTaxVal# = Round#(TaxBill(1).RealValue + TaxBill(1).PersValue)
  

  'PRINT #RptFile, TAB(13); USING "########,#"; TaxBill(1).PersValue; TaxBill(1).RealValue; NetTaxVal#;
  PRINT #RptFile, TAB(22); USING "########,#"; TaxBill(1).PersValue; TAB(33); TaxBill(1).RealValue;
  PRINT #RptFile, TAB(45); USING "########,#"; NetTaxVal#;
  NetTaxVal# = Round#(NetTaxVal# - TaxBill(1).ExptValue)
  'PRINT #RptFile, TAB(33); USING "######.##"; (TaxBill(1).TotalBillDue - TaxBill(1).LateTaxDue);
  PRINT #RptFile, TAB(59); USING "######,#"; TaxBill(1).ExptValue;

  'PRINT #RptFile, TAB(54); USING "####.##"; TaxBill(1).LateTaxDue;
  'PRINT #RptFile, TAB(66); USING "#####.##"; TaxBill(1).TotalBillDue
  TwoPct# = Round#(TaxBill(1).TotalBillDue * .02)
  PRINT #RptFile, TAB(72); USING "######,#"; NetTaxVal#
  PRINT #RptFile,
  PRINT #RptFile,

  PRINT #RptFile, USING "#.##"; TAB(22); TaxRate#;
  PRINT #RptFile, TAB(30); USING "######.##"; TaxBill(1).PriorYrBalance; TAB(41); TaxBill(1).LateTaxDue; TAB(54); TaxBill(1).TotalBillDue
  PRINT #RptFile, 'USING "######.##"; TAB(54); TwoPct#

  PRINT #RptFile, TAB(22); PINTemp(1)
  'PRINT #RptFile, TAB(22); LEFT$(TaxBill(1).RDesc1, 21)
  PRINT #RptFile,
  PRINT #RptFile, TAB(54); USING "######.##"; TaxBill(1).TotalBillDue
  PRINT #RptFile, TAB(22); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(22); LEFT$(TaxBill(1).CustAdd1, 25)
  PRINT #RptFile, TAB(22); LEFT$(TaxBill(1).CustAdd2, 25)
  PRINT #RptFile, TAB(22); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, '"BOTTOM"; TAB(70); "BOTTOM"


