  PRINT #RptFile, "~"
  PRINT #RptFile, USING "#####"; PrnCnt
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(16); USING "######"; TaxBill(1).BillNumber;
  IF TaxBill(1).RealValue < -1000 THEN
    TaxBill(1).RealValue = 0
  END IF
  IF TaxBill(1).PersValue < -1000 THEN
    TaxBill(1).PersValue = 0
  END IF
  PRINT #RptFile, TAB(48); USING "########,#"; TaxBill(1).RealValue
  PRINT #RptFile, TAB(16); TaxBill(1).TaxYear;
  PRINT #RptFile, TAB(48); USING "########,#"; TaxBill(1).PersValue
  IF TaxBill(1).RealTaxRate# = 0 THEN
    TaxBill(1).RealTaxRate# = TaxBill(1).PersTaxRate#
  END IF
  PRINT #RptFile, TAB(16); USING "#.###"; TaxBill(1).RealTaxRate#;
  NetTaxVal# = round#(TaxBill(1).RealValue + TaxBill(1).PersValue)
  PRINT #RptFile, TAB(48); USING "########,#"; NetTaxVal#;
  PRINT #RptFile, TAB(81); USING "$$######,#.##"; TaxBill(1).TotalBillDue - TaxBill(1).LateTaxDue
  PRINT #RptFile, TAB(20); USING "######"; TaxBill(1).CustRec;
  PRINT #RptFile, TAB(55); USING "######,#"; TaxBill(1).ExptValue;
  PRINT #RptFile, TAB(81); USING "$$######,#.##"; TaxBill(1).LateTaxDue
  PRINT #RptFile, TAB(14); TaxBill(1).TownShip;
  NetTaxVal# = round#(NetTaxVal# - TaxBill(1).ExptValue)
  PRINT #RptFile, TAB(48); USING "########,#"; NetTaxVal#
  PRINT #RptFile, TAB(18); TaxBill(1).LotOrAcre;
  'PRINT #RptFile, "HERE"
'  IF TaxBill(1).LASize < 0 THEN
'    TaxBill(1).LASize = 0
'  END IF
  PRINT #RptFile, " "; TaxBill(1).LASize
  PRINT #RptFile, TAB(18); TaxBill(1).RealPin;
  PRINT #RptFile, TAB(81); USING "$$#######,#.##"; TaxBill(1).TotalBillDue
  PRINT #RptFile, TAB(18); TaxBill(1).RDesc1
  PRINT #RptFile, TAB(18); TaxBill(1).RDesc2
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(8); LEFT$(CustName$, 35)
  PRINT #RptFile, TAB(8); LEFT$(TaxBill(1).CustAdd1, 35)
  PRINT #RptFile, TAB(8); LEFT$(TaxBill(1).CustAdd2, 35)
  PRINT #RptFile, TAB(8); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  
  IF BillNum = 0 THEN
    BillNum = 1
   ELSE
    BillNum = 0
    PRINT #RptFile, ""
  END IF

