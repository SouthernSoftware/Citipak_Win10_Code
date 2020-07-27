  TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF
  
  PRINT #RptFile, CHR$(27); CHR$(58); "~";
  PRINT #RptFile, TAB(35); TaxBill(1).TaxYear;
  PRINT #RptFile, USING "#.##"; TAB(46); TaxRate#;
  PRINT #RptFile, TAB(78); USING "######"; TaxBill(1).CustRec;
  PRINT #RptFile, TAB(90); USING "######"; TaxBill(1).BillNumber
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  NetTaxVal# = Round#(TaxBill(1).RealValue + TaxBill(1).PersValue)
  PRINT #RptFile, USING "######,#"; TaxBill(1).PersValue;

  PRINT #RptFile, TAB(13); USING "#######,#"; TaxBill(1).RealValue;

  PRINT #RptFile, TAB(26); USING "#######,#"; NetTaxVal#;
  NetTaxVal# = Round#(NetTaxVal# - TaxBill(1).ExptValue)
  PRINT #RptFile, TAB(37); USING "######.##"; (TaxBill(1).TotalBillDue - TaxBill(1).LateTaxDue);
  PRINT #RptFile, TAB(51); USING "######,#"; TaxBill(1).ExptValue;
  PRINT #RptFile, TAB(81); USING "####.##"; TaxBill(1).LateTaxDue;
  PRINT #RptFile, TAB(89); USING "#####.##"; TaxBill(1).TotalBillDue
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(8); "IF TAXES ARE ESCROWED SEND BILL TO"
  PRINT #RptFile, TAB(8); "MORTGAGE COMPANY."
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(8); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(8); LEFT$(TaxBill(1).CustAdd1, 25)
  PRINT #RptFile, TAB(8); LEFT$(TaxBill(1).CustAdd2, 25)
  PRINT #RptFile, TAB(8); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, "~"
'  LPCnt = LPCnt + 1
'  IF LPCnt > 1 THEN
'    PRINT #RptFile, CHR$(12)
'    LPCnt = 0
'  END IF


