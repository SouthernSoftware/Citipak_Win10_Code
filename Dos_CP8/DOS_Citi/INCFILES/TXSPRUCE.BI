  TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF

  PRINT #RptFile, QPTrim$(STR$(TaxBill(1).TaxYear)); "-"; QPTrim$(STR$(TaxBill(1).BillNumber)); "~"; QPTrim$(CustName$); "~";
  PRINT #RptFile, QPTrim$(TaxBill(1).CustAdd1); "~"; QPTrim$(TaxBill(1).CustAdd2); "~";
  PRINT #RptFile, QPTrim$(TaxBill(1).CustAdd3); " "; QPTrim$(TaxBill(1).CustZip); "~";
  PRINT #RptFile, QPTrim$(TaxBill(1).RDesc3); "~"; PINTemp(1); "~";
  Desc$ = QPTrim$(TaxBill(1).RDesc1) + " " + QPTrim$(TaxBill(1).RDesc2)
  PRINT #RptFile, Desc$; "~";
  PRINT #RptFile, QPTrim$(STR$(TaxBill(1).RealValue)); "~";
  PRINT #RptFile, QPTrim$(STR$(TaxBill(1).PersValue)); "~";
  PRINT #RptFile, QPTrim$(STR$(TaxBill(1).ExptValue)); "~";

  NetTaxVal# = Round#(TaxBill(1).RealValue + TaxBill(1).PersValue)
  NetTaxVal# = Round#(NetTaxVal# - TaxBill(1).ExptValue)
  PRINT #RptFile, QPTrim$(STR$(NetTaxVal#)); "~";
  PRINT #RptFile, USING "$$##.##"; TaxRate#;
  PRINT #RptFile, "~";
  PRINT #RptFile, USING "$$######.##"; TaxBill(1).TotalBillDue

