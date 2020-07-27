  'TaxBillRec (1)
  TaxRate# = TaxBillRec(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBillRec(1).PersTaxRate
  END IF

  PRINT #RptFile, QPTrim$(STR$(TaxBillRec(1).TaxYear)); "~";
  PRINT #RptFile, QPTrim$(STR$(TaxBillRec(1).BillNumber)); "~";
  PRINT #RptFile, QPTrim$(CustName$); "~";
  PRINT #RptFile, QPTrim$(TaxBillRec(1).CustAdd1); "~";
  PRINT #RptFile, QPTrim$(TaxBillRec(1).CustAdd2); "~";
  PRINT #RptFile, QPTrim$(TaxBillRec(1).CustAdd3); "~";
  PRINT #RptFile, QPTrim$(TaxBillRec(1).CustZip); "~";
  PRINT #RptFile, QPTrim$(TaxBillRec(1).RDesc3); "~";
  PRINT #RptFile, QPTrim$(PINTemp(1)); "~";
  Desc$ = QPTrim$(TaxBillRec(1).RDesc1) + " " + QPTrim$(TaxBillRec(1).RDesc2)
  PRINT #RptFile, Desc$; "~";
  PRINT #RptFile, QPTrim$(STR$(TaxBillRec(1).RealValue)); "~";
  PRINT #RptFile, QPTrim$(STR$(TaxBillRec(1).ExptValue)); "~";
  NetTaxVal# = Round#(TaxBillRec(1).RealValue + TaxBillRec(1).PersValue)
  NetTaxVal# = Round#(NetTaxVal# - TaxBillRec(1).ExptValue)
  PRINT #RptFile, QPTrim$(STR$(NetTaxVal#)); "~";
  PRINT #RptFile, QPTrim$(STR$(TaxRate#)); "~";
  PRINT #RptFile, QPTrim$(FUSING$(STR$(TaxBillRec(1).TotalBillDue), "######.##"))

