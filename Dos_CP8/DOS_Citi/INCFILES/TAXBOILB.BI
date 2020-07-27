  TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF

  Map$ = QPTrim$(LEFT$(TaxBill(1).RDesc2, 6))
  Blk$ = QPTrim$(MID$(TaxBill(1).RDesc2, 7, 6))
  Lot$ = QPTrim$(MID$(TaxBill(1).RDesc2, 14, 6))

'  RSET PINTemp(1) = TaxBill(1).RealPIN

  PRINT #RptFile, "~"; TAB(30); TaxBill(1).TaxYear; USING "#.##"; TAB(41); TaxRate#;
  PRINT #RptFile, TAB(72); QPTrim$(PINTemp(1));
  PRINT #RptFile, TAB(80); USING "######"; TaxBill(1).BillNumber
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(68); USING "######"; TaxBill(1).CustRec
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  NetTaxVal# = round#(TaxBill(1).RealValue + TaxBill(1).PersValue)
  PRINT #RptFile, TAB(12); USING "######,#"; TaxBill(1).PersValue;
  PRINT #RptFile, TAB(22); USING "#######,#"; TaxBill(1).RealValue;
  NetTaxVal# = round#(NetTaxVal# - TaxBill(1).ExptValue)
  PRINT #RptFile, TAB(35); USING "#######,#"; NetTaxVal#;
  'PRINT #RptFile, TAB(39); USING "######.##"; (TaxBill(1).TotalBillDue - TaxBill(1).LateTaxDue);
  PRINT #RptFile, TAB(52); USING "######,#"; TaxBill(1).ExptValue;
  PRINT #RptFile, TAB(68); USING "###.##"; TaxBill(1).LateTaxDue;
  PRINT #RptFile, TAB(78); USING "#####.##"; TaxBill(1).TotalBillDue

  PRINT #RptFile,
  PRINT #RptFile, TAB(4); "M: "; Map$
  PRINT #RptFile, TAB(4); "B: "; Blk$
  PRINT #RptFile, TAB(4); "L: "; Lot$

'  PRINT #RptFile,
'  PRINT #RptFile, LEFT$(TaxBill(1).RDesc1, 21)
'  PRINT #RptFile, LEFT$(TaxBill(1).RDesc2, 21)

'  PRINT #RptFile, 'TAB(5); "IF TAXES ARE ESCROWED SEND BILL TO"
'  PRINT #RptFile, 'TAB(5); "MORTGAGE COMPANY."
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(6); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(6); LEFT$(TaxBill(1).CustAdd1, 25)
  PRINT #RptFile, TAB(6); LEFT$(TaxBill(1).CustAdd2, 25)
  PRINT #RptFile, TAB(6); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  PRINT #RptFile,
  PRINT #RptFile, TAB(4); TaxBill(1).RDesc1
  PRINT #RptFile,
  PRINT #RptFile, USING "######.##"; TaxBill(1).RealTaxDue; TaxBill(1).PersTaxDue
  PRINT #RptFile, "~"

