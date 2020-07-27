  IF BPrnCnt = 3 THEN
    PRINT #RptFile, CHR$(12);
    BPrnCnt = 0
  END IF
  IF BPrnCnt = 0 THEN
    PRINT #RptFile, CHR$(27); "&l3A"; CHR$(27); "&l6D";
  END IF
  
  TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF

  PRINT #RptFile, "TOP"
  PRINT #RptFile, TAB(12); TaxBill(1).TaxYear;
  PRINT #RptFile, USING "#.##"; TAB(22); TaxRate#;
  PRINT #RptFile, TAB(54); USING "######"; TaxBill(1).CustRec;
  PRINT #RptFile, TAB(64); USING "######"; TaxBill(1).BillNumber

  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  NetTaxVal# = Round#(TaxBill(1).RealValue + TaxBill(1).PersValue)
  PRINT #RptFile, USING "######"; TaxBill(1).PersValue;
  PRINT #RptFile, TAB(9); USING "########"; TaxBill(1).RealValue;
  PRINT #RptFile, TAB(20); USING "#######,#"; NetTaxVal#;
  NetTaxVal# = Round#(NetTaxVal# - TaxBill(1).ExptValue)
  PRINT #RptFile, TAB(31); USING "#####.##"; (TaxBill(1).TotalBillDue - TaxBill(1).LateTaxDue);
  PRINT #RptFile, TAB(41); USING "######"; TaxBill(1).ExptValue;
  IF TaxBill(1).LateTaxDue > 0 THEN
    PRINT #RptFile, TAB(66); USING "####.##"; TaxBill(1).LateTaxDue;
  END IF
  PRINT #RptFile, TAB(73); USING "#####.##"; TaxBill(1).TotalBillDue
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
  'PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, "~ BN"; USING "#####"; PrnCnt
  'PRINT #RptFile, "bottom!!"
'  IF PrnCnt MOD 3 = 0 THEN
'    PRINT #RptFile, CHR$(12)
'  END IF
   BPrnCnt = BPrnCnt + 1

