  BYear = (VAL(RIGHT$(DATE$, 4)) + 1)
  DueDate$ = "01-05-" + QPTrim$(STR$(BYear))
  TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF
   
  Desc$ = QPTrim$(TaxBill(1).RDesc1) + " " + QPTrim$(TaxBill(1).RDesc2)
  PRINT #RptFile, CHR$(27); CHR$(58); CHR$(27); CHR$(48);
 ' PRINT #RptFile, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
  PRINT #RptFile, "~"; TAB(5); USING "####"; PrnCnt
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(34); TaxBill(1).TaxYear; TAB(46); USING "#####"; TaxBill(1).BillNumber; TAB(63); TaxBill(1).CustRec
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, USING "####"; TaxBill(1).TaxYear;
  PRINT #RptFile, TAB(6); PINTemp(1);
  PRINT #RptFile, TAB(24); USING "#####"; TaxBill(1).CustRec
  PRINT #RptFile, 'TAB(10); LEFT$(Desc$, 20)
  PRINT #RptFile,
  PRINT #RptFile,
  NetTaxVal# = Round#(TaxBill(1).REALVALUE + TaxBill(1).PERSVALUE)
  PRINT #RptFile, TAB(39); USING "Total Due:#######,#.##"; TaxBill(1).TotalBillDue + TaxBill(1).PriorYrBalance
  PRINT #RptFile,
  PRINT #RptFile, USING "#######,#"; TaxBill(1).REALVALUE; TaxBill(1).PERSVALUE; TaxBill(1).ExptValue
  PRINT #RptFile,
  NetTaxVal# = Round#(NetTaxVal# - TaxBill(1).ExptValue)
  PRINT #RptFile, TAB(39); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(39); LEFT$(TaxBill(1).CustAdd1, 25)
  PRINT #RptFile, TAB(39); LEFT$(TaxBill(1).CustAdd2, 25)
  PRINT #RptFile, TAB(39); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  PRINT #RptFile,
  PRINT #RptFile, USING "#####,#.##"; TaxBill(1).LateTaxDue;
  PRINT #RptFile, USING "#.##"; TAB(13); TaxRate#;
  PRINT #RptFile, TAB(20); USING "#####.##"; Round#(TaxBill(1).TotalBillDue + TaxBill(1).PriorYrBalance)
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, "~"

