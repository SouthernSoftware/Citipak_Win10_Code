  BYear = (VAL(RIGHT$(DATE$, 4)) + 1)
  DueDate$ = "01-05-" + QPTrim$(STR$(BYear))
  TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF
  TotalTax# = Round(TaxBill(1).TotalBillDue + TaxBill(1).PriorYrBalance)
  Desc$ = QPTrim$(TaxBill(1).RDesc1) + " " + QPTrim$(TaxBill(1).RDesc2)

  PRINT #RptFile, CHR$(27); CHR$(58); CHR$(27); CHR$(48);
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
  PRINT #RptFile,
  PRINT #RptFile, USING "####"; TaxBill(1).TaxYear;
  PRINT #RptFile, USING "######"; TaxBill(1).BillNumber;
  PRINT #RptFile, TAB(14); USING "######"; TaxBill(1).CustRec;
  PRINT #RptFile, TAB(25); TaxBill(1).RealPin;
  NetTaxVal# = Round#(TaxBill(1).REALVALUE + TaxBill(1).PERSVALUE)

  PRINT #RptFile, TAB(41); TaxBill(1).TaxYear; TAB(50); TaxBill(1).CustRec; TAB(57); USING "$$#####,#.##"; TotalTax#
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(12); Desc$
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,

  PRINT #RptFile, USING "########,#"; TaxBill(1).REALVALUE; TAB(13); TaxBill(1).PERSVALUE;
  PRINT #RptFile, TAB(25); USING "########,#"; TaxBill(1).ExptValue
  'PRINT #RptFile, TAB(56); USING "$$#####,#.##"; TotalTax#
  PRINT #RptFile, TAB(47); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(47); LEFT$(TaxBill(1).CustAdd1, 25)
  PRINT #RptFile, TAB(47); LEFT$(TaxBill(1).CustAdd2, 25)
  PRINT #RptFile, USING "########,#"; NetTaxVal#;
  PRINT #RptFile, TAB(20); USING "#.##"; TaxRate#;
  PRINT #RptFile, TAB(24); USING "$$#####,#.##"; TotalTax#;

  PRINT #RptFile, TAB(47); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(2); USING "#######.##"; TaxBill(1).PriorYrBalance; TAB(26); TotalTax#
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, '"  Bill Date: "; DATE$
  'PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, "~"

  'NetTaxVal# = Round#(NetTaxVal# - TaxBill(1).ExptValue)

