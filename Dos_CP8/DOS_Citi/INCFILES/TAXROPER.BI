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
'  PRINT #RptFile,
'  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, USING "####"; TaxBill(1).TaxYear;
  PRINT #RptFile, TAB(9); USING "######"; TaxBill(1).BillNumber
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, USING "#######"; TaxBill(1).CustRec; TAB(40); TaxBill(1).TaxYear; TAB(50); TaxBill(1).BillNumber; TAB(60); TaxBill(1).CustRec
  PRINT #RptFile,
  PRINT #RptFile, USING "########,#"; TaxBill(1).REALVALUE; TAB(13); TaxBill(1).PERSVALUE;
  NetTaxVal# = Round#(TaxBill(1).REALVALUE + TaxBill(1).PERSVALUE)
  PRINT #RptFile, TAB(25); USING "########,#"; NetTaxVal#;
  PRINT #RptFile, TAB(56); USING "$$#####,#.##"; TotalTax#
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(26); USING "#######.##"; TaxBill(1).PriorYrBalance
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(26); USING "#######.##"; TaxBill(1).ExptValue
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(43); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(43); LEFT$(TaxBill(1).CustAdd1, 25)
  PRINT #RptFile, TAB(26); USING "#######.##"; TaxBill(1).LateTaxDue;
  PRINT #RptFile, TAB(43); LEFT$(TaxBill(1).CustAdd2, 25)
  PRINT #RptFile, TAB(43); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  'PRINT #RptFile, USING "#.##"; TAB(21); TaxRate#;
  'PRINT #RptFile, TAB(28); USING "####,#.##"; Round#(TaxBill(1).TotalBillDue - TaxBill(1).LateTaxDue)
  NetTaxVal# = Round#(NetTaxVal# - TaxBill(1).ExptValue)
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(22); USING "$$#####,#.##"; TotalTax#
  PRINT #RptFile,
  PRINT #RptFile, "  Bill Date: "; DATE$
  PRINT #RptFile, "       Rate: "; USING "#.##"; TaxRate#
  PRINT #RptFile,
  PRINT #RptFile, "~"

  'PRINT #RptFile, USING "########,#"; TaxBill(1).ExptValue;
  '
  'PRINT #RptFile, TAB(21); USING "#######,#.##"; Round#(TaxBill(1).TotalBillDue - TaxBill(1).LateTaxDue);
  'PRINT #RptFile,
  'PRINT #RptFile, TAB(21); USING "#######,#.##"; TaxBill(1).LateTaxDue
  'PRINT #RptFile,
  'PRINT #RptFile, TAB(2); USING "#####,#.##"; TaxBill(1).PriorYrBalance;
  'PRINT #RptFile, TAB(21); USING "#######,#.##"; TaxBill(1).TotalBillDue + TaxBill(1).PriorYrBalance; TAB(47); TaxBill(1).TotalBillDue + TaxBill(1).PriorYrBalance
  'PRINT #RptFile,
  'PRINT #RptFile,

  'PRINT #RptFile, 'TAB(24); PINTemp(1); TAB(36); USING "####"; TaxBill(1).TaxYear;
  'PRINT #RptFile, 'TAB(41); USING "########"; TaxBill(1).BillNumber;
  'PRINT #RptFile, 'TAB(16)
  'PRINT #RptFile,

'    FOR ll = 1 TO 7: LPRINT " ": NEXT ll
'    LPRINT TAB(2); year$; TAB(7);
'    IF count = 0 THEN LPRINT "REPRINT";  ELSE LPRINT USING "######"; count;
'    LPRINT TAB(26); pin$
'    LPRINT : LPRINT
'    LPRINT TAB(2); ACCT$; TAB(14); "SMITH CREEK"
'    LPRINT
'    LPRINT
'    LPRINT TAB(2); LEFT$(PDESC$, 35); TAB(42); year$; TAB(48);
'    IF count = 0 THEN
'       LPRINT "REPRINT";
'     ELSE
'       LPRINT USING "######"; count;
'     END IF
'    LPRINT TAB(61); ACCT$
'    LPRINT
'    LPRINT
'    LPRINT TAB(56); USING "$$####,#.##"; TAXOWED!
'    LPRINT TAB(2); USING "#######,#"; REALVALUE#;
'    LPRINT TAB(14); USING "#######,#"; PERSVALUE#;
'    LPRINT TAB(26); USING "#######,#"; REALVALUE# + PERSVALUE#
'    LPRINT
'    LPRINT TAB(28); USING "####,#.##"; TAGTOTAL#
'    LPRINT
'    LPRINT
'    LPRINT
'    LPRINT
'    LPRINT TAB(22); USING ".##"; grate;
'    LPRINT TAB(28); USING "####,#.##"; TAXOWED! - TAGTOTAL#
'    LPRINT TAB(42); N$
'    LPRINT TAB(28); USING "####,#.##"; latetax!;
'    LPRINT TAB(42); A1$
'    LPRINT TAB(42); A2$
'    LPRINT TAB(42); CY$; " "; st$; " "; zp$
'    LPRINT TAB(4); DUEDATE$; TAB(22); USING "$$#####,#.##"; TAXOWED!
'    LPRINT
'    LPRINT
'    LPRINT
'    LPRINT

