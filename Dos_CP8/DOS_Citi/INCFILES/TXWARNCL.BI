  GOSUB GetRealInfo
  GOSUB GetPersInfo

  BillNum$ = ParseBillNum$(TaxTransRec.Description)
  BillNum$ = LEFT$(BillNum$, 9)
  BilNumVal& = VAL(BillNum$)

'  BYear = (VAL(RIGHT$(DATE$, 4)) + 1)
'  DueDate$ = "01-05-" + QPTrim$(STR$(BYear))
'  TaxRate# = TaxBill(1).RealTaxRate
'  RealCalcVal#; PersCalcVal#
'  IF TaxRate# = 0 THEN
'    TaxRate# = TaxBill(1).PersTaxRate
'  END IF

  'TotalTax# = Round(TaxBill(1).TotalBillDue + TaxBill(1).PriorYrBalance)
  
  'Desc$ = QPTrim$(TaxBill(1).RDesc1) + " " + QPTrim$(TaxBill(1).RDesc2)

  PRINT #RptFile, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
  PRINT #RptFile, "~"; TAB(5); USING "####"; PrnCnt
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, USING "####"; TaxTransRec.TaxYear;
  PRINT #RptFile, TAB(9); BillNum$
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, USING "#######"; CustAcct&
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, USING "#######"; TAB(38); TaxTransRec.TaxYear; TAB(48); BilNumVal&; TAB(58); CustAcct&
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, USING "########,#"; RealCalcVal#; TAB(13); PersCalcVal#;
  NetTaxVal# = Round#(RealCalcVal# + PersCalcVal#)
  PRINT #RptFile, TAB(25); USING "########,#"; NetTaxVal#;
  PRINT #RptFile, TAB(56); USING "$$#####,#.##"; Balance#
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(47); LEFT$(CustName$, 25)
  PRINT #RptFile, " "; PayDate$; TAB(24); USING "$$#####,#.##"; Balance#;
  PRINT #RptFile, TAB(47); LEFT$(TaxCustRec.Addr1, 25)
  PRINT #RptFile, TAB(47); LEFT$(TaxCustRec.Addr2, 25)
  PRINT #RptFile, TAB(47); QPTrim$(TaxCustRec.City) + " "; TaxCustRec.State + " " + LEFT$(TaxCustRec.Zip, 5) + "-" + MID$(TaxCustRec.Zip, 6, 4)
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, "~"


