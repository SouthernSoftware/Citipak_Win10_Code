  GOSUB GetRealInfo
  GOSUB GetPersInfo

  BillNum$ = ParseBillNum$(TaxTransRec.Description)
  BillNum$ = LEFT$(BillNum$, 9)

  CalcVal# = Round#(RealCalcVal# + PersCalcVal#)

'  PRINT #RptFile, CHR$(27); CHR$(58); CHR$(27); CHR$(48);
  PRINT #RptFile, CHR$(27); CHR$(48); CHR$(27); CHR$(77);

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
  PRINT #RptFile, TAB(34); TaxTransRec.TaxYear; TAB(46); BillNum$; TAB(63); USING "#####"; CustAcct&
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, USING "######"; TaxTransRec.TaxYear;
  PRINT #RptFile, TAB(14); PINTemp(1);
  PRINT #RptFile, TAB(26); USING "######"; CustAcct&
  PRINT #RptFile, 'TAB(10); LEFT$(Desc$, 20)
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(39); USING "Total Due:#######,#.##"; Balance#
  PRINT #RptFile,
  PRINT #RptFile, USING "########,#"; RealCalcVal#; PersCalcVal#
  PRINT #RptFile,
  PRINT #RptFile, TAB(39); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(39); LEFT$(TaxCustRec.Addr1, 25)
  PRINT #RptFile, TAB(39); LEFT$(TaxCustRec.Addr2, 25)
  PRINT #RptFile, TAB(39); QPTrim$(TaxCustRec.City) + " "; TaxCustRec.State + " " + LEFT$(TaxCustRec.Zip, 5) + "-" + MID$(TaxCustRec.Zip, 6, 4)
  PRINT #RptFile,
  'PRINT #RptFile, USING "#####,#.##"; TaxBill(1).LateTaxDue;
  'PRINT #RptFile, USING "#.##"; TAB(14); TaxRate#;
  PRINT #RptFile, TAB(22); USING "####,#.##"; Balance#
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  'PRINT #RptFile, "  DELINQUENT-REMIT UPON RECEIPT"
  PRINT #RptFile, STR$(WhatYear); " TAXES ADVERTISED ON: "; AdvDate$; " TO AVOID COST,"
  PRINT #RptFile, " TAXES MUST BE PAID IN FULL BY: "; PayDate$
  PRINT #RptFile, "~"


'  GOSUB GetRealInfo
'  GOSUB GetPersInfo
'
'  BillNum$ = ParseBillNum$(TaxTransRec.Description)
'  BillNum$ = LEFT$(BillNum$, 9)
'
'  CalcVal# = Round#(RealCalcVal# + PersCalcVal#)
'
'  TaxRate# = TaxBill(1).RealTaxRate
'  IF TaxRate# = 0 THEN
'    TaxRate# = TaxBill(1).PersTaxRate
'  END IF
'  NetTaxVal# = Round#(TaxBill(1).RealValue + TaxBill(1).PersValue)
'
'  PRINT #RptFile, "~"; TAB(5); USING "####"; PrnCnt
'  PRINT #RptFile,
'  PRINT #RptFile,
'  PRINT #RptFile, USING "######"; TaxTransRec.TaxYear;
'  'PRINT #RptFile, TAB(50); TaxBill(1).TaxYear;
'
'  PRINT #RptFile, TAB(56); BillNum$;
'  'PRINT #RptFile, TAB(58); USING "######"; TaxBill(1).BillNumber;
'
'  PRINT #RptFile, TAB(69); PinTemp(1)
'  PRINT #RptFile,
'
'  PRINT #RptFile, TAB(50); USING "######"; CustAcct&
'
'  PRINT #RptFile,
'  PRINT #RptFile,
'  PRINT #RptFile,
'  PRINT #RptFile, TAB(50); USING "########"; RealCalcVal#;
'  PRINT #RptFile, TAB(59); USING "######"; PersCalcVal#;
'  PRINT #RptFile, TAB(69); USING "#########,"; CalcVal#
'  PRINT #RptFile,
'  PRINT #RptFile, TAB(69); USING "########"; PersExmp#
'  'PRINT #RptFile, TAB(69); USING "########"; TaxBill(1).ExptValue
'  PRINT #RptFile,
'
''  PRINT #RptFile, TAB(6); LEFT$(CustName$, 25)
''  PRINT #RptFile, TAB(6); LEFT$(TaxBill(1).CustAdd1, 25); TAB(68); USING "#.##"; TaxRate#
''  PRINT #RptFile, TAB(6); LEFT$(TaxBill(1).CustAdd2, 25)
''  PRINT #RptFile, TAB(6); QPTrim$(TaxBill(1).CustAdd3); " "; QPTrim$(TaxBill(1).CustZip)
''; TAB(69); USING "#####.##"; TaxBill(1).LateTaxDue

'  PRINT #RptFile, TAB(36); LEFT$(CustName$, 25)
'  PRINT #RptFile, TAB(36); LEFT$(TaxCustRec.Addr1, 25)
'  PRINT #RptFile, TAB(36); LEFT$(TaxCustRec.Addr2, 25)
'  PRINT #RptFile, TAB(36); QPTrim$(TaxCustRec.City) + " "; TaxCustRec.State + " " + LEFT$(TaxCustRec.Zip, 5) + "-" + MID$(TaxCustRec.Zip, 6, 4)
'  PRINT #RptFile,
'  PRINT #RptFile, TAB(20); USING "#######,#.##"; Balance#
'  'PRINT #RptFile, TAB(44); LEFT$(TaxBill(1).RDesc1, 21); TAB(69); USING "#####.##"; TaxBill(1).TotalBillDue
'  PRINT #RptFile,
'  PRINT #RptFile, "~ BN"; USING "#####"; PrnCnt


