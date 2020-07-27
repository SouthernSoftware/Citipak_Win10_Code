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
  PRINT #RptFile, QPTrim$(STR$(WhatYear - 1)); " TAXES ADVERTISED ON: "; AdvDate$; " TO AVOID COST,"
  PRINT #RptFile, "TAXES MUST BE PAID IN FULL BY: "; PayDate$
  PRINT #RptFile,
  PRINT #RptFile, "~"


