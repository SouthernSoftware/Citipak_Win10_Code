

  CLOSE TaxFile
  TaxFile = FREEFILE
  OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, TaxBillRec(1).CustRec, TaxCust(1)

  CustName$ = RTRIM$(TaxCust(1).FNAME) + " " + LTRIM$(TaxCust(1).LName)
  CustName$ = LTRIM$(CustName$)

 'Must Calc Late Fee Here
  PenAmt# = Round#((TaxBillRec(1).TotalBillDue) * .05)
  '+4
  PRINT #RptFile, TAB(7); "~"; TAB(73); "BN"; USING "#####"; PrnCnt
  PRINT #RptFile, ""
  PRINT #RptFile, 'TAB(78); USING "#####"; TaxBillRec(1).BillNumber
  PRINT #RptFile, TAB(78); USING "#####"; TaxBillRec(1).BillNumber
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(5); USING "######,#"; TaxBillRec(1).RealValue - TaxBillRec(1).PersValue;
  PRINT #RptFile, TAB(17); USING "######,#"; TaxBillRec(1).PersValue;
  PRINT #RptFile, TAB(28); USING "######,#"; TaxBillRec(1).RealValue;
  PRINT #RptFile, TAB(38); USING "#####,#.##"; TaxBillRec(1).TotalBillDue;
  PRINT #RptFile, TAB(64); USING "#####,#.##"; TaxBillRec(1).TotalBillDue;
  PenAmt# = Round#(TaxBillRec(1).TotalBillDue * .05)
  PRINT #RptFile, TAB(78); USING "####.##"; PenAmt#
  
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(8); QPTrim$(TaxBillRec(1).RDesc1)
  PRINT #RptFile, TAB(8); QPTrim$(TaxBillRec(1).RDesc2)
  PRINT #RptFile,
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(15); "ACCT # "; TaxBillRec(1).CustRec
  PRINT #RptFile,
  PRINT #RptFile, TAB(15); LEFT$(CustName$, 45)
  PRINT #RptFile, TAB(15); LEFT$(TaxBillRec(1).CustAdd1, 35)
  PRINT #RptFile, TAB(15); LEFT$(TaxBillRec(1).CustAdd2, 35)
  PRINT #RptFile, TAB(15); QPTrim$(TaxBillRec(1).CustAdd3); " "; TaxBillRec(1).CustZip
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(7); "~"


