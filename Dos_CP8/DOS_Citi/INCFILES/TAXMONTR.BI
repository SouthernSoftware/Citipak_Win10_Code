

  CLOSE TaxFile
  TaxFile = FREEFILE
  OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, TaxBillRec(1).CustRec, TaxCust(1)

  CustName$ = QPTrim$(QPTrim$(TaxCust(1).FNAME) + " " + QPTrim$(TaxCust(1).LName))
  PenAmt# = Round#(TaxBillRec(1).TotalBillDue * .05)
  BilNum& = (BillInfo(1).BILLNUM + PrnCnt) - 1
 'Must Calc Late Fee Here

  PRINT #RptFile, "~"  '"TOP"; TAB(70); "TOP"
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(42); USING "##.##"; BillInfo(1).REALRATE
  PRINT #RptFile,
  PRINT #RptFile, TAB(7); LEFT$(QPTrim$(TaxBillRec(1).RDesc1), 20); TAB(32); USING "######,#"; TaxBillRec(1).RealValue - TaxBillRec(1).PersValue;
  PRINT #RptFile, TAB(41); USING "#####,#.##"; TaxBillRec(1).TotalBillDue
  PRINT #RptFile, TAB(7); QPTrim$(TaxBillRec(1).RDesc2)
  PRINT #RptFile, ""
  PRINT #RptFile,
  PRINT #RptFile, TAB(53); Num2Date$(BillInfo(1).DueDate); TAB(64); BilNum&;
  PRINT #RptFile, TAB(70); TaxBillRec(1).RealPin
  PRINT #RptFile,
  PRINT #RptFile, TAB(67); USING "#####,#.##"; TaxBillRec(1).TotalBillDue
  PRINT #RptFile, TAB(41); USING "#####,#.##"; TaxBillRec(1).TotalBillDue
  PRINT #RptFile, ""
  PRINT #RptFile,
  PRINT #RptFile, TAB(30); TaxBillRec(1).CustRec
  PRINT #RptFile, TAB(18); LEFT$(CustName$, 45)
  PRINT #RptFile, TAB(18); LEFT$(TaxCust(1).Addr1, 35)
  PRINT #RptFile, TAB(18); LEFT$(TaxCust(1).Addr2, 35)
  PRINT #RptFile, TAB(18); QPTrim$(TaxBillRec(1).CustAdd3); " "; TaxBillRec(1).CustZip; TAB(53); "Acct:"; TaxBillRec(1).CustRec
  'PRINT #RptFile,
  PRINT #RptFile, TAB(53); "Cust: "; LEFT$(CustName$, 30)
  PRINT #RptFile,
  PRINT #RptFile, "~"   '"BOTTOM"; TAB(70); "BOTTOM"


  'PRINT #RptFile, TAB(5); USING "######,#"; TaxBillRec(1).RealValue - TaxBillRec(1).PersValue;
  'PRINT #RptFile, TAB(15); USING "######,#"; TaxBillRec(1).PersValue;
  'PRINT #RptFile, TAB(25); USING "######,#"; TaxBillRec(1).RealValue;
  'PRINT #RptFile, TAB(35); USING "#####,#.##"; TaxBillRec(1).TotalBillDue;
  'PRINT #RptFile, TAB(57); USING "#####,#.##"; TaxBillRec(1).TotalBillDue;

'     LPRINT
'     LPRINT
'     LPRINT
'     LPRINT TAB(42); USING "##.##"; taxrate
'     LPRINT
'     LPRINT
'     LPRINT TAB(7); LEFT$(TBRECust(1).PropDesc1, 20); TAB(32); USING "######,#"; (TBRECust(1).LandValue + TBRECust(1).BldgValue) - TBRECust(1).LandUseValue;
'     LPRINT TAB(41); USING "#####,#.##"; TBRECust(1).LBCurYrTaxAmount
'     LPRINT TAB(7); LEFT$(TBRECust(1).PropDesc2, 20)
'     LPRINT TAB(7); LEFT$(TBRECust(1).PropDesc3, 20)
'     LPRINT TAB(7); LEFT$(TBRECust(1).PropDesc4, 20); TAB(52); DueDate$; TAB(63); TBRECust(1).LBCurYrBillNumber;
'     LPRINT TAB(72); LEFT$(TBRECust(1).MapNumber, 4); "-"; MID$(TBRECust(1).MapNumber, 5, 2); "-"; MID$(TBRECust(1).MapNumber, 7, 2); "-"; MID$(TBRECust(1).MapNumber, 9, 4)
'     LPRINT
'     LPRINT TAB(67); USING "#####,#.##"; TBRECust(1).LBCurYrTaxAmount
'     LPRINT TAB(41); USING "#####,#.##"; TBRECust(1).LBCurYrTaxAmount
'     LPRINT
'     LPRINT
'     LPRINT TAB(30); AcctNumber!
'     LPRINT TAB(18); TBRECust(1).CustName
'     LPRINT TAB(18); TBRECust(1).Address1
'     LPRINT TAB(18); TBRECust(1).Address2
'     LPRINT TAB(18); RTRIM$(TBRECust(1).Address3); " "; TBRECust(1).ZipCode
'     LPRINT
'     LPRINT TAB(52); "ACCT # "; AcctNumber!
'     LPRINT TAB(52); "CUST : "; TBRECust(1).CustName
'     LPRINT
'

