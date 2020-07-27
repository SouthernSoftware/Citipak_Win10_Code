
  IF MortFlag AND NotDone THEN
    Cust$ = LEFT$(CustName$, 25)
    Addr1$ = LEFT$(TaxBill(1).CustAdd1, 25)
    Addr2$ = LEFT$(TaxBill(1).CustAdd2, 25)
    Addr3$ = QPTrim$(TaxBill(1).CustAdd3) + " " + TaxBill(1).CustZip
  ELSEIF MortFlag THEN
    Cust$ = LEFT$(QPTrim$(MortCodeRec(1).BName), 25)
    Addr1$ = LEFT$(QPTrim$(MortCodeRec(1).Add1), 25)
    Addr2$ = LEFT$(QPTrim$(MortCodeRec(1).Add2), 25)
    Addr3$ = QPTrim$(MortCodeRec(1).Add3)
  ELSE
    Cust$ = LEFT$(CustName$, 25)
    Addr1$ = LEFT$(TaxBill(1).CustAdd1, 25)
    Addr2$ = LEFT$(TaxBill(1).CustAdd2, 25)
    Addr3$ = QPTrim$(TaxBill(1).CustAdd3) + " " + TaxBill(1).CustZip
  END IF

  PRINT #RptFile, "~"; TAB(50); USING "####"; PrnCnt
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TaxBill(1).TaxYear;
  PRINT #RptFile, TAB(7); USING "######"; TaxBill(1).BillNumber;
  PRINT #RptFile, TAB(16); USING "#####"; TaxBill(1).CustRec;
  PRINT #RptFile, TAB(24); PINTemp(1); TAB(36); USING "####"; TaxBill(1).TaxYear;
  PRINT #RptFile, TAB(41); USING "########"; TaxBill(1).BillNumber;
  PRINT #RptFile, TAB(51); USING "######"; TaxBill(1).CustRec
  PRINT #RptFile, TAB(11); LEFT$(CustName$, 21)

  PRINT #RptFile, TAB(11); LEFT$(TaxBill(1).RDesc1, 21)
  PRINT #RptFile, TAB(11); LEFT$(TaxBill(1).RDesc2, 21)
  'v line 12
  PRINT #RptFile,
  
  PRINT #RptFile, USING "########,#"; TaxBill(1).RealValue; TAB(15); TaxBill(1).PersValue;
  NetTaxVal# = Round#(TaxBill(1).RealValue + TaxBill(1).PersValue)
  NetTaxVal# = Round#(NetTaxVal# - TaxBill(1).ExptValue)

  PRINT #RptFile, TAB(25); USING "######,#"; TaxBill(1).ExptValue;
  PRINT #RptFile, TAB(46); USING "#######,#.##"; TaxBill(1).TotalBillDue;
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, USING "########,#"; NetTaxVal#;
  
  TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF

  PRINT #RptFile, USING "#.##"; TAB(16); TaxRate#;
  PRINT #RptFile, TAB(21); USING "#######,#.##"; Round#(TaxBill(1).TotalBillDue - TaxBill(1).LateTaxDue);
  PRINT #RptFile, TAB(36); Cust$
  PRINT #RptFile, TAB(36); Addr1$
  PRINT #RptFile, TAB(36); Addr2$
  PRINT #RptFile, TAB(36); Addr3$
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  TwoPct# = Round#(TaxBill(1).TotalBillDue * .02)
  PRINT #RptFile, TAB(21); USING "#######,#.##"; Round#(TaxBill(1).TotalBillDue - TwoPct#)
  PRINT #RptFile, "~"


