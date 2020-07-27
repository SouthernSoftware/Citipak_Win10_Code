  TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF
  
  PRINT #RptFile, "~"
  PRINT #RptFile, TAB(29); TaxBill(1).TaxYear; TAB(38); USING "#.##"; TaxRate#;
   PRINT #RptFile, TAB(69); USING "######"; TaxBill(1).CustRec;
   PRINT #RptFile, TAB(80); USING "######"; TaxBill(1).BillNumber
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(69); TaxBill(1).RealPin

  FOR LC! = 5 TO 7
   PRINT #RptFile, ""
  NEXT LC!

  NetTaxVal# = Round#(TaxBill(1).RealValue + TaxBill(1).PersValue)

  PRINT #RptFile, USING "######,#"; TaxBill(1).PersValue;
   PRINT #RptFile, TAB(13); USING "#######,#"; TaxBill(1).RealValue;
   PRINT #RptFile, TAB(25); USING "#######,#"; NetTaxVal#;

  NetTaxVal# = Round#(NetTaxVal# - TaxBill(1).ExptValue)


   PRINT #RptFile, TAB(35); USING "######.##"; (TaxBill(1).TotalBillDue - TaxBill(1).LateTaxDue);
   PRINT #RptFile, TAB(46); USING "######,#"; TaxBill(1).ExptValue;

   PRINT #RptFile, TAB(57); USING "####.##"; TaxBill(1).CarShore;
   PRINT #RptFile, TAB(72); USING "####.##"; TaxBill(1).LateTaxDue;
   PRINT #RptFile, TAB(79); USING "#####.##"; Round#(TaxBill(1).TotalBillDue + TaxBill(1).CarShore)
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); "IF TAXES ARE ESCROWED SEND BILL TO"
  PRINT #RptFile, TAB(5); "MORTGAGE COMPANY."
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxBill(1).CustAdd1, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxBill(1).CustAdd2, 25)
  PRINT #RptFile, TAB(5); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, "~"

'  LPCnt = LPCnt + 1
'  IF LPCnt > 1 THEN
'    PRINT #RptFile, CHR$(12)
'    LPCnt = 0
'  END IF


