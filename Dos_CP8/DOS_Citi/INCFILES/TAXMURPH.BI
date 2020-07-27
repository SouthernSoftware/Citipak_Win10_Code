  TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF
  NetTaxVal# = Round#(TaxBill(1).Realvalue + TaxBill(1).Persvalue)
  PRINT #RptFile, "CN"; PrnCnt
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(55); USING "Bill_####### "; TaxBill(1).BillNumber
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(10); USING "######"; TaxBill(1).Realvalue;
  PRINT #RptFile, TAB(20); USING "######"; TaxBill(1).Persvalue;
  PRINT #RptFile, TAB(33); USING "#######"; NetTaxVal#;
  PRINT #RptFile, TAB(41); USING "$$#####.##"; Round#(TaxBill(1).TotalBillDue - TaxBill(1).LateTaxDue);
  PRINT #RptFile, TAB(65); USING "##.##"; TaxRate#;
  PRINT #RptFile, TAB(75); USING "#####.##"; TaxBill(1).LateTaxDue;
  PRINT #RptFile, TAB(84); USING "$$#####,#.##"; TaxBill(1).TotalBillDue
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(20); USING "#####.##"; TaxBill(1).ExptValue
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(17); LEFT$(TaxBill(1).RDesc1, 21)
  PRINT #RptFile,
  PRINT #RptFile, TAB(20); USING "######"; TaxBill(1).CustRec
  PRINT #RptFile,
  PRINT #RptFile, TAB(10); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(10); LEFT$(TaxBill(1).CustAdd1, 25)
  PRINT #RptFile, TAB(10); LEFT$(TaxBill(1).CustAdd2, 25)
  PRINT #RptFile, TAB(10); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  PRINT #RptFile,
  PRINT #RptFile, "~"
'
'printbill:
'3150    REM GET NAME INTO N$
'        K = INSTR(firstnam$, "  ")
'        IF K = 0 THEN K = 25
'        IF K = 1 THEN N$ = lnam$ ELSE N$ = LEFT$(lnam$, K) + firstnam$
'        LOT = INSTR(prop$, "LOT")
'        IF LOT = 0 THEN GOTO CNT
'        LOT$ = MID$(prop$, LOT, 9)
'CNT:    WIDTH LPRINT 132
'        LPRINT TAB(55); "Bill # "; count
'        LPRINT
'        LPRINT
'        LPRINT
'        LPRINT TAB(10); USING "######"; realvalue#;
'        LPRINT TAB(20); USING "######"; persvalue#;
'        LPRINT TAB(33); USING "#######"; realvalue# + persvalue#;
'        LPRINT TAB(41); USING "$$#####.##"; regtax#;
'        LPRINT TAB(65); USING "##.##"; grate;
'        LPRINT TAB(75); USING "#####.##"; latetax!;
'        LPRINT TAB(84); USING "$$#####,#.##"; taxowed!
'        LPRINT
'        LPRINT
'        LPRINT
'        LPRINT TAB(15); USING "#####.##"; CVS(scexp$)
'        LPRINT
'        LPRINT
'        LPRINT TAB(17); pdesc$
'        IF BILLINFO$ = "R" THEN mestoprint$ = "": GOTO 3051
'        IF CVS(totlate$) = 0 THEN
'          mestoprint$ = ""
'        ELSE
'          mestoprint$ = "BACK TAXES OWED!"
'        END IF

'3051    LPRINT TAB(7); mestoprint$
'        LPRINT TAB(20); acct$
'        LPRINT
'        LPRINT
'        LPRINT TAB(10); N$
'        LPRINT TAB(10); A1$
'        LPRINT TAB(10); A2$
'        LPRINT TAB(10); cy$; " "; st$; " "; zp$
'        LPRINT
'        LPRINT
'        LPRINT
'        LPRINT
'        RETURN


