'  CLOSE TaxFile
'  TaxFile = FREEFILE
'  OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
'  GET TaxFile, TaxBill(1).CustRec, TaxCust(1)
'
'  CustName$ = RTRIM$(TaxCust(1).FNAME) + " " + LTRIM$(TaxCust(1).LName)
'  CustName$ = LTRIM$(CustName$)

  'Must Calc Late Fee Here

  DueDate$ = "12-31-" + QPTrim$(STR$(TaxBill(1).TaxYear))
  TaxRate# = TaxBill(1).RealTaxRate
  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
  END IF

  NetTaxVal# = Round#(TaxBill(1).RealValue + TaxBill(1).PersValue)
  NetTaxVal# = Round#(NetTaxVal# - TaxBill(1).ExptValue)

  'IF TaxBill(1).ExptValue > 0 THEN STOP

  PRINT #RptFile,
  PRINT #RptFile, TAB(29); "TOWN OF MAGGIE VALLEY"
  PRINT #RptFile, TAB(29); "    3987 SOCO RD."
  PRINT #RptFile, TAB(29); "MAGGIE VALLEY NC 28751"
  PRINT #RptFile, TAB(29); "  PROPERTY TAX BILL"

  FOR LC = 1 TO 3
    PRINT #RptFile, " "
  NEXT

  PRINT #RptFile, TAB(12); "ACCT # "; TaxBill(1).CustRec;
  PRINT #RptFile, TAB(65); "BILL #"; USING "######"; TaxBill(1).BillNumber
  PRINT #RptFile, TAB(12); LEFT$(CustName$, 25);
  PRINT #RptFile, TAB(63); "TAX YEAR "; TaxBill(1).TaxYear
  PRINT #RptFile, TAB(12); LEFT$(TaxBill(1).CustAdd1, 25);
  PRINT #RptFile, TAB(63); "TAX RATE "; USING "#.###"; TaxRate#
  PRINT #RptFile, TAB(12); LEFT$(TaxBill(1).CustAdd2, 25)
  PRINT #RptFile, TAB(12); QPTrim$(TaxBill(1).CustAdd3); " "; LEFT$(TaxBill(1).CustZip, 5) + "-" + MID$(TaxBill(1).CustZip, 6, 4)
  FOR LC = 1 TO 4
    PRINT #RptFile, " "
  NEXT
  PRINT #RptFile, TAB(39); "[--------- VALUATIONS --------]"
  PRINT #RptFile, TAB(2); "PROPERTY DESCRIPTION"; TAB(30); "RATE"; TAB(40); "REAL"; TAB(48); "PERSONAL"; TAB(61); "EXEMPT"; TAB(72); "TOTAL"
  PRINT #RptFile, " "
  'Line 23 Starts Here
  PRINT #RptFile, TAB(30); USING ".##"; TaxRate#;
  PRINT #RptFile, TAB(35); USING "########,#"; TaxBill(1).RealValue;
  PRINT #RptFile, TAB(47); USING "#######,#"; TaxBill(1).PersValue;
  PRINT #RptFile, TAB(59); USING "######,#"; TaxBill(1).ExptValue;
  PRINT #RptFile, TAB(68); USING "########,#"; (TaxBill(1).PersValue + TaxBill(1).RealValue)
  PRINT #RptFile, TAB(2); QPTrim$(TaxBill(1).RDesc1)
  PRINT #RptFile, TAB(2); QPTrim$(PINTemp(1))


   PRINT #RptFile, ""
   PRINT #RptFile, ""
   PRINT #RptFile, ""
   PRINT #RptFile, TAB(2); "NOTE: "
   PRINT #RptFile, TAB(2); "      A 2% PENALTY WILL BE ADDED AFTER DUE DATE."
   PRINT #RptFile, TAB(2); "      .75% ADDED ON FIRST OF EACH MONTH THEREAFTER."
   PRINT #RptFile, ""
   PRINT #RptFile, TAB(2); "      PLEASE SUBMIT YOUR ACCOUNT# OR, PARCEL ID# ON CHECK"
   PRINT #RptFile, TAB(2); "      TO PROCESS YOUR PAYMENT. CONTACT THE TAX COLLECTOR"
   PRINT #RptFile, TAB(2); "      IF YOU HAVE IF ANY QUESTIONS."
   PRINT #RptFile, TAB(2); "      (PHONE: 828-926-0866  EXT. 101)"
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile, TAB(39); " TAX DUE DATE: "; DueDate$
   PRINT #RptFile,
   PRINT #RptFile, TAB(39); USING "TOTAL TAX DUE: $$####,#.##"; TaxBill(1).TotalBillDue
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile, "BN"; USING "#####"; PrnCnt
   PRINT #RptFile, CHR$(12);
  

