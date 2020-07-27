  

  CLOSE TaxFile
  TaxFile = FREEFILE
  OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, TaxBillRec(1).CustRec, TaxCust(1)

  CustName$ = RTRIM$(TaxCust(1).FNAME) + " " + LTRIM$(TaxCust(1).LName)
  CustName$ = LTRIM$(CustName$)

 'Get Mortgage Code Information Here
  IF TaxCust(1).FirstPropRec > 0 THEN
   RealFile = FREEFILE
   OPEN TaxPropFile FOR RANDOM SHARED AS RealFile LEN = RealRecLen
   GET RealFile, TaxCust(1).FirstPropRec, PropRec(1)
   MortCode$ = qptrim$(PropRec(1).MortCode)
   CLOSE RealFile
  END IF
   
 'If MortCode$ is Not Blank the Get the Mort Code Name
  MortName$ = ""
   IF LEN(MortCode$) > 0 THEN
    MortFile = FREEFILE
    OPEN "TAXMORT.DAT" FOR RANDOM AS MortFile LEN = MortCodeRecLen
    NumOfCatRecs = LOF(MortFile) \ MortCodeRecLen
    FOR MCnt = 1 TO NumOfCatRecs
      GET MortFile, MCnt, MortCodeRec(1)
      IF RTRIM$(MortCode$) = RTRIM$(MortCodeRec(1).MortCode) THEN MortName$ = MortCodeRec(1).BName: CLOSE MortFile: EXIT FOR
    NEXT MCnt
    CLOSE MortFile
   END IF
  
  CLOSE TaxFile
  


  'Must Calc Late Fee Here
  

  PRINT #RptFile, "~"
  PRINT #RptFile, TAB(75); USING "#####"; TaxBillRec(1).BillNumber
  PRINT #RptFile, TAB(75); WhatYear
  FOR LC = 4 TO 10
   PRINT #RptFile, " "
  NEXT LC
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(5); "Acct # "; TaxBillRec(1).CustRec
  PRINT #RptFile, TAB(5); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxBillRec(1).CustAdd1, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxBillRec(1).CustAdd2, 25)
  PRINT #RptFile, TAB(5); qptrim$(TaxBillRec(1).CustAdd3); " "; TaxBillRec(1).CustZip
  FOR LC = 18 TO 20
   PRINT #RptFile, " "
  NEXT LC
   PRINT #RptFile, TAB(43); "LAND"; TAB(52); "BUILDING"; TAB(66); "TOTAL"; TAB(75); "TOTAL DUE"
   PRINT #RptFile, " "
  'Line 23 Starts Here
  PRINT #RptFile, TAB(2); qptrim$(TaxBillRec(1).RDesc1);
   PRINT #RptFile, TAB(33); USING ".##"; RealTaxRate#;
   PRINT #RptFile, TAB(37); USING "########,#"; (TaxBillRec(1).RealValue);
   PRINT #RptFile, TAB(51); USING "#######,#"; TaxBillRec(1).PersValue;
   PRINT #RptFile, TAB(62); USING "#######,#"; (TaxBillRec(1).RealValue + TaxBillRec(1).PersValue);
   PRINT #RptFile, TAB(74); USING "######,#.##"; (TaxBillRec(1).RealTaxDue)
  PRINT #RptFile, TAB(2); qptrim$(TaxBillRec(1).RDesc2)
  PRINT #RptFile, TAB(2); "GPIN # "; TaxBillRec(1).RealPin
  IF TaxBillRec(1).ExptValue > 0 THEN
   PRINT #RptFile, TAB(2); "Exemption Amount: "; TAB(74); USING "######,#.##"; (TaxBillRec(1).ExptValue)
  ELSE
   PRINT #RptFile, ""
  END IF
  
  IF LEN(MortName$) > 0 THEN
  PRINT #RptFile, " Mortgage Company: "; MortName$
  ELSE
  PRINT #RptFile, ""
  END IF

  'Lines 28 to 36 are blank
   FOR LCnt! = 28 TO 32: PRINT #RptFile, "": NEXT LCnt!
   PRINT #RptFile, ""
   IF TaxBillRec(1).PriorYrBalance > 0 THEN
    PRINT #RptFile, TAB(2); "** Delinquent Balance Due On This Account ** "
   ELSE
   PRINT #RptFile, ""
   END IF
   PRINT #RptFile, ""
   PRINT #RptFile, ""
 'Line 37 for Totals
   PRINT #RptFile, ""
    PRINT #RptFile, TAB(49); "Total Tax Due "; DueDate$;
    PRINT #RptFile, TAB(74); USING "$$####,#.##"; TaxBillRec(1).TotalBillDue - TaxBillRec(1).ExptValue
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile,
   PRINT #RptFile, "BN"; USING "#####"; PrnCnt
   PRINT #RptFile, "~"
  

