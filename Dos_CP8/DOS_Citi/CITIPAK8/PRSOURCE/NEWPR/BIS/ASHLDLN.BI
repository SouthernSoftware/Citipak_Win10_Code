


  'Must Calc Late Fee Here
  

  PRINT #RptFile, "~"
  PRINT #RptFile, TAB(75); USING "#####"; BillNumber
  PRINT #RptFile, TAB(75); WhatYear
  FOR LC = 4 TO 10
   PRINT #RptFile, " "
  NEXT LC
  PRINT #RptFile, TAB(5); LEFT$(TaxCust(1).CSSN, 3); "-"; MID$(TaxCust(1).CSSN, 4, 2); "-"; MID$(TaxCust(1).CSSN, 6, 4)
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(5); "Acct # "; CustRec
  PRINT #RptFile, TAB(5); LEFT$(CustName$, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxCust(1).CustAdd1, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxCust(1).CustAdd2, 25)
  PRINT #RptFile, TAB(5); QPTRIM$(TaxCust(1).CustAdd3); " "; TaxBill(1).CustZip
  FOR LC = 18 TO 20
   PRINT #RptFile, " "
  NEXT LC
   PRINT #RptFile, TAB(41); "TAX DUE"; TAB(51); "PENALTY"; ta(62); "INTEREST"; TAB(75); "TOTAL DUE"
   PRINT #RptFile, " "
  'Line 23 Starts Here
  PRINT #RptFile, TAB(2); "Personal Property/Mach Tools"; TAB(33); USING ".###"; PersTaxRate#;
   PRINT #RptFile, TAB(38); USING "#######,#"; (TaxBill(1).PersValue + TaxBill(1).MTValue);
   PRINT #RptFile, TAB(51); USING "#####,#.##"; (TaxBill(1).PersTaxDue + TaxBill(1).MTTaxDue);
   PRINT #RptFile, TAB(74); USING "#####,#.##"; (TaxBill(1).PersTaxDue + TaxBill(1).MTTaxDue)
  PRINT #RptFile, TAB(2); "Mobile Homes"; TAB(33); USING ".###"; MHTaxRate#;
   PRINT #RptFile, TAB(38); USING "#######,#"; TaxBill(1).MHValue;
   PRINT #RptFile, TAB(51); USING "#####,#.##"; TaxBill(1).MHTaxDue;
   PRINT #RptFile, TAB(74); USING "#####,#.##"; TaxBill(1).MHTaxDue
  PRINT #RptFile, TAB(2); "Qualify Volunteer"; TAB(33); USING ".###"; FETaxRate#;
   PRINT #RptFile, TAB(38); USING "#######,#"; TaxBill(1).FEValue;
   PRINT #RptFile, TAB(51); USING "#####,#.##"; TaxBill(1).FETaxDue;
   PRINT #RptFile, TAB(74); USING "#####,#.##"; TaxBill(1).FETaxDue

  'Line 26 to 35 here to print vehicles
  
    CarCount = 0

   WhatPers& = TaxCust(1).FirstPersRec
    DO
     GET PersFile, WhatPers&, PersRec(1)
     IF LEFT$(PersRec(1).Desc5, 1) = "Y" THEN
      PRINT #RptFile, TAB(2); "*" + QPTRIM$(PersRec(1).Desc4); " "; QPTRIM$(PersRec(1).Desc2); TAB(55); "Value:"; USING "$#####,#.##"; PersRec(1).PersVal
      CarCount = CarCount + 1
     END IF
     IF CarCount >= 10 THEN
        PRINT #RptFile, ""
        PRINT #RptFile, ""
        PRINT #RptFile, TAB(50); "Total Tax Due "; DueDate$;
         PRINT #RptFile, TAB(74); USING "######,#.##"; TaxBill(1).TotalBillDue
        PRINT #RptFile,
        PRINT #RptFile, TAB(50); "* Vehicle May Qualify Under PPTRA"
        PRINT #RptFile,
        PRINT #RptFile,
        PRINT #RptFile, "BN"; USING "#####"; PrnCnt
        PRINT #RptFile, "~"

        PRINT #RptFile, "~"
        PRINT #RptFile, TAB(75); USING "#####"; TaxBill(1).BillNumber
        PRINT #RptFile, TAB(75); WhatYear
        FOR LC = 4 TO 12
        PRINT #RptFile, " "
        NEXT LC
        PRINT #RptFile, TAB(5); "Acct # "; TaxBill(1).CustRec; " Vehicle Listing Cont'd"
        PRINT #RptFile, TAB(5); CustName$
        PRINT #RptFile, TAB(5); LEFT$(TaxBill(1).CustAdd1, 25)
        PRINT #RptFile, TAB(5); LEFT$(TaxBill(1).CustAdd2, 25)
        PRINT #RptFile, TAB(5); QPTRIM$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
        FOR LC = 18 TO 20
         PRINT #RptFile, " "
        NEXT LC
        PRINT #RptFile, TAB(41); "VALUE"; TAB(51); "TAX AMOUNT"; TAB(63); "LATE AMOUNT"; TAB(75); "TOTAL DUE"
        PRINT #RptFile, ""
        PRINT #RptFile, "Vehicle Listing Continued ..."
        PRINT #RptFile, ""
        PRINT #RptFile, ""
        CarCount = 0
    END IF
       WhatPers& = PersRec(1).NextRec
    LOOP WHILE WhatPers& > 0

   ' Finish the bill up here
     IF CarCount < 10 THEN
      FOR LCnt = CarCount TO 10: PRINT #RptFile, "": NEXT LCnt
     END IF
      PRINT #RptFile, ""
        PRINT #RptFile, TAB(50); "Total Tax Due "; DueDate$;
        'Put Late Here and Add to Total
        PRINT #RptFile, TAB(74); USING "#####,#.##"; TaxBill(1).TotalBillDue
   
  PRINT #RptFile,
  PRINT #RptFile,
  IF CarCount > 0 THEN
   PRINT #RptFile, TAB(50); "* Upon Payment, Vehicle Qualifies"
   PRINT #RptFile, TAB(50); "  For A Partial State Refund"
  ELSE
   PRINT #RptFile,
   PRINT #RptFile,
  END IF
  PRINT #RptFile, "BN"; USING "#####"; PrnCnt
  PRINT #RptFile, "~"


