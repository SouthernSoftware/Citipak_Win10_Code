   ' Remington, VA Tax Bill

  CLOSE PersFile
  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  CLOSE TaxFile

  TaxFile = FREEFILE
  OPEN PPTaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, TaxBill(1).CustRec, TaxCust(1)


     PRINT #RptFile, TAB(12); "DUE ON OR BEFORE "; DueDate$; "      Tax Year "; TYear$
     PRINT #RptFile, TAB(4); "TOWN OF REMINGTON   P.O. BOX 276   REMINGTON, VA 22734"; TAB(60); USING "######"; TaxBill(1).BillNumber
     PRINT #RptFile, ""
     PRINT #RptFile, ""
     PRINT #RptFile, TAB(5); "PERSONAL PROPERTY";
      PRINT #RptFile, TAB(41); USING "#.##"; TaxBill(1).PersTaxRate;
       PRINT #RptFile, TAB(47); USING "$$######,#"; TaxBill(1).PersValue;
        PRINT #RptFile, TAB(59); USING "$$######,#.##"; TaxBill(1).PersTaxDue



    CarCount = 0

   WhatPers& = TaxCust(1).FirstPersRec
     WHILE WhatPers& <> 0
     GET PersFile, WhatPers&, PersRec(1)
     IF LEFT$(PersRec(1).Desc5, 1) = "Y" OR LEFT$(PersRec(1).Desc5, 1) = "N" THEN
       VehLine$ = qptrim$(PersRec(1).Desc4) + " " + qptrim$(PersRec(1).Desc2) + " " + qptrim$(PersRec(1).Desc1) + " " + qptrim$(MID$(PersRec(1).Desc5, 2, 2)) + " " + LEFT$(PersRec(1).Desc5, 1)
      PRINT #RptFile, TAB(4); LEFT$(VehLine$, 40);
       PRINT #RptFile, TAB(47); USING "$$######,#"; PersRec(1).PersVal
      CarCount = CarCount + 1
     END IF
     IF CarCount = 4 AND PersRec(1).NextRec > 0 THEN
        PRINT #RptFile, TAB(5); "THERE IS A 10% PENALTY IF PAID AFTER "; DueDate$
        PRINT #RptFile, ""
        PRINT #RptFile, TAB(7); "Acct#"; TaxBill(1).CustRec; TAB(47); DueDate$; TAB(59); USING "$$######,#.##"; TaxBill(1).TotalBillDue
        PRINT #RptFile,
        PRINT #RptFile, TAB(7); TaxBill(1).CustName
        PRINT #RptFile, TAB(7); TaxBill(1).CustAdd1
        PRINT #RptFile, TAB(7); TaxBill(1).CustAdd2
        PRINT #RptFile, TAB(7); RTRIM$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
        PRINT #RptFile,
        PRINT #RptFile,
        PRINT #RptFile, TAB(2); TaxCust(1).CSSN
        PRINT #RptFile,

     PRINT #RptFile, TAB(12); "DUE ON OR BEFORE "; DueDate$; "      Tax Year "; TYear$
     PRINT #RptFile, TAB(4); "TOWN OF REMINGTON   P.O. BOX 276   REMINGTON, VA 22734"; TAB(60); USING "######"; TaxBill(1).BillNumber
     PRINT #RptFile, ""
     PRINT #RptFile, ""
     PRINT #RptFile, TAB(5); "PERSONAL PROPERTY VEHICLE LISTING CONT'D"
     CarCount = 0
     Continued = 1
     END IF

     WhatPers& = PersRec(1).NextRec
    WEND
      
     FOR SkipLine = CarCount + 1 TO 4
      PRINT #RptFile, " "
     NEXT SkipLine

     PRINT #RptFile, TAB(5); "THERE IS A 10% PENALTY IF PAID AFTER "; DueDate$
     
     PRINT #RptFile, ""
     PRINT #RptFile, TAB(7); "Acct#"; TaxBill(1).CustRec;

     IF Continued = 1 THEN
      PRINT #RptFile, ""
      ELSE
      PRINT #RptFile, TAB(47); DueDate$; TAB(59); USING "$$######,#.##"; TaxBill(1).TotalBillDue
     END IF
     PRINT #RptFile,
     REM LINE 14
     PRINT #RptFile, TAB(7); TaxBill(1).CustName
     PRINT #RptFile, TAB(7); TaxBill(1).CustAdd1
     PRINT #RptFile, TAB(7); TaxBill(1).CustAdd2
     PRINT #RptFile, TAB(7); RTRIM$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
     PRINT #RptFile,
     PRINT #RptFile,
     PRINT #RptFile, TAB(2); TaxCust(1).CSSN
     PRINT #RptFile,
     Continued = 0


