  'Determine Perc Discnt
  IF WhatYear = 1999 THEN PERC! = 27.5
  IF WhatYear = 2000 THEN PERC! = 47.5
  IF WhatYear >= 2001 THEN PERC! = 70
  
  CLOSE PersFile
  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  
  CLOSE TaxFile
  TaxFile = FREEFILE
  OPEN PPTaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, TaxBill(1).CustRec, TaxCust(1)
  
  CustName$ = RTRIM$(TaxCust(1).FNAME) + " " + RTRIM$(TaxCust(1).LName)
  CustName$ = LTRIM$(CustName$)
  
  'Must Calc Late Fee Here
  
  PRINT #RptFile, "~"
  PRINT #RptFile, TAB(65); "Tax Year: "; WhatYear
  PRINT #RptFile, TAB(75); USING "#####"; TaxBill(1).BillNumber
  
  FOR LC = 4 TO 12
    PRINT #RptFile, " "
  NEXT LC
  PRINT #RptFile, TAB(5); "Acct # "; TaxBill(1).CustRec
  PRINT #RptFile, TAB(5); CustName$
  PRINT #RptFile, TAB(5); TaxBill(1).CustAdd1
  PRINT #RptFile, TAB(5); TaxBill(1).CustAdd2
  PRINT #RptFile, TAB(5); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  FOR LC = 18 TO 21
    PRINT #RptFile, " "
  NEXT LC
  
  PRINT #RptFile, TAB(38); "PROP VALUE"; TAB(51); "TAX AMOUNT"; TAB(63); "PPTRA DIS."; TAB(74); "TOT DUE"
  PRINT #RptFile, " "
  
  'Line 24 Starts Here
  PRINT #RptFile, TAB(2); "Machinery/Tools"; TAB(33); USING "#.##"; MTTaxRate#;
  PRINT #RptFile, TAB(38); USING "#######,#"; TaxBill(1).MTValue;
  PRINT #RptFile, TAB(51); USING "######.##"; TaxBill(1).MTTaxDue;
  PRINT #RptFile, TAB(72); USING "######.##"; TaxBill(1).MTTaxDue
  
  PRINT #RptFile, TAB(2); "Personal Property"; TAB(33); USING "#.##"; PersTaxRate#;
  PRINT #RptFile, TAB(38); USING "#######,#"; (TaxBill(1).PersValue);
  PRINT #RptFile, TAB(51); USING "######.##"; (TaxBill(1).PersTaxDue);
  PRINT #RptFile, TAB(62); USING "#####.##"; TaxBill(1).PPTRADiscnt;
  PRINT #RptFile, TAB(72); USING "######.##"; (TaxBill(1).PersTaxDue - TaxBill(1).PPTRADiscnt)
  
  PRINT #RptFile, TAB(2); "Mobile Home"; TAB(33); USING "#.##"; MHTaxRate#;
  PRINT #RptFile, TAB(38); USING "#######,#"; (TaxBill(1).MHValue);
  PRINT #RptFile, TAB(51); USING "######.##"; (TaxBill(1).MHTaxDue);
  'PRINT #RptFile, TAB(62); USING "#####.##"; TaxBill(1).PPTRADiscnt;
  PRINT #RptFile, TAB(72); USING "######.##"; TaxBill(1).MHTaxDue
  
  PRINT #RptFile, " Property Information"
  'Line 27 to 35 here to print vehicles
  CarCount = 0
  WhatPers& = TaxCust(1).FirstPersRec
  DO
    GET PersFile, WhatPers&, PersRec(1)
    PropYear& = VAL(QPTrim$(RIGHT$(PersRec(1).DESC5, 6)))
    IF PropYear& = WhatYear THEN
      VehDesc$ = "Desc: " + QPTrim$(PersRec(1).Desc1)
      VehDesc$ = QPTrim$(VehDesc$)
      IF (LEFT$(PersRec(1).DESC5, 1) = "Y") THEN 'OR (PersRec(1).MHValue > 0) THEN
        IF Round#(PersRec(1).PERSVAL) > 20000 THEN
          PPTRAVal# = 20000
        ELSE
          PPTRAVal# = Round#(PersRec(1).PERSVAL)
        END IF
        IF PPTRAVal# <= 1000 THEN
          PPTRADiscount# = Round#((Round#(PPTRAVal# / 100) * PersTaxRate#))
        ELSE
          PPTRADiscount# = Round#((Round#((PPTRAVal# / 100) * (PERC! / 100)) * PersTaxRate#))
        END IF
        TaxAmt# = (PersTaxRate# / 100) * PersRec(1).PERSVAL
        PRINT #RptFile, TAB(2); "*" + VehDesc$;
        PRINT #RptFile, TAB(38); USING "#######,#"; PersRec(1).PERSVAL;
        PRINT #RptFile, TAB(51); USING "#####,#.##"; TaxAmt#;
        PRINT #RptFile, TAB(63); USING "####,#.##"; PPTRADiscount#
        CarCount = CarCount + 1
      ELSEIF (PersRec(1).MHValue > 0) THEN
        TaxAmt# = (MHTaxRate# / 100) * PersRec(1).MHValue
        PRINT #RptFile, TAB(2); "*" + VehDesc$;
        PRINT #RptFile, TAB(38); USING "#######,#"; PersRec(1).MHValue;
        PRINT #RptFile, TAB(51); USING "#####,#.##"; TaxAmt#
        CarCount = CarCount + 1
      ELSE
        TaxAmt# = (PersTaxRate# / 100) * PersRec(1).PERSVAL
        PRINT #RptFile, TAB(2); "*" + VehDesc$;
        PRINT #RptFile, TAB(38); USING "#######,#"; PersRec(1).PERSVAL;
        PRINT #RptFile, TAB(51); USING "######.##"; TaxAmt#
        CarCount = CarCount + 1
      END IF
    
      IF CarCount >= 8 THEN
        PRINT #RptFile, ""
        PRINT #RptFile, ""
        PRINT #RptFile, TAB(50); "Total Tax Due "; DueDate$;
        PRINT #RptFile, TAB(71); USING "$$#####.##"; TaxBill(1).TotalBillDue
        PRINT #RptFile,
        PRINT #RptFile,
        PRINT #RptFile,
        PRINT #RptFile,
        PRINT #RptFile, "BN"; USING "#####"; PrnCnt
        PRINT #RptFile, "~"
      
        PRINT #RptFile, "~"
        PRINT #RptFile, TAB(75); WhatYear
        PRINT #RptFile, TAB(75); USING "#####"; TaxBill(1).BillNumber
        FOR LC = 4 TO 12
          PRINT #RptFile, " "
        NEXT
        PRINT #RptFile, TAB(5); "Acct # "; TaxBill(1).CustRec; " Vehicle Listing Cont'd"
        PRINT #RptFile, TAB(5); CustName$
        PRINT #RptFile, TAB(5); TaxBill(1).CustAdd1
        PRINT #RptFile, TAB(5); TaxBill(1).CustAdd2
        PRINT #RptFile, TAB(5); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
        FOR LC = 18 TO 21
          PRINT #RptFile, " "
        NEXT
        PRINT #RptFile, TAB(38); "PROP VALUE"; TAB(51); "TAX AMOUNT"; TAB(63); "PPTRA DIS."; TAB(74); "TOT DUE"
        PRINT #RptFile, " "
        PRINT #RptFile, "Property Information Continued ..."
        PRINT #RptFile, ""
        PRINT #RptFile, ""
        CarCount = 0
      END IF

    END IF
    WhatPers& = PersRec(1).NextRec
  LOOP WHILE WhatPers& > 0
  
  ' Finish the bill up here
  IF CarCount < 8 THEN
    FOR LCnt = CarCount TO 8
      PRINT #RptFile, ""
    NEXT LCnt
  END IF
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(50); "Tax Due "; DueDate$;
  'Put Late Here and Add to Total
  PRINT #RptFile, TAB(70); USING "$$######.##"; TaxBill(1).TotalBillDue
  
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, "BN"; USING "#####"; PrnCnt
  PRINT #RptFile, "~"
  
  
  
  
  
  

