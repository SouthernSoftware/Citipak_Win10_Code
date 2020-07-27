 'INDEPENDENCE PERS PROPERTY INCLUDE FILE

 'Added May 10, 2002 per independence va instruction
  TxFile = FREEFILE
  OPEN "TAXPP97.DAT" FOR RANDOM AS #TxFile LEN = LEN(Tax97PPRate)
  IF LOF(TxFile) > 0 THEN
    GET TxFile, 1, Tax97PPRate
    TaxRate97! = Tax97PPRate.Rate
  ELSE
    TaxRate97! = PersTaxRate#
  END IF
  CLOSE TxFile
 
 'Determine Perc Discnt
  IF WhatYear = 1999 THEN PERC! = 27.5
  IF WhatYear = 2000 THEN PERC! = 47.5
  IF WhatYear >= 2001 THEN PERC! = 70
  'IF WhatYear = 2002 THEN PERC! = 70

  CLOSE PersFile
  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen
  CLOSE TaxFile
  TaxFile = FREEFILE
  OPEN PPTaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, TaxBill(1).CustRec, TaxCust(1)
  CustName$ = RTRIM$(TaxCust(1).LName) + "  " + RTRIM$(TaxCust(1).FNAME)
  CustName$ = LTRIM$(CustName$)
  
  'Must Calc Late Fee Here
  
  PRINT #RptFile,
  PRINT #RptFile, TAB(30); "TOWN OF INDEPENDENCE"
  PRINT #RptFile, TAB(30); "    PO BOX 99"
  PRINT #RptFile, TAB(29); "INDEPENDENCE, VA 24348"
  PRINT #RptFile, TAB(27); "PERSONAL PROPERTY TAX BILL"
  FOR LC = 6 TO 8
    PRINT #RptFile, " "
  NEXT LC
  PRINT #RptFile, TAB(12); "ACCT # "; TaxBill(1).CustRec;
  PRINT #RptFile, TAB(65); "BILL #"; USING "######"; TaxBill(1).BillNumber
  PRINT #RptFile, TAB(12); LEFT$(CustName$, 25);
  PRINT #RptFile, TAB(63); "TAX YEAR "; WhatYear
  PRINT #RptFile, TAB(12); LEFT$(TaxBill(1).CustAdd1, 35)
  PRINT #RptFile, TAB(12); LEFT$(TaxBill(1).CustAdd2, 35)
  PRINT #RptFile, TAB(12); QPTRIM$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  FOR LC = 14 TO 17
    PRINT #RptFile, " "
  NEXT LC
  PRINT #RptFile, TAB(40); "PROPERTY"; TAB(51); " TAX"; TAB(61); " PPTRA"
  PRINT #RptFile, TAB(31); "Rate"; TAB(41); "VALUE"; TAB(51); "AMOUNT"; TAB(61); "DISCOUNT"; TAB(71); "TOTAL DUE"
   'Line 23 Starts Here

  PRINT #RptFile, TAB(2); "Personal Property";
  PRINT #RptFile, TAB(30); USING "#.###"; PersTaxRate#;
  PRINT #RptFile, TAB(38); USING "######,#"; TaxBill(1).PersValue;
  PRINT #RptFile, TAB(48); USING "#####,#.##"; (TaxBill(1).PersTaxDue);
  PRINT #RptFile, TAB(60); USING "####.##"; TaxBill(1).PPTRADiscnt;
  PRINT #RptFile, TAB(70); USING "#####,#.##"; (TaxBill(1).PersTaxDue - TaxBill(1).PPTRADiscnt)
  PRINT #RptFile, TAB(2); "Mobile Homes";
  PRINT #RptFile, TAB(30); USING "#.###"; MHTaxRate#;
  PRINT #RptFile, TAB(38); USING "######,#"; TaxBill(1).MHValue;
  PRINT #RptFile, TAB(48); USING "#####,#.##"; TaxBill(1).MHTaxDue;
  PRINT #RptFile, TAB(70); USING "#####,#.##"; TaxBill(1).MHTaxDue
  PRINT #RptFile, TAB(2); "Farm Equipment";
  PRINT #RptFile, TAB(30); USING "#.###"; FETaxRate#;
  PRINT #RptFile, TAB(38); USING "######,#"; TaxBill(1).FEValue;
  PRINT #RptFile, TAB(48); USING "#####,#.##"; TaxBill(1).FETaxDue;
  PRINT #RptFile, TAB(70); USING "#####,#.##"; TaxBill(1).FETaxDue
  PRINT #RptFile, TAB(2); "Machinery/Tools";
  PRINT #RptFile, TAB(30); USING "#.###"; MTTaxRate#;
  PRINT #RptFile, TAB(38); USING "######,#"; TaxBill(1).MTValue;
  PRINT #RptFile, TAB(48); USING "#####,#.##"; TaxBill(1).MTTaxDue;
  PRINT #RptFile, TAB(70); USING "#####,#.##"; TaxBill(1).MTTaxDue
  CarCount = 0
  PRINT #RptFile,
  PRINT #RptFile, " PPTRA Information"
  
  'Do Not Process This Record

   WhatPers& = TaxCust(1).FirstPersRec
    DO
     GET PersFile, WhatPers&, PersRec(1)

    'Determine Year of Car
     PYear$ = LTRIM$(RIGHT$(PersRec(1).Desc5, 5))
     PYear = VAL(PYear$)

   IF LEFT$(PersRec(1).Desc5, 1) = "Y" AND PYear = WhatYear THEN
  
      IF Round#(PersRec(1).PERSVAL) > 20000 THEN
        PPTRAVal# = 20000
       ELSE
        PPTRAVal# = Round#(PersRec(1).PERSVAL)
      END IF
      IF PPTRAVal# <= 1000 THEN
        PPTRADiscount# = Round#((Round#(PPTRAVal# / 100) * TaxRate97!))
        ELSE
        PPTRADiscount# = Round#((Round#((PPTRAVal# / 100) * (PERC! / 100)) * TaxRate97!))
      END IF

      VehDesc$ = " VIN# " + QPTRIM$(PersRec(1).Desc1)
      VehDesc$ = QPTRIM$(VehDesc$)
      TaxAmt# = (PersTaxRate# / 100) * PersRec(1).PERSVAL

      PRINT #RptFile, TAB(2); "*" + VehDesc$;
       PRINT #RptFile, TAB(38); USING "######,#"; PersRec(1).PERSVAL;
       PRINT #RptFile, TAB(48); USING "#####,#.##"; TaxAmt#;
       PRINT #RptFile, TAB(60); USING "####.##"; PPTRADiscount#
      CarCount = CarCount + 1
    END IF

      WhatPers& = PersRec(1).NextRec
    LOOP WHILE WhatPers& > 0

   ' Finish the bill up here
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(40); "Total Tax Due by "; DueDate$;
  'Put Late Here and Add to Total
  PRINT #RptFile, TAB(68); USING "$$#####,#.##"; TaxBill(1).TotalBillDue
   
  PRINT #RptFile,
  PRINT #RptFile,
  IF CarCount > 0 THEN
    PRINT #RptFile, " The tax on the first $20,000 of value of your personal car, motorcycle,"
    PRINT #RptFile, " and pickup or panel truck under 7,501 pounds, which is a qualifying "
    PRINT #RptFile, " vehicle, has been reduced by 27.5% this year.  In future years, it is"
    PRINT #RptFile, " scheduled to be reduced by 47.5% for the year 2000, 70% for 2001, and "
    PRINT #RptFile, " beyond.  For qualifying vehicles valued at $1,000 or less, your"
    PRINT #RptFile, " obligation to pay this tax has been eliminated.  These reductions are"
    PRINT #RptFile, " based on local tax rates in effect on August 1, 1997."
  END IF
  PRINT #RptFile,
  PRINT #RptFile, " BN"; USING "#####"; PrnCnt
  PRINT #RptFile, CHR$(12);


