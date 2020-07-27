  IF WhatYear = 1999 THEN PERC! = 27.5
  IF WhatYear = 2000 THEN PERC! = 47.5
  IF WhatYear = 2001 THEN PERC! = 70
  IF WhatYear >= 2002 THEN PERC! = 70

  CLOSE PersFile
  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersRecLen

  CLOSE TaxFile
  TaxFile = FREEFILE
  OPEN PPTaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, TaxBill(1).CustRec, TaxCust(1)

  CustName$ = RTRIM$(TaxCust(1).LName) + "  " + RTRIM$(TaxCust(1).FNAME)
  CustName$ = LTRIM$(CustName$)

  PRINT #RptFile, "~"
  PRINT #RptFile, TAB(63); "TAX YEAR: "; WhatYear
  PRINT #RptFile, TAB(75); USING "#####"; TaxBill(1).BillNumber
  PRINT #RptFile, " "
  PRINT #RptFile, " "
  PRINT #RptFile, TAB(5); TownName$
  PRINT #RptFile, TAB(5); Add1$
  PRINT #RptFile, TAB(5); Add2$
  PRINT #RptFile, TAB(5); Add3$
  PRINT #RptFile, " "
  PRINT #RptFile, " "
  PRINT #RptFile, " "
  PRINT #RptFile, TAB(5); "Acct # "; TaxBill(1).CustRec
  PRINT #RptFile, TAB(5); CustName$
  PRINT #RptFile, TAB(5); TaxBill(1).CustAdd1
  PRINT #RptFile, TAB(5); TaxBill(1).CustAdd2
  PRINT #RptFile, TAB(5); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  FOR LC = 18 TO 21
   PRINT #RptFile, " "
  NEXT LC
  PRINT #RptFile, TAB(37); "PROP VALUE"; TAB(50); "TAX AMOUNT"; TAB(62); "PPTRA DIS"; TAB(72); "TOTAL DUE"
  PRINT #RptFile, " "
  PRINT #RptFile, "Personal Property"; TAB(32); USING "#.##"; PersTaxRate#;
   PRINT #RptFile, TAB(37); USING "#######,#"; TaxBill(1).PersValue;
   PRINT #RptFile, TAB(50); USING "#####,#.##"; TaxBill(1).PersTaxDue;
   PRINT #RptFile, TAB(62); USING "####,#.##"; TaxBill(1).PPTRADiscnt;
   PRINT #RptFile, TAB(71); USING "#####,#.##"; (TaxBill(1).PersTaxDue - TaxBill(1).PPTRADiscnt)
  PRINT #RptFile, "Machinery/Tools"; TAB(32); USING "#.##"; MTTaxRate#;
   PRINT #RptFile, TAB(37); USING "#######,#"; TaxBill(1).MTValue;
   PRINT #RptFile, TAB(50); USING "#####,#.##"; TaxBill(1).MTTaxDue;
   PRINT #RptFile, TAB(71); USING "#####,#.##"; TaxBill(1).MTTaxDue
  PRINT #RptFile, "Mobile Homes";
   PRINT #RptFile, TAB(32); USING "#.##"; MHTaxRate#;
   PRINT #RptFile, TAB(37); USING "#######,#"; TaxBill(1).MHValue;
   PRINT #RptFile, TAB(50); USING "#####,#.##"; TaxBill(1).MHTaxDue;
   PRINT #RptFile, TAB(71); USING "#####,#.##"; TaxBill(1).MHTaxDue
  PRINT #RptFile, "Farm Equipment";
   PRINT #RptFile, TAB(32); USING "#.##"; FETaxRate#;
   PRINT #RptFile, TAB(37); USING "#######,#"; TaxBill(1).FEValue;
   PRINT #RptFile, TAB(50); USING "#####,#.##"; TaxBill(1).FETaxDue;
   PRINT #RptFile, TAB(71); USING "#####,#.##"; TaxBill(1).FETaxDue
  PRINT #RptFile, ""
  PRINT #RptFile, "*PPTRA Vehicle Information"
 'Line 30 to 35 here to print vehicles
   CarCount = 0
   WhatPers& = TaxCust(1).FirstPersRec
    DO
     GET PersFile, WhatPers&, PersRec(1)

    PYear$ = RIGHT$(PersRec(1).Desc5, 5)
    PYear = VAL(PYear$)
    IF PYear > 0 AND PYear <> WhatYear THEN
    'Do Not Process This Record
    ELSE


   IF LEFT$(PersRec(1).Desc5, 1) = "Y" AND PersRec(1).PersVal > 0 THEN
       
       Rate = VAL(MID$(PersRec(1).Desc5, 2, 2))
       IF Rate > 0 THEN Factor! = Rate / 12: ProRateFlag$ = "Y" ELSE Factor! = 1
      
      IF Round#(PersRec(1).PersVal * Factor!) > 20000 THEN
        PPTRAVal# = 20000
       ELSE
        PPTRAVal# = Round#(PersRec(1).PersVal * Factor!)
      END IF
      IF PPTRAVal# <= (1000 * Factor!) THEN
        PPTRADiscount# = Round#((PPTRAVal# / 100) * PersTaxRate#)
        ELSE
        PPTRADiscount# = Round#(((PPTRAVal# / 100) * (PERC! / 100)) * PersTaxRate#)
      END IF
      
      VehDesc$ = RTRIM$(PersRec(1).Desc2$)
      VehDesc$ = QPTrim$(VehDesc$)
      TaxAmt# = (PersTaxRate# / 100) * (PersRec(1).PersVal * Factor!)
      PRINT #RptFile, "*" + VehDesc$;
       PRINT #RptFile, TAB(37); USING "#######,#"; PersRec(1).PersVal;
       PRINT #RptFile, TAB(50); USING "#####,#.##"; TaxAmt#;
       PRINT #RptFile, TAB(62); USING "####,#.##"; PPTRADiscount#
      PRINT #RptFile, " VIN# "; RTRIM$(PersRec(1).Desc1);
       IF VAL(MID$(PersRec(1).Desc5, 2, 2)) > 0 THEN
        PRINT #RptFile, " MTHS "; MID$(PersRec(1).Desc5, 2, 2)
        ELSE
        PRINT #RptFile, ""
       END IF

      CarCount = CarCount + 2
    END IF
    END IF
     IF CarCount >= 6 AND PersRec(1).NextRec <> 0 THEN
        PRINT #RptFile, ""
        PRINT #RptFile, TAB(48); "Total Tax Due ";
         PRINT #RptFile, USING "$$#####,#.##"; TaxBill(1).TotalBillDue
        PRINT #RptFile, TAB(48); "Tax Due Date: "; DueDate$
        PRINT #RptFile, ""
        PRINT #RptFile, TAB(48); "10% Penalty After 01/06/"; QPTrim$(STR$(WhatYear + 1)); "."
        PRINT #RptFile, TAB(48); "10% Interest Accrues Annually."
        PRINT #RptFile, 'TAB(48); "After 04/01/"; QPTrim$(STR$(WhatYear + 1)); "."
        PRINT #RptFile, "BN"; USING "#####"; PrnCnt
        PRINT #RptFile, "~"

        PRINT #RptFile, "~"
        PRINT #RptFile, TAB(62); "TAX YEAR: "; WhatYear
        PRINT #RptFile, TAB(75); USING "#####"; TaxBill(1).BillNumber
        PRINT #RptFile, " "
        PRINT #RptFile, " "
        PRINT #RptFile, TAB(5); TownName$
        PRINT #RptFile, TAB(5); Add1$
        PRINT #RptFile, TAB(5); Add2$
        PRINT #RptFile, TAB(5); Add3$
        PRINT #RptFile, " "
        PRINT #RptFile, " "
        PRINT #RptFile, " "
        PRINT #RptFile, TAB(5); "Acct # "; TaxBill(1).CustRec; " Vehicle Listing Cont'd"
        PRINT #RptFile, TAB(5); CustName$
        PRINT #RptFile, TAB(5); TaxBill(1).CustAdd1
        PRINT #RptFile, TAB(5); TaxBill(1).CustAdd2
        PRINT #RptFile, TAB(5); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
        FOR LC = 18 TO 21
         PRINT #RptFile, " "
        NEXT LC
        PRINT #RptFile, TAB(37); "PROP VALUE"; TAB(50); "TAX AMOUNT"; TAB(62); "PPTRA DIS"; TAB(72); "TOTAL DUE"
        PRINT #RptFile, " "
        PRINT #RptFile, "Vehicle Listing Continued ..."
        PRINT #RptFile, ""
        PRINT #RptFile, ""
        PRINT #RptFile, " PPTRA Vehicle Information"
        PRINT #RptFile, ""
        PRINT #RptFile, ""
        CarCount = 0
    END IF
       WhatPers& = PersRec(1).NextRec
    LOOP WHILE WhatPers& > 0

   ' Finish the bill up here
     IF CarCount <= 6 THEN
      FOR LCnt = CarCount TO 6: PRINT #RptFile, "": NEXT LCnt
     END IF
  PRINT #RptFile, TAB(48); "Total Tax Due ";
  PRINT #RptFile, USING "$$#####,#.##"; TaxBill(1).TotalBillDue
  PRINT #RptFile, TAB(48); "Tax Due Date: "; DueDate$
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(48); "10% Penalty After 01/06/"; QPTrim$(STR$(WhatYear + 1)); "."
  PRINT #RptFile, TAB(48); "10% Interest Accrues Annually"
  PRINT #RptFile, '
  PRINT #RptFile, '
  PRINT #RptFile, "~ BN"; USING "#####"; PrnCnt































