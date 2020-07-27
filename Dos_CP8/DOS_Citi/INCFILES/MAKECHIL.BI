 'Determine Perc Discnt
  DoneCnt = DoneCnt + 1
  IF WhatYear = 1999 THEN PERC! = 27.5
  IF WhatYear = 2000 THEN PERC! = 47.5
  IF WhatYear >= 2001 THEN PERC! = 70
  'IF WhatYear = 2002 THEN PERC! = 100

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
  PRINT #RptFile, " "
  PRINT #RptFile, TAB(5); "Acct # "; TaxBill(1).CustRec
  PRINT #RptFile, TAB(5); CustName$
  PRINT #RptFile, TAB(5); TaxBill(1).CustAdd1
  PRINT #RptFile, TAB(5); TaxBill(1).CustAdd2
  PRINT #RptFile, TAB(5); QPTRIM$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  PRINT #RptFile, " "
  PRINT #RptFile, TAB(37); "PROP VALUE"; TAB(50); "TAX AMOUNT"; TAB(62); "PPTRA DIS"; TAB(72); "TOTAL DUE"
  PRINT #RptFile, " "

 'Line 24 Starts Here
  PRINT #RptFile, "Personal Property"; TAB(32); USING "#.##"; PersTaxRate#;
   PRINT #RptFile, TAB(37); USING "#######,#"; TaxBill(1).PersValue;
   PRINT #RptFile, TAB(50); USING "#####,#.##"; TaxBill(1).PersTaxDue + TaxBill(1).ChillHowieFudge;
   PRINT #RptFile, TAB(62); USING "####,#.##"; TaxBill(1).PPTRADiscnt;
   PRINT #RptFile, TAB(71); USING "#####,#.##"; (TaxBill(1).PersTaxDue + TaxBill(1).ChillHowieFudge - TaxBill(1).PPTRADiscnt)
'  PRINT #RptFile, "Machinery/Tools"; TAB(32); USING "#.##"; MTTaxRate#;
'   PRINT #RptFile, TAB(37); USING "#######,#"; TaxBill(1).MTValue;
'   PRINT #RptFile, TAB(50); USING "#####,#.##"; TaxBill(1).MTTaxDue;
'   PRINT #RptFile, TAB(71); USING "#####,#.##"; TaxBill(1).MTTaxDue
'  PRINT #RptFile, "Farm Equipment";
'   PRINT #RptFile, TAB(32); USING "#.##"; FETaxRate#;
'   PRINT #RptFile, TAB(37); USING "#######,#"; TaxBill(1).FEValue;
'   PRINT #RptFile, TAB(50); USING "#####,#.##"; TaxBill(1).FETaxDue;
'   PRINT #RptFile, TAB(71); USING "#####,#.##"; TaxBill(1).FETaxDue
'  PRINT #RptFile, "Mobile Homes";
'   PRINT #RptFile, TAB(32); USING "#.##"; MHTaxRate#;
'   PRINT #RptFile, TAB(37); USING "#######,#"; TaxBill(1).MHValue;
'   PRINT #RptFile, TAB(50); USING "#####,#.##"; TaxBill(1).MHTaxDue;
'   PRINT #RptFile, TAB(71); USING "#####,#.##"; TaxBill(1).MHTaxDue
'  PRINT #RptFile, "Merchant Capital";
'   PRINT #RptFile, TAB(32); USING "#.##"; MCTaxRate#;
'   PRINT #RptFile, TAB(37); USING "#######,#"; TaxBill(1).MCValue;
'   PRINT #RptFile, TAB(50); USING "#####,#.##"; TaxBill(1).MCTaxDue;
'   PRINT #RptFile, TAB(71); USING "#####,#.##"; TaxBill(1).MCTaxDue
  PRINT #RptFile, " PPTRA Vehicle Information"
 'Line 30 to 35 here to print vehicles
   CarCount = 0
   WhatPers& = TaxCust(1).FirstPersRec
    DO
     GET PersFile, WhatPers&, PersRec(1)

    PYear$ = RIGHT$(PersRec(1).Desc5, 4)
    PYear = VAL(PYear$)
    IF PYear > 0 AND PYear <> WhatYear THEN
    'Do Not Process This Record
    ELSE


   IF LEFT$(PersRec(1).Desc5, 1) = "Y" THEN
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

      VehDesc$ = " VIN# " + QPTRIM$(PersRec(1).Desc1)
      VehDesc$ = QPTRIM$(VehDesc$)
      TaxAmt# = (PersTaxRate# / 100) * PersRec(1).PERSVAL
      PRINT #RptFile, "*" + VehDesc$;
      PRINT #RptFile, TAB(37); USING "#######,#"; PersRec(1).PERSVAL;
      PRINT #RptFile, TAB(50); USING "#####,#.##"; TaxAmt#;
      PRINT #RptFile, TAB(62); USING "####,#.##"; PPTRADiscount#
      CarCount = CarCount + 1
    END IF
    
    END IF
     IF CarCount >= 6 THEN
        PRINT #RptFile, ""
        PRINT #RptFile, TAB(48); "Total Tax Due ";
         PRINT #RptFile, USING "$$#####,#.##"; TaxBill(1).TotalBillDue
        PRINT #RptFile, TAB(48); "Tax Due Date: "; DueDate$
        PRINT #RptFile, ""
        PRINT #RptFile,
        PRINT #RptFile,
        PRINT #RptFile,
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
        PRINT #RptFile, TAB(5); QPTRIM$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
        FOR LC = 18 TO 21
         PRINT #RptFile, " "
        NEXT LC
        PRINT #RptFile, TAB(37); "PROP VALUE"; TAB(50); "TAX AMOUNT"; TAB(62); "PPTRA DIS."; TAB(73); "TOTAL DUE"
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
     IF CarCount < 6 THEN
      FOR LCnt = CarCount TO 6: PRINT #RptFile, "": NEXT LCnt
     END IF
  PRINT #RptFile, TAB(48); "Total Tax Due ";
  PRINT #RptFile, USING "$$#####,#.##"; TaxBill(1).TotalBillDue
  PRINT #RptFile, TAB(48); "Tax Due Date: "; DueDate$;
  PRINT #RptFile,
  PRINT #RptFile, "BN"; USING "#####"; PrnCnt
  PRINT #RptFile, "~"
  IF DoneCnt >= 3 THEN
    DoneCnt = 0
    PRINT #RptFile, CHR$(12)
  END IF






