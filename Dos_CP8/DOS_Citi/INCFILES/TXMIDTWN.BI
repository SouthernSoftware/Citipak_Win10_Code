
  'Determine Perc Discnt
  IF WhatYear = 1999 THEN PERC! = 27.5
  IF WhatYear = 2000 THEN PERC! = 47.5
  IF WhatYear >= 2001 THEN PERC! = 70
'  IF WhatYear = 2002 THEN PERC! = 100

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
  PRINT #RptFile, TAB(30); "TOWN OF MIDDLETOWN"
  PRINT #RptFile, TAB(30); "   PO BOX 696"
  PRINT #RptFile, TAB(29); "MIDDLETOWN, VA 22645"
  PRINT #RptFile, TAB(30); "  (540-869-2226)"
  PRINT #RptFile, TAB(27); "PERSONAL PROPERTY TAX BILL"
  FOR LC = 6 TO 8
    PRINT #RptFile, " "
  NEXT
  PRINT #RptFile, TAB(10); "ACCT # "; TaxBill(1).CustRec;
  PRINT #RptFile, TAB(65); "BILL #"; USING "######"; TaxBill(1).BillNumber
  PRINT #RptFile, TAB(10); LEFT$(CustName$, 25);
  PRINT #RptFile, TAB(63); "TAX YEAR "; WhatYear
  PRINT #RptFile, TAB(10); LEFT$(TaxBill(1).CustAdd1, 25);
  PRINT #RptFile, TAB(63); "TAX RATE "; USING "#.###"; PersTaxRate#
  PRINT #RptFile, TAB(10); LEFT$(TaxBill(1).CustAdd2, 25)
  PRINT #RptFile, TAB(10); QPTRIM$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  FOR LC = 14 TO 17
    PRINT #RptFile, " "
  NEXT
  PRINT #RptFile, TAB(40); "PROPERTY"; TAB(51); " TAX"; TAB(61); " PPTRA"
  PRINT #RptFile, TAB(41); "VALUE"; TAB(51); "AMOUNT"; TAB(61); "DISCOUNT"; TAB(71); "TOTAL DUE"
  'Line 23 Starts Here
  PRINT #RptFile, TAB(2); "Personal Property";
  PRINT #RptFile, TAB(38); USING "######,#"; TaxBill(1).PersValue;
  PRINT #RptFile, TAB(48); USING "#####,#.##"; (TaxBill(1).PersTaxDue);
  PRINT #RptFile, TAB(60); USING "####.##"; TaxBill(1).PPTRADiscnt;
  PRINT #RptFile, TAB(70); USING "#####,#.##"; TaxBill(1).TotalBillDue

  CarCount = 0
  PRINT #RptFile,
  PRINT #RptFile, " PPTRA Information"

  WhatPers& = TaxCust(1).FirstPersRec
  DO
    GET PersFile, WhatPers&, PersRec(1)
    IF LEFT$(PersRec(1).DESC5, 1) = "Y" THEN
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
    PRINT #RptFile, " The Personal Property Relief Act provides that the tax on the first"
    PRINT #RptFile, " $20,000 of value of your personal car, motorcycle, and pickup or panel"
    PRINT #RptFile, " truck under 7,501 pounds, which is a qualifying vehicle, has been"
    PRINT #RptFile, " reduced by 70% this year. If your qualifying vehicle's value is"
    PRINT #RptFile, " $1,000 or less, your tax has been eliminated. These reductions are"
    PRINT #RptFile, " based on the local tax rates in effect on July 1 or August 1, 1997,"
    PRINT #RptFile, " whichever was higher."
    PRINT #RptFile, " Please contact the Town Office at the phone number above with any"
    PRINT #RptFile, " questions."
  END IF

  PRINT #RptFile,
  PRINT #RptFile, " BN"; USING "#####"; PrnCnt
  PRINT #RptFile, CHR$(12);

