  
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
  
  CustName$ = RTRIM$(TaxCust(1).FNAME) + " " + RTRIM$(TaxCust(1).LName)
  CustName$ = LTRIM$(CustName$)

'***************************************
  BusFurn = 0
  WhatPers& = TaxCust(1).FirstPersRec
  DO
    GET PersFile, WhatPers&, PersRec(1)
    IF LEFT$(PersRec(1).DESC5, 1) = "Y" THEN
      IF LEFT$(QPTrim$(PersRec(1).Desc1), 2) = "BF" THEN
        BusFurn = -1
        EXIT DO
      END IF
    END IF
    WhatPers& = PersRec(1).NextRec
  LOOP WHILE WhatPers& > 0
'***************************************

  'Must Calc Late Fee Here
  PenAmt# = round#((TaxBill(1).TotalBillDue + TaxBill(1).PPTRADiscnt) * .05)
  PRINT #RptFile, TAB(4); "~"
  PRINT #RptFile, " "
  PRINT #RptFile, TAB(4); "SSN: "; TaxCust(1).CSSN; TAB(75); USING "#####"; TaxBill(1).BillNumber
  PRINT #RptFile, " "
  PRINT #RptFile, " "
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(4); USING "#######,#"; TaxBill(1).PersValue + TaxBill(1).MTValue;
  PRINT #RptFile, TAB(18); USING "#####.##"; TaxBill(1).TotalBillDue + TaxBill(1).PPTRADiscnt;
  PRINT #RptFile, TAB(27); USING "####.##"; TaxBill(1).PPTRADiscnt;
  PRINT #RptFile, TAB(36); USING "#####.##"; TaxBill(1).TotalBillDue;
  PRINT #RptFile, TAB(47); USING "####.##"; PenAmt#;
  PRINT #RptFile, TAB(56); USING "#####,#.##"; TaxBill(1).TotalBillDue + PenAmt#;
  IF TaxBill(1).MTValue > 0 THEN
    PRINT #RptFile, TAB(69); "(Mach/Tools)"
    PRINT #RptFile, ""
  ELSE
    PRINT #RptFile, TAB(69); "(Vehicle)"
    IF BusFurn THEN
      PRINT #RptFile, TAB(69); "(Bus Furn)"
    ELSE
      PRINT #RptFile, ""
    END IF
  END IF

  
  CarCount = 0
  WhatPers& = TaxCust(1).FirstPersRec
  DO
    GET PersFile, WhatPers&, PersRec(1)
    
    IF LEFT$(PersRec(1).DESC5, 1) = "Y" THEN
      IF round#(PersRec(1).PERSVAL) > 20000 THEN
        PPTRAVal# = 20000
      ELSE
        PPTRAVal# = round#(PersRec(1).PERSVAL)
      END IF
      IF PPTRAVal# <= 1000 THEN
        PPTRADiscount# = round#((round#(PPTRAVal# / 100) * PersTaxRate#))
      ELSE
        PPTRADiscount# = round#((round#((PPTRAVal# / 100) * (PERC! / 100)) * PersTaxRate#))
      END IF
      IF LEFT$(QPTrim$(PersRec(1).Desc1), 2) = "BF" THEN
        BusFurn = -1
      END IF
      VehDesc$ = LEFT$(PersRec(1).Desc4, 4) + " " + LEFT$(PersRec(1).Desc2, 7) + " "
      VehDesc$ = VehDesc$ + RTRIM$(PersRec(1).Desc1)
      VehDesc$ = QPTrim$(VehDesc$)
      
      TaxAmt# = (PersTaxRate# / 100) * PersRec(1).PERSVAL
      PRINT #RptFile, TAB(4); "*" + LEFT$(VehDesc$, 31);
      PRINT #RptFile, TAB(38); USING "######,#"; PersRec(1).PERSVAL;
      PRINT #RptFile, TAB(48); USING "#####.##"; TaxAmt#;
      PRINT #RptFile, TAB(57); USING "####,#.##"; PPTRADiscount#
      CarCount = CarCount + 1
    END IF
    WhatPers& = PersRec(1).NextRec
  LOOP WHILE WhatPers& > 0
  ' Finish the bill up here
  IF CarCount < 5 THEN
    FOR LCnt = CarCount TO 4
      PRINT #RptFile, ""
    NEXT LCnt
  END IF
  IF CarCount > 0 THEN
    PRINT #RptFile, TAB(4); "* PPTRA Qualifying Vehicle"
  ELSE
    PRINT #RptFile,
  END IF
  PRINT #RptFile, TAB(19); "Acct # "; TaxBill(1).CustRec
  PRINT #RptFile,
  PRINT #RptFile, TAB(19); CustName$
  PRINT #RptFile, TAB(19); TaxBill(1).CustAdd1
  PRINT #RptFile, TAB(19); TaxBill(1).CustAdd2
  PRINT #RptFile, TAB(19); QPTrim$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
  PRINT #RptFile,
  PRINT #RptFile, TAB(4); "BN"; USING "#####"; PrnCnt
  

