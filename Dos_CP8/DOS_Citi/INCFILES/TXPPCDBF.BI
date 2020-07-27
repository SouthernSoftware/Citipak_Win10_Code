  'Determine Perc Discnt
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
   PRINT #RptFile, "~"; TAB(34); STR$(WhatYear); " PERSONAL PROPERTY"
   PRINT #RptFile, TAB(5); TownName$
   PRINT #RptFile, TAB(5); Add1$
   PRINT #RptFile, TAB(5); Add2$
   PRINT #RptFile, TAB(5); Add3$
   PRINT #RptFile, TAB(10); TaxCust(1).CSSN; TAB(65); USING "PP  ##.###"; PersTaxRate#
   PRINT #RptFile, " "
   PRINT #RptFile, " "

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
        ELSE
          PPTRADiscount# = 0
        END IF
        VehDesc$ = QPTRIM$(PersRec(1).Desc4) + " " + LEFT$(PersRec(1).Desc2, 22) + "   " + LEFT$(PersRec(1).Desc5, 1)
        VehDesc$ = QPTRIM$(VehDesc$)
        CarCount = CarCount + 1
        IF PersRec(1).PERSVAL <= 0 THEN
          TaxAmt# = (MHTaxRate# / 100) * PersRec(1).MHValue
          PRINT #RptFile, VehDesc$;
          PRINT #RptFile, TAB(33); USING "#######,#"; PersRec(1).MHValue;
          PRINT #RptFile, TAB(44); USING "#####,#.##"; TaxAmt#;
          PRINT #RptFile, TAB(54); USING "####,#.##"; PPTRADiscount#;
          PRINT #RptFile, TAB(64); USING "####,#.##"; TaxAmt#
        ELSE
          TaxAmt# = (PersTaxRate# / 100) * PersRec(1).PERSVAL
          PRINT #RptFile, VehDesc$;
          PRINT #RptFile, TAB(33); USING "#######,#"; PersRec(1).PERSVAL;
          PRINT #RptFile, TAB(44); USING "#####,#.##"; TaxAmt#;
          PRINT #RptFile, TAB(54); USING "####,#.##"; PPTRADiscount#;
          PRINT #RptFile, TAB(64); USING "####,#.##"; Round#(TaxAmt# - PPTRADiscount#)
        END IF
      END IF
      IF (CarCount >= 5) AND (PersRec(1).NextRec > 0) THEN
        PRINT #RptFile, " "
        PRINT #RptFile, ""
        PRINT #RptFile, TAB(10); USING "#####"; PrnCnt;
        PRINT #RptFile, TAB(36); DueDate$; TAB(66); "CONT'D"                 '; USING "$$#####,#.##"; TaxBill(1).TotalBillDue
        PRINT #RptFile,
        PRINT #RptFile, TAB(9); CustName$
        PRINT #RptFile, TAB(9); TaxBill(1).CustAdd1
        PRINT #RptFile, TAB(9); TaxBill(1).CustAdd2
        PRINT #RptFile, TAB(9); QPTRIM$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
        PRINT #RptFile,
        PRINT #RptFile,
        PRINT #RptFile, "~"
        PRINT #RptFile, "~"
        PRINT #RptFile, TAB(5); TownName$
        PRINT #RptFile, TAB(5); Add1$
        PRINT #RptFile, TAB(5); Add2$
        PRINT #RptFile, TAB(5); Add3$
        PRINT #RptFile, TAB(10); TaxCust(1).CSSN
        PRINT #RptFile, " "
        PRINT #RptFile, '"Vehicle Listing Continued ..."
        CarCount = 0
      ELSEIF (CarCount >= 5) AND (PersRec(1).NextRec <= 0) THEN
        PRINT #RptFile, ""
      '  CarCount = 0
      END IF
      WhatPers& = PersRec(1).NextRec
    LOOP WHILE WhatPers& > 0

  ' Finish the bill up here
    IF CarCount < 5 THEN
      FOR LCnt = CarCount TO 5
        PRINT #RptFile, ""
      NEXT
    END IF
'
    PRINT #RptFile, ""
    PRINT #RptFile, TAB(10); USING "#####"; PrnCnt;
    PRINT #RptFile, TAB(36); DueDate$; TAB(62); USING "$$#####,#.##"; TaxBill(1).TotalBillDue
    PRINT #RptFile,
    PRINT #RptFile, TAB(9); CustName$
    PRINT #RptFile, TAB(9); TaxBill(1).CustAdd1
    PRINT #RptFile, TAB(9); TaxBill(1).CustAdd2
    PRINT #RptFile, TAB(9); QPTRIM$(TaxBill(1).CustAdd3); " "; TaxBill(1).CustZip
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile, "~"



