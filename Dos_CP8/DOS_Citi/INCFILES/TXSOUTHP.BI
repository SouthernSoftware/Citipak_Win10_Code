  'TaxBillRec (1)
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

  CustName$ = QPTrim$(TaxCust(1).FNAME) + " " + QPTrim$(TaxCust(1).LName)
  CustName$ = QPTrim$(CustName$)


'  TaxRate# = TaxBill(1).RealTaxRate
'  IF TaxRate# = 0 THEN
    TaxRate# = TaxBill(1).PersTaxRate
'  END IF

  'TaxBill(1).PersTaxDue

  PRINT #RptFile, QPTrim$(STR$(TaxBill(1).TaxYear)); "~";
  PRINT #RptFile, QPTrim$(STR$(TaxBill(1).BillNumber)); "~";
  PRINT #RptFile, QPTrim$(CustName$); "~";
  PRINT #RptFile, QPTrim$(TaxBill(1).CustAdd1); "~";
  PRINT #RptFile, QPTrim$(TaxBill(1).CustAdd2); "~";
  PRINT #RptFile, QPTrim$(TaxBill(1).CustAdd3); "~";
  PRINT #RptFile, QPTrim$(TaxBill(1).CustZip); "~";
  'PRINT #RptFile, QPTrim$(PINTemp(1)); "~";
  Desc$ = QPTrim$(TaxBill(1).RDesc1) + " " + QPTrim$(TaxBill(1).RDesc2)
  PRINT #RptFile, QPTrim$(Desc$); "~";
  PRINT #RptFile, QPTrim$(FUsing$(STR$(PersTaxRate#), "##.##")); "~";
  PRINT #RptFile, QPTrim$(FUsing$(STR$(MHTaxRate#), "##.##")); "~";
  PRINT #RptFile, QPTrim$(FUsing$(STR$(MCTaxRate#), "##.##")); "~";
  PRINT #RptFile, QPTrim$(FUsing$(STR$(FETaxRate#), "##.##")); "~";
  PRINT #RptFile, QPTrim$(FUsing$(STR$(MTTaxRate#), "##.##")); "~";

  PRINT #RptFile, QPTrim$(STR$(TaxBill(1).PersValue)); "~";
  PRINT #RptFile, QPTrim$(STR$(TaxBill(1).MHValue)); "~";
  PRINT #RptFile, QPTrim$(STR$(TaxBill(1).MCValue)); "~";
  PRINT #RptFile, QPTrim$(STR$(TaxBill(1).FEValue)); "~";
  PRINT #RptFile, QPTrim$(STR$(TaxBill(1).MTValue)); "~";

  PRINT #RptFile, QPTrim$(FUsing$(STR$(TaxBill(1).PersTaxDue), "######.##")); "~";
  PRINT #RptFile, QPTrim$(FUsing$(STR$(TaxBill(1).MHTaxDue), "######.##")); "~";
  PRINT #RptFile, QPTrim$(FUsing$(STR$(TaxBill(1).MCTaxDue), "######.##")); "~";
  PRINT #RptFile, QPTrim$(FUsing$(STR$(TaxBill(1).FETaxDue), "######.##")); "~";
  PRINT #RptFile, QPTrim$(FUsing$(STR$(TaxBill(1).MTTaxDue), "######.##")); "~";
  PRINT #RptFile, QPTrim$(FUsing$(STR$(TaxBill(1).ExptValue), "######.##")); "~";
  PRINT #RptFile, QPTrim$(FUsing$(STR$(TaxBill(1).PPTRADiscnt), "######.##")); "~";
  PRINT #RptFile, QPTrim$(FUsing$(STR$(TaxBill(1).TotalBillDue), "######.##")); "~";

  PRINT #RptFile, DueDate$; "~VEH_START~";

  CarCount = 0
  WhatPers& = TaxCust(1).FirstPersRec
  DO
    GET PersFile, WhatPers&, PersRec(1)
    PYear$ = RIGHT$(PersRec(1).Desc5, 4)
    PYear = VAL(PYear$)

    IF PYear > 0 AND PYear <> WhatYear THEN
    
    ELSE
      IF LEFT$(PersRec(1).Desc5, 1) = "Y" THEN
        IF Round#(PersRec(1).PERSVAL) > 20000 THEN
          PPTRAVal# = 20000
        ELSE
          PPTRAVal# = Round#(PersRec(1).PERSVAL)
        END IF
        IF PPTRAVal# <= 1000 THEN
          PPTRADiscount# = Round#((PPTRAVal# / 100) * TaxRate97!)
        ELSE
          PPTRADiscount# = Round#(((PPTRAVal# / 100) * (PERC! / 100)) * TaxRate97!)
        END IF

        VehDesc$ = "VIN:" + QPTrim$(PersRec(1).Desc1)
        VehDesc$ = QPTrim$(VehDesc$)
        TaxAmt# = Round#((PersTaxRate# / 100) * PersRec(1).PERSVAL)

        PRINT #RptFile, VehDesc$; "~";
        PRINT #RptFile, QPTrim$(FUsing$(STR$(PersRec(1).PERSVAL), "########")); "~";
        PRINT #RptFile, QPTrim$(FUsing$(STR$(TaxAmt#), "######.##")); "~";
        PRINT #RptFile, QPTrim$(FUsing$(STR$(PPTRADiscount#), "######.##")); "~";
        CarCount = CarCount + 1
      END IF
    END IF
    WhatPers& = PersRec(1).NextRec
  LOOP WHILE WhatPers& > 0
  PRINT #RptFile, "~VEH_END~"

  'NetTaxVal# = Round#(TaxBill(1).PersValue)
  'NetTaxVal# = Round#(NetTaxVal# - TaxBill(1).ExptValue)

  'PRINT #RptFile, QPTrim$(STR$(NetTaxVal#)); "~";
  'PRINT #RptFile, QPTrim$(STR$(PersTaxRate#));

