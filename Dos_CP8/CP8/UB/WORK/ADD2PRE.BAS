'          '02-20-97 Add revenue totals by rate code
'          IF WhatRate > 0 THEN
'            RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
'          END IF
'          ^^^^ Old way.
'040804 Corrects problem with accounts that have both flat & ratecoded
'       services to the same revenue, skewing the revenue by ratecode totals.
          IF WhatRate > 0 THEN
            RateTotals(WhatRate) = Round#(RateTotals(WhatRate) + UBBillRec(1).RevAmt(TRevCnt))
            FOR FRCnt = 1 TO 4
              WhatService = UBCustRec(1).FlatRates(FRCnt).REVSRC
              IF UBCustRec(1).FlatRates(FRCnt).REVSRC > 0 THEN
                IF WhatService = TRevCnt THEN
                  IF UBCustRec(1).FlatRates(FRCnt).FRAMT <> 0 THEN
                    Multi = UBCustRec(1).FlatRates(FRCnt).NumMin
                    IF Multi < 1 THEN
                      Multi = 1
                    END IF
                    FlatAmt# = Round#(UBCustRec(1).FlatRates(FRCnt).FRAMT * Multi)
                    IF FlatAmt# <> 0 THEN
                      RateTotals(WhatRate) = Round#(RateTotals(WhatRate) - FlatAmt#)
                    END IF
                  END IF
                END IF
              END IF
            NEXT
          END IF
'***********************************************************************
