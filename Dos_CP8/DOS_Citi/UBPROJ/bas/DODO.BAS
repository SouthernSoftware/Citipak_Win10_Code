          IF (TMeterConsp& - UBRateTbls(WhatTbl).MINUNITS) <= 0 THEN
            '062697 fix for min consump test to actual (NumUsers * MINUNITS)
'071201 Added fix for prorating
'pro dale
            ProRevAmt# = 0
            IF ProRateFlag AND ProrateServ(RCnt) THEN
              ProRevAmt# = Round#((AddRevAmt# + UBRateTbls(WhatTbl).MINAMT) * ProPct#)
              UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + ProRevAmt#)
              UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + ProRevAmt#)
            ELSE
              UBBillRec(1).RevAmt(RCnt) = Round#(UBBillRec(1).RevAmt(RCnt) + (AddRevAmt# + UBRateTbls(WhatTbl).MINAMT))
              UBBillRec(1).TransAmt = Round#(UBBillRec(1).TransAmt + (AddRevAmt# + UBRateTbls(WhatTbl).MINAMT))
            END IF
            GOTO GotAmt
          END IF

