              IF ReadAmt& > 0 THEN
                IF CubMtr THEN
                  ReadAmt& = ReadAmt& * 7.481
                END IF
                RateConsump(WhatRate) = RateConsump(WhatRate) + (ReadAmt& * MeterMulti&)
                BookConsump(WhatBook).Consump(TRevCnt) = BookConsump(WhatBook).Consump(TRevCnt) + (ReadAmt& * MeterMulti&)
                ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + (ReadAmt& * MeterMulti&)
              ELSEIF ReadAmt& < 0 THEN
                'Meter has rolled over or, been misread
                MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(MCCnt))) - 1)
                ReadAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MCCnt)) + UBBillRec(1).CurRead(MCCnt)
                IF CubMtr THEN
                  ReadAmt& = ReadAmt& * 7.481
                END IF
                RateConsump(WhatRate) = RateConsump(WhatRate) + (ReadAmt& * MeterMulti&)
                BookConsump(WhatBook).Consump(TRevCnt) = BookConsump(WhatBook).Consump(TRevCnt) + (ReadAmt& * MeterMulti&)
                ConsumpTot(TRevCnt, 1) = ConsumpTot(TRevCnt, 1) + (ReadAmt& * MeterMulti&)
              END IF

