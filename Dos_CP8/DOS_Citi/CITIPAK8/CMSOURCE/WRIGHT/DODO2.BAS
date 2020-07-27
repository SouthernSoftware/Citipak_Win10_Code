  REDIM TransList(1 TO TotalRecs&)    AS FLen2
  REDIM Picked(1 TO TotalRecs&) AS INTEGER

  WhatCnt& = 1

  FOpenS PayFileName$, PHandle
  NumOfRecs& = FLof(PHandle) \ UBPayRecLen
  IF NumOfRecs& > 0 THEN
    FOR Cnt& = 1 TO NumOfRecs&
      FGetRTA PHandle, UBPaymentRec(1), Cnt&, UBPayRecLen
      LSET TransList(WhatCnt&).V = UBPaymentRec(1).CUSTNAME
      MID$(TransList(WhatCnt&).V, 28) = "Payment"
      MID$(TransList(WhatCnt&).V, 44) = FUsing$(STR$(UBPaymentRec(1).AMTPAID), Fmt$)
      WhatCnt& = WhatCnt& + 1
    NEXT
  END IF
  FClose PHandle
  PayListCnt& = NumOfRecs&

  MaxLen = 64
  Col = ((80 - MaxLen) \ 2) - 2

  BlockClear

