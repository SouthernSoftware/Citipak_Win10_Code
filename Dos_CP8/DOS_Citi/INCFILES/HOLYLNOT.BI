'Holly Springs

  'NoticeInfo.MsgLine1
  CODate$ = MakeLongDate$(QPTrim$(NoticeInfo.MsgLine1))
  CmaPos = INSTR(CODate$, ",")
  IF CmaPos > 0 THEN
    WhatDay$ = LEFT$(CODate$, CmaPos - 1)
  END IF

  PRINT #UBRpt, "~"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(16); LongNDate$
  PRINT #UBRpt, TAB(16); "Town of Holly Springs"
  PRINT #UBRpt, TAB(16); "P.O. Box 8"
  PRINT #UBRpt, TAB(16); "Holly Springs, N.C. 27540"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(12); "Acct: "; Acct$
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(14); "We have not received payment on your water bill account which"
  PRINT #UBRpt, TAB(12); "results in your account being past due. We feel this is an oversight."
  PRINT #UBRpt, TAB(12); "If payment has already been made please disregard this notice. If"
  PRINT #UBRpt, TAB(12); "not, please remit payment by 5:00 PM on "; LongPDate$;
  PRINT #UBRpt, TAB(12); "to avoid a penalty charge of $25.00 and disconnection of service"
  PRINT #UBRpt, TAB(12); "at 8:00 AM. on "; CODate$
  PRINT #UBRpt, TAB(14); "If water is cut off for non-payment, service will only be restored"
  PRINT #UBRpt, TAB(12); "upon total payment of water bill, past and current amounts due, plus"
  PRINT #UBRpt, TAB(12); "the $25.00 penalty."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(14); "Total amount due as of "; LongNDate$; " "; USING "$$#####.##"; TotalBal#;
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(14); "Phone: 557-3920"
  PRINT #UBRpt, TAB(50); LEFT$(UBCustRec(1).CustName, 25)
  PRINT #UBRpt, TAB(50); LEFT$(UBCustRec(1).ADDR1, 25)
  PRINT #UBRpt, TAB(50); LEFT$(UBCustRec(1).ADDR2, 25)
  PRINT #UBRpt, TAB(50); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, "~"

'****************************************
'  CODate$ = MakeLongDate$(Num2Date(NoticeInfo.PayByDate + 1))
'  CmaPos = INSTR(CODate$, ",")
'  IF CmaPos > 0 THEN
'    WhatDay$ = LEFT$(CODate$, CmaPos - 1)
'  END IF
'  PRINT #UBRpt, "~"
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt, TAB(16); "Town of Fuquay-Varina"
'  PRINT #UBRpt, TAB(16); "1300 E. Academy St."
'  PRINT #UBRpt, TAB(16); "Fuqay-Varina, N.C. 27526"
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt, TAB(12); LongNDate$; "   WATER & SEWER PAYMENTS  552-1405"
'  PRINT #UBRpt,
'  PRINT #UBRpt, TAB(14); "WE HAVE NOT RECEIVED PAYMENT FOR YOUR WATER BILL AND FEEL"
'  PRINT #UBRpt, TAB(12); "THERE MUST BE AN OVERSIGHT. PLEASE CHECK ON THIS SINCE THE"
'  PRINT #UBRpt, TAB(12); "BELOW AMOUNT MUST BE PAID BY 5:00 PM "; LongPDate$
'  PRINT #UBRpt, TAB(12); "IN ORDER TO AVOID A PENALTY OF $10.00 WHICH WILL BE CHARGED"
'  PRINT #UBRpt, TAB(12); "AT 8:00 AM ON "; WhatDay$; " AND WATER SERVICE WILL BE DISCONNECTED"
'  PRINT #UBRpt, TAB(12); "ON "; CODate$; "."
'  PRINT #UBRpt, TAB(14); "**IF WATER IS CUT OFF FOR NON-PAYMENT, SERVICE WILL NOT BE"
'  PRINT #UBRpt, TAB(12); "RESTORED UNTIL BILL PLUS $10.00 PENALTY FEE HAS BEEN PAID IN"
'  PRINT #UBRpt, TAB(12); "FULL.  A NIGHT DROP BOX IS AVAILABLE FOR YOUR CONVENIENCE AT"
'  PRINT #UBRpt, TAB(12); "THE DRIVE THRU WINDOW.**"
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt, TAB(50); UBCustRec(1).CustName
'  PRINT #UBRpt, TAB(50); UBCustRec(1).ADDR1
'  PRINT #UBRpt, TAB(50); UBCustRec(1).ADDR2
'  PRINT #UBRpt, TAB(12); USING "Balance Due: $$#####.##"; TotalBal#;
'  PRINT #UBRpt, TAB(50); QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt, "~"
'
''****************************************

