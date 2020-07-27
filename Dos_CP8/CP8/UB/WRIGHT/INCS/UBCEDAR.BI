'Cedar Bluff VA
'last revised 09-09-98

  CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)

  TotalTax# = 0
  FOR TaxCnt = 1 TO MaxRevsCnt
    TotalTax# = Round(TotalTax# + UBBillRec(1).TaxAmt(TaxCnt))
  NEXT

  Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

  NPrintCnt = 0

  PRINT #UBRpt, "~"; TAB(47); USING "########"; PrintedCnt
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, USING "#######"; CustAcctNo&;
  PRINT #UBRpt, TAB(19); DateRead$
  PRINT #UBRpt,
  PRINT #UBRpt, 'TAB(35); USING "Acct: #####"; CustAcctNo&
  
  DidWMrt = False
  FOR WMtrCnt = 1 TO 7
    SELECT CASE UBBillRec(1).MtrTypes(WMtrCnt)
    CASE MtrWaterOnly, MtrSewerOnly, MtrCombined, MtrTouchRead
      IF UBBillRec(1).PrevRead(WMtrCnt) < 0 THEN
        UBBillRec(1).PrevRead(WMtrCnt) = 0
      END IF
      IF UBBillRec(1).CurRead(WMtrCnt) < 0 THEN
        UBBillRec(1).CurRead(WMtrCnt) = 0
      END IF
      PRINT #UBRpt, USING "#########"; UBBillRec(1).PrevRead(WMtrCnt);
      PRINT #UBRpt, TAB(11); USING "#########"; UBBillRec(1).CurRead(WMtrCnt);
      UsageAmt& = UBBillRec(1).CurRead(WMtrCnt) - UBBillRec(1).PrevRead(WMtrCnt)
      IF UsageAmt& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WMtrCnt))) - 1)
        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WMtrCnt)) + UBBillRec(1).CurRead(WMtrCnt)
      END IF
      PRINT #UBRpt, TAB(20); USING "#########"; UsageAmt&
      DidWMrt = True
      EXIT FOR
    END SELECT
  NEXT
  
  IF NOT DidWMrt THEN
    PRINT #UBRpt,
  END IF
  PRINT #UBRpt, TAB(35); USING "Acct: #####"; CustAcctNo&
  IF UBBillRec(1).RevAmt(1) > 0 THEN
    PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(1).RevName, 10); TAB(21); USING "#######.##"; UBBillRec(1).RevAmt(1);
  END IF
  PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).CustName, 27)

  IF UBBillRec(1).RevAmt(2) > 0 THEN
    PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(2).RevName, 10); TAB(21); USING "#######.##"; UBBillRec(1).RevAmt(2);
  END IF
  PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).ADDR1, 27)

  IF UBBillRec(1).RevAmt(3) > 0 THEN
    PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(3).RevName, 10); TAB(21); USING "#######.##"; UBBillRec(1).RevAmt(3);
  END IF
  PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).CITY, 14); " "; UBCustRec(1).STATE; " "; UBCustRec(1).ZIPCODE

  IF UBBillRec(1).RevAmt(4) > 0 THEN
    PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(4).RevName, 10); TAB(21); USING "#######.##"; UBBillRec(1).RevAmt(4)
  ELSE
    PRINT #UBRpt,
  END IF
  
  IF Previous# <> 0 THEN
    PRINT #UBRpt, TAB(2); "Previous:"; TAB(21); USING "#######.##"; Previous#
  ELSE
    PRINT #UBRpt,
  END IF
  
  IF FinalFlag AND CDeposit# THEN
    PRINT #UBRpt, "Deposit:"; TAB(21); USING "#######.##"; -UBCustRec(1).DepositAmt
    TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
  ELSE
    PRINT #UBRpt,
  END IF

  IF TotalTax# <> 0 THEN
    PRINT #UBRpt, TAB(2); "  Tax"; TAB(21); USING "#####.##"; TotalTax#;
  END IF

  PRINT #UBRpt, TAB(35); UBCustRec(1).ServAddr
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(21); USING "#######.##"; TotalAmt#;
  PRINT #UBRpt, TAB(35); PastDueDate$;
  PRINT #UBRpt, TAB(49); USING "#####.##"; TotalAmt#
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(2); Message$
  PRINT #UBRpt, TAB(2); Msg2$
  PRINT #UBRpt, TAB(2); Msg3$
  PRINT #UBRpt, TAB(2); Msg4$
  PRINT #UBRpt, "~"

'  PRINT #UBRpt, TAB(2); CustMsg$
  
  
  

