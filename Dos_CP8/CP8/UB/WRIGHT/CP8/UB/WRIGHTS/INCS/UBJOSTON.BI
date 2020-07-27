'Johnston County 08-07-97

  Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)
  
'  CDeposit# = UBCustRec(1).DepositAmt
  IF FinalFlag AND CDeposit# THEN
    TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
  END IF
  
  NPrintCnt = 0
  Did2Rev = False

'*** Look for meter readings
  PrevRead& = 0
  CurrRead& = 0
  UsageAmt& = 0
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
      PrevRead& = UBBillRec(1).PrevRead(WMtrCnt)
      CurrRead& = UBBillRec(1).CurRead(WMtrCnt)
      UsageAmt& = CurrRead& - PrevRead&
      IF UsageAmt& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WMtrCnt))) - 1)
        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WMtrCnt)) + UBBillRec(1).CurRead(WMtrCnt)
      END IF
      EXIT FOR
    END SELECT
  NEXT
'**** Find a meter

  PRINT #UBRpt, "~"
  PRINT #UBRpt, TAB(2); CustAcctNo&
  PRINT #UBRpt,

  IF UBBillRec(1).RevAmt(1) > 0 THEN
    PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(1).RevName, 10); TAB(23); USING "######.##"; UBBillRec(1).RevAmt(1);
    PRINT #UBRpt, TAB(35); USING "#####"; PrintedCnt
  ELSE
    PRINT #UBRpt, TAB(35); USING "#####"; PrintedCnt
  END IF

  IF UBBillRec(1).RevAmt(2) > 0 THEN
    Did2Rev = True
    PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(2).RevName, 10); TAB(23); USING "######.##"; UBBillRec(1).RevAmt(2)
  ELSE
    PRINT #UBRpt,
  END IF

  IF UBBillRec(1).RevAmt(3) > 0 THEN
    PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(3).RevName, 10); TAB(23); USING "######.##"; UBBillRec(1).RevAmt(3)
  ELSE
    PRINT #UBRpt,
  END IF
  
  IF UBBillRec(1).RevAmt(4) > 0 THEN
    PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(4).RevName, 10); TAB(23); USING "######.##"; UBBillRec(1).RevAmt(4);
  END IF
  PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).SERVADDR, 27)

  IF Previous# <> 0 THEN
    PRINT #UBRpt, TAB(4); "     Previous:"; TAB(23); USING "######.##"; Previous#
  ELSE
    PRINT #UBRpt,
  END IF

'deposit
  IF CDeposit# > 0 THEN
    PRINT #UBRpt, TAB(4); " Less Deposit:"; TAB(23); USING "######.##"; CDeposit#;
  END IF

  PRINT #UBRpt, TAB(33); LEFT$(BillDate$, 6); RIGHT$(BillDate$, 2); TAB(42); LEFT$(PrevDate$, 6); RIGHT$(PrevDate$, 2); TAB(51); LEFT$(DateRead$, 6); RIGHT$(DateRead$, 2)
  PRINT #UBRpt,

  PRINT #UBRpt, TAB(23); USING "######.##"; TotalAmt#;
  PRINT #UBRpt, TAB(33); USING "########"; PrevRead&; TAB(42); CurrRead&; TAB(51); UsageAmt&
  PRINT #UBRpt,
  PRINT #UBRpt,

  PRINT #UBRpt, TAB(12); CustAcctNo&; TAB(23); USING "######.##"; TotalAmt#;
  PRINT #UBRpt, TAB(33); LEFT$(BillDate$, 6); RIGHT$(BillDate$, 2);
  PRINT #UBRpt, TAB(42); LEFT$(PastDueDate$, 6); RIGHT$(PastDueDate$, 2);
  PRINT #UBRpt, TAB(52); LEFT$(Message$, 8)
  PRINT #UBRpt,

  PRINT #UBRpt, TAB(8); UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; TAB(33); LEFT$(UBCustRec(1).CustName, 27)
  PRINT #UBRpt, TAB(33); LEFT$(UBCustRec(1).ADDR1, 27)
  PRINT #UBRpt, TAB(33); LEFT$(UBCustRec(1).ADDR2, 27)
  PRINT #UBRpt, TAB(33); LEFT$(UBCustRec(1).CITY, 14); " ";
  PRINT #UBRpt, UBCustRec(1).STATE; " "; UBCustRec(1).ZIPCODE
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(7); Msg2$
  PRINT #UBRpt, TAB(7); Msg3$
  PRINT #UBRpt, TAB(7); Msg4$
  PRINT #UBRpt, "~"             ' "LAST LINE"


