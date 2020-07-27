'Hillsville VA Custom bill
'

    CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)
    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    END IF
    
    FoundAMtr = False
    FOR MChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(MChk) > 0 THEN
        FoundAMtr = True
        EXIT FOR
      END IF
    NEXT

    IF FoundAMtr = False THEN
      'if no metered services then adjust read dates to billdate
      'and billdate - 30
      DateRead$ = Num2Date$(UBBillRec(1).BillDate)
      PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
    END IF

'*** Look for meter readings
  PrevRead& = 0
  CurrRead& = 0
  UsageAmt& = 0
  DidWMrt = False
  Multi& = 0
  Zero$ = ""
  FOR WMtrCnt = 1 TO 7
    SELECT CASE UBBillRec(1).MtrTypes(WMtrCnt)
    CASE MtrWaterOnly, MtrSewerOnly, MtrCombined, MtrTouchRead
      Multi& = UBCustRec(1).LocMeters(WMtrCnt).MTRMulti
      IF Multi& < 1 THEN Multi& = 1
      IF UBBillRec(1).PrevRead(WMtrCnt) < 0 THEN
        UBBillRec(1).PrevRead(WMtrCnt) = 0
      END IF
      IF UBBillRec(1).CurRead(WMtrCnt) < 0 THEN
        UBBillRec(1).CurRead(WMtrCnt) = 0
      END IF
      PrevRead& = (UBBillRec(1).PrevRead(WMtrCnt) * Multi&)
      CurrRead& = UBBillRec(1).CurRead(WMtrCnt) * Multi&

      UsageAmt& = CurrRead& - PrevRead&
      IF UsageAmt& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WMtrCnt))) - 1)
        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WMtrCnt)) + UBBillRec(1).CurRead(WMtrCnt)
      END IF
      'UsageAmt& = UsageAmt& * Multi&
      EXIT FOR
    END SELECT
  NEXT
'**** Find a meter
    
    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF

    IF TotalAmt# > 0 THEN
      'TenPct# = 10
      TenPct# = Round#(TotalAmt# * .1)
    ELSE
      TenPct# = 0
    END IF

    AcctNum = UBBillRec(1).CustAcctNo
    Acct$ = QPTrim$(STR$(AcctNum))
    SELECT CASE AcctNum
    CASE IS < 10
      Acct$ = "00" + Acct$
    CASE IS < 100
      Acct$ = "0" + Acct$
    END SELECT
    AcctLen = LEN(Acct$)

    PRINT #UBRpt, TAB(50); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, USING "##########"; UBBillRec(1).CustAcctNo;
    PRINT #UBRpt, TAB(15); LEFT$(UBCustRec(1).ServAddr, 19); TAB(50); USING "########"; UBBillRec(1).CustAcctNo;
    PRINT #UBRpt, TAB(62); Num2Date$(UBBillRec(1).BillDate)
    PRINT #UBRpt,

    PRINT #UBRpt, TAB(50); Num2Date$(UBBillRec(1).PastDueDate); TAB(64); USING "#####.##"; TotalAmt#
    PRINT #UBRpt, TAB(3); Num2Date$(UBBillRec(1).BillDate); TAB(15); PrevDate$; TAB(26); DateRead$;
     'Only Print Days if Greater than 0
     IF DaysINRead > 0 THEN
       PRINT #UBRpt, TAB(40); USING "####"; DaysINRead
     ELSE
       PRINT #UBRpt, ""
     END IF

    PRINT #UBRpt, TAB(50); Num2Date$(UBBillRec(1).PastDueDate); TAB(64); USING "#####.##"; TotalAmt# + TenPct#
    
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    
    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(1).RevName, 5);
      IF UBBillRec(1).MtrTypes(1) > 0 THEN
        PRINT #UBRpt, TAB(7); USING "#########"; CurrRead&;
        PRINT #UBRpt, TAB(18); USING "#########"; PrevRead&;
        PRINT #UBRpt, TAB(28); USING "#######"; UsageAmt&;
      END IF
      PRINT #UBRpt, TAB(36); USING "#####.##"; UBBillRec(1).RevAmt(1)
    ELSE
      PRINT #UBRpt,
    END IF

    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(2).RevName, 5);
      IF UBBillRec(1).MtrTypes(1) > 0 THEN
        PRINT #UBRpt, TAB(7); USING "#########"; CurrRead&;
        PRINT #UBRpt, TAB(18); USING "#########"; PrevRead&;
        PRINT #UBRpt, TAB(28); USING "#######"; UsageAmt&;
      END IF
      PRINT #UBRpt, TAB(36); USING "#####.##"; UBBillRec(1).RevAmt(2)
    ELSE
      PRINT #UBRpt,
    END IF

    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(3).RevName, 5);
      PRINT #UBRpt, TAB(36); USING "#####.##"; UBBillRec(1).RevAmt(3)
    ELSE
      PRINT #UBRpt,
    END IF

    IF UBBillRec(1).RevAmt(4) <> 0 THEN
      PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(4).RevName, 5);
      PRINT #UBRpt, TAB(36); USING "#####.##"; UBBillRec(1).RevAmt(4);
    END IF


    PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).CustName, 29)
    PRINT #UBRpt, TAB(47); UBCustRec(1).Addr1
    PRINT #UBRpt, TAB(47); UBCustRec(1).Addr2
    PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
    
    IF TotalTax# > 0 THEN
      PRINT #UBRpt, "                       TAX:  "; USING "$$,######.##"; TotalTax#
    ELSE
      PRINT #UBRpt, ""
    END IF
    
    IF Previous# <> 0 THEN
      PRINT #UBRpt, "                  Previous:  "; USING "$$,######.##"; Previous#
    ELSE
      PRINT #UBRpt,
    END IF
      PRINT #UBRpt, "                   Current:  "; USING "$$,######.##"; UBBillRec(1).TransAmt
    PRINT #UBRpt, "                           --------------"

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, "                   Deposit:  "; USING "$$,######.##"; -UBCustRec(1).DepositAmt
    ELSE
      PRINT #UBRpt,
    END IF

    IF TotalAmt# < 0 AND FinalFlag THEN
      PRINT #UBRpt, "                Refund Due:  "; USING "$$,######.##"; ABS(TotalAmt#)
    ELSE
      'STOP
      PRINT #UBRpt, "                     Total:  "; USING "$$,######.##"; TotalAmt#
    END IF

    PRINT #UBRpt,
    PRINT #UBRpt, "  "; CustMsg$
    PRINT #UBRpt, "  "; Message$
    PRINT #UBRpt,
    PRINT #UBRpt, "~"

