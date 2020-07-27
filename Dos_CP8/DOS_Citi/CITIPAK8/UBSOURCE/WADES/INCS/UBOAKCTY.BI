'Oak City Bill format 11-05-97 Modifed std bill
'
  CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)
  IF NOT LPIFlag THEN
    LPIFlag = -2
    PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
    'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
    ' put printer in     8 lpi             12 cpi  oki mode
  END IF
  
  TotalTax# = 0
  FOR TaxCnt = 1 TO MaxRevsCnt
    TotalTax# = Round(TotalTax# + UBBillRec(1).TaxAmt(TaxCnt))
  NEXT

  FoundAMtr = False
  
  FOR mChk = 1 TO 7
    IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
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
  
  AcctNum = UBBillRec(1).CustAcctNo
  Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)
  
  IF TotalAmt# > 0 THEN
    TenPct# = 10
    'TenPct# = Round#(UBBillRec(1).TransAmt * .1)
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
  PRINT #UBRpt, TAB(15); LEFT$(UBCustRec(1).SERVADDR, 19); TAB(50); USING "########"; UBBillRec(1).CustAcctNo;
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
      PRINT #UBRpt, " WATER"; TAB(7); USING "#########"; UBBillRec(1).CurRead(WMtrCnt);
      PRINT #UBRpt, TAB(17); USING "#########"; UBBillRec(1).PrevRead(WMtrCnt);
      UsageAmt& = UBBillRec(1).CurRead(WMtrCnt) - UBBillRec(1).PrevRead(WMtrCnt)
      IF UsageAmt& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WMtrCnt))) - 1)
        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WMtrCnt)) + UBBillRec(1).CurRead(WMtrCnt)
      END IF
      PRINT #UBRpt, TAB(28); USING "#######"; UsageAmt&;
      PRINT #UBRpt, TAB(36); USING "#####.##"; UBBillRec(1).RevAmt(1)

      
      DidWMrt = True
      EXIT FOR
    END SELECT
  NEXT
  
  IF NOT DidWMrt THEN
    PRINT #UBRpt,
  END IF
  
  DidEMrt = False
  FOR EMtrCnt = 1 TO 7
    IF UBBillRec(1).MtrTypes(EMtrCnt) = 4 THEN
      IF UBBillRec(1).PrevRead(EMtrCnt) < 0 THEN
        UBBillRec(1).PrevRead(EMtrCnt) = 0
      END IF
      IF UBBillRec(1).CurRead(EMtrCnt) < 0 THEN
        UBBillRec(1).CurRead(EMtrCnt) = 0
      END IF
      PRINT #UBRpt, " ELEC."; TAB(7); USING "#########"; UBBillRec(1).CurRead(EMtrCnt);
      PRINT #UBRpt, TAB(17); USING "#########"; UBBillRec(1).PrevRead(EMtrCnt);
      UsageAmt& = UBBillRec(1).CurRead(EMtrCnt) - UBBillRec(1).PrevRead(EMtrCnt)
      IF UsageAmt& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(EMtrCnt))) - 1)
        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(EMtrCnt)) + UBBillRec(1).CurRead(EMtrCnt)
      END IF
      PRINT #UBRpt, TAB(28); USING "#######"; UsageAmt&;
      PRINT #UBRpt, TAB(36); USING "#####.##"; UBBillRec(1).RevAmt(2)
      DidEMrt = True
      EXIT FOR
    END IF
  NEXT
  
  IF NOT DidEMrt THEN
    PRINT #UBRpt,
  END IF
  
  PCnt = 2
  
  FOR WRevCnt = 3 TO 7
    PCnt = PCnt + 1
    IF UBBillRec(1).RevAmt(WRevCnt) <> 0 THEN
      PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 5);
      PRINT #UBRpt, TAB(36); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt);
    END IF
    
    SELECT CASE PCnt
    CASE 4
      PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).CustName, 29)
    CASE 5
      PRINT #UBRpt, TAB(47); UBCustRec(1).ADDR1
    CASE 6
      PRINT #UBRpt, TAB(47); UBCustRec(1).Addr2
    CASE 7
      PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).CITY, 14); " "; UBCustRec(1).STATE; " "; LEFT$(UBCustRec(1).ZIPCODE, 5)
    CASE ELSE
      PRINT #UBRpt,
    END SELECT
  NEXT
  
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
    TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
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

