'McCormick
'last revised 050399

  REDIM RevChk(1 TO 15)  AS INTEGER

  TotalTax# = 0
  DidWMrt = False
  DidEMrt = False
  NPrintCnt = 0

  FOR TaxCnt = 1 TO MaxRevsCnt
    TotalTax# = Round(TotalTax# + UBBillRec(1).TaxAmt(TaxCnt))
  NEXT

  Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

  IF (TotalAmt# > 0) AND (EPPFlag = 0) THEN
    TenPct# = Round#((UBBillRec(1).TransAmt - TotalTax#) * .1)
  END IF
  
  PRINT #UBRpt, "~"; TAB(50); CustAcctNo&
  PRINT #UBRpt, TAB(40); USING "#####"; PrintedCnt
  PRINT #UBRpt,
  PRINT #UBRpt,
  
  FOR EMtrCnt = 1 TO 7
    IF UBBillRec(1).MtrTypes(EMtrCnt) = 4 THEN
      IF UBBillRec(1).PrevRead(EMtrCnt) < 0 THEN
        UBBillRec(1).PrevRead(EMtrCnt) = 0
      END IF
      IF UBBillRec(1).CurRead(EMtrCnt) < 0 THEN
        UBBillRec(1).CurRead(EMtrCnt) = 0
      END IF
      PRINT #UBRpt, TAB(10); USING "#########"; UBBillRec(1).PrevRead(EMtrCnt);
      PRINT #UBRpt, TAB(19); USING "#########"; UBBillRec(1).CurRead(EMtrCnt);
      UsageAmt& = UBBillRec(1).CurRead(EMtrCnt) - UBBillRec(1).PrevRead(EMtrCnt)
      IF UsageAmt& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(EMtrCnt))) - 1)
        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(EMtrCnt)) + UBBillRec(1).CurRead(EMtrCnt)
      END IF
      PRINT #UBRpt, TAB(28); USING "#########"; UsageAmt&
      DidEMrt = True
      EXIT FOR
    END IF
  NEXT

  IF NOT DidEMrt THEN
    PRINT #UBRpt,
  END IF

  FOR WMtrCnt = 1 TO 7
    SELECT CASE UBBillRec(1).MtrTypes(WMtrCnt)
    CASE MtrWaterOnly, MtrSewerOnly, MtrCombined, MtrTouchRead
      IF UBBillRec(1).PrevRead(WMtrCnt) < 0 THEN
        UBBillRec(1).PrevRead(WMtrCnt) = 0
      END IF
      IF UBBillRec(1).CurRead(WMtrCnt) < 0 THEN
        UBBillRec(1).CurRead(WMtrCnt) = 0
      END IF
      PRINT #UBRpt, TAB(10); USING "#########"; UBBillRec(1).PrevRead(WMtrCnt);
      PRINT #UBRpt, TAB(19); USING "#########"; UBBillRec(1).CurRead(WMtrCnt);
      UsageAmt& = UBBillRec(1).CurRead(WMtrCnt) - UBBillRec(1).PrevRead(WMtrCnt)
      IF UsageAmt& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WMtrCnt))) - 1)
        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WMtrCnt)) + UBBillRec(1).CurRead(WMtrCnt)
      END IF
      PRINT #UBRpt, TAB(28); USING "#########"; UsageAmt&
      DidWMrt = True
      EXIT FOR
    END SELECT
  NEXT

  IF NOT DidWMrt THEN
    PRINT #UBRpt,
  END IF

  PRINT #UBRpt, TAB(12); UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb;
  PRINT #UBRpt, TAB(28); BillDate$; TAB(39); LEFT$(UBCustRec(1).SERVADDR, 27)
  PRINT #UBRpt,

  IF UBBillRec(1).RevAmt(1) > 0 THEN
    PRINT #UBRpt, TAB(7); LEFT$(UBSetUpRec(1).Revenues(1).RevName, 3); TAB(27); USING "######.##"; UBBillRec(1).RevAmt(1)
    RevChk(1) = True
  ELSE
    PRINT #UBRpt,
  END IF

  IF UBBillRec(1).RevAmt(2) > 0 THEN
    PRINT #UBRpt, TAB(7); LEFT$(UBSetUpRec(1).Revenues(2).RevName, 3); TAB(27); USING "######.##"; UBBillRec(1).RevAmt(2);
    RevChk(2) = True
  END IF
  PRINT #UBRpt, TAB(40); LEFT$(UBCustRec(1).CustName, 27)

  IF UBBillRec(1).RevAmt(3) > 0 THEN
    PRINT #UBRpt, TAB(7); LEFT$(UBSetUpRec(1).Revenues(3).RevName, 3); TAB(27); USING "######.##"; UBBillRec(1).RevAmt(3);
    RevChk(3) = True
  END IF
  PRINT #UBRpt, TAB(40); LEFT$(UBCustRec(1).ADDR1, 27)

  IF UBBillRec(1).RevAmt(9) > 0 THEN
    PRINT #UBRpt, TAB(7); LEFT$(UBSetUpRec(1).Revenues(9).RevName, 3); TAB(27); USING "######.##"; UBBillRec(1).RevAmt(9);
    RevChk(9) = True
  END IF

  PRINT #UBRpt, TAB(40); LEFT$(UBCustRec(1).CITY, 14); " ";
  PRINT #UBRpt, UBCustRec(1).STATE; " "; UBCustRec(1).ZIPCODE

'******
'051099 Corrected
  FOR FRCnt = 1 TO 2
    Do2Skip = False
    RevSrc = UBCustRec(1).FlatRates(FRCnt).RevSrc
    IF RevSrc > 0 THEN
      IF RevChk(RevSrc) = True THEN
        NPrintCnt = NPrintCnt + 1
        Do2Skip = True
      END IF
    END IF
    
    IF NOT Do2Skip THEN
      Desc$ = QPTrim$(UBCustRec(1).FlatRates(FRCnt).FRDesc)
      IF LEN(Desc$) > 0 THEN
        PRINT #UBRpt, TAB(7); LEFT$(Desc$, 3); TAB(27); USING "######.##"; UBBillRec(1).RevAmt(UBCustRec(1).FlatRates(FRCnt).RevSrc)
        RevChk(RevSrc) = True
      ELSE
        NPrintCnt = NPrintCnt + 1
      END IF
    END IF
  NEXT
'******

  IF TotalTax# <> 0 THEN
    PRINT #UBRpt, TAB(7); " Tax"; TAB(27); USING "######.##"; TotalTax#;
  END IF

  SELECT CASE EPPFlag
  CASE True
    PRINT #UBRpt, TAB(40); USING "EPP BAL: #####.##"; EPPTotBal#
  CASE False
    PRINT #UBRpt,
  CASE 1
    PRINT #UBRpt, TAB(40); "EPP YEAR END"
  END SELECT

  IF NPrintCnt > 0 THEN
    FOR NPCnt = 1 TO NPrintCnt
      PRINT #UBRpt,
    NEXT
  END IF

  PRINT #UBRpt,
  PRINT #UBRpt, TAB(7); USING "#####.##"; Previous#;
  PRINT #UBRpt, TAB(16); USING "######.##"; UBBillRec(1).TransAmt;
  PRINT #UBRpt, TAB(27); USING "######.##"; TotalAmt#;
  PRINT #UBRpt, TAB(39); UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb;
  PRINT #UBRpt, TAB(52); USING "######.##"; TotalAmt#
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(11); PastDueDate$;
  'LatePayAmt# = Round#(UBBillRec(1).TransAmt + (UBBillRec(1).TransAmt * .1))
  PRINT #UBRpt, TAB(27); USING "######.##"; Round#(TotalAmt# + TenPct#);
  PRINT #UBRpt, TAB(39); PastDueDate$; TAB(52); USING "######.##"; Round#(TotalAmt# + TenPct#)
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(7); Message$
  PRINT #UBRpt, TAB(7); Msg2$
  PRINT #UBRpt, "~"             ' "LAST LINE"

