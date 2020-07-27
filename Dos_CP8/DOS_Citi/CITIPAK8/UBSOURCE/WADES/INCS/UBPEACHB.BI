
    CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)
    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    END IF

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
      IF UBBillRec(1).curread(WMtrCnt) < 0 THEN
        UBBillRec(1).curread(WMtrCnt) = 0
      END IF
      PrevRead& = UBBillRec(1).PrevRead(WMtrCnt)
      CurrRead& = UBBillRec(1).curread(WMtrCnt)
      UsageAmt& = CurrRead& - PrevRead&
      IF UsageAmt& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WMtrCnt))) - 1)
        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WMtrCnt)) + UBBillRec(1).curread(WMtrCnt)
      END IF
      EXIT FOR
    END SELECT
  NEXT
'**** Find a meter
  AcctNum = UBBillRec(1).CustAcctNo
  Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)

  TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

  IF TotalAmt# > 0 THEN
    'TenPct# = 10
    TenPct# = Round#(TotalAmt# * .1)
  ELSE
    TenPct# = 0
  END IF

  AcctNum = UBBillRec(1).CustAcctNo

  PRINT #UBRpt, "~"; TAB(25); USING "########"; FBillNO& + PrintedCnt
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(2); USING "########"; CurrRead&; TAB(15); PrevRead&; TAB(25); UsageAmt&
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(11); PrevDate$; TAB(24); DateRead$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(3); UBCustRec(1).Serv(1).RateCode; TAB(25); USING "#####.##"; UBBillRec(1).RevAmt(1)
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(3); UBCustRec(1).Serv(2).RateCode; TAB(25); USING "#####.##"; UBBillRec(1).RevAmt(2)
  'PRINT #UBRpt,
  PRINT #UBRpt, TAB(51); USING "#####.##"; TotalAmt#
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(52); USING "#####"; AcctNum
  PRINT #UBRpt, TAB(25); USING "#####.##"; Previous#

  PRINT #UBRpt, TAB(49); PastDueDate$
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(25); USING "#####.##"; TotalAmt#

  PRINT #UBRpt, TAB(36); LEFT$(UBCustRec(1).CustName, 25)
  PRINT #UBRpt, TAB(25); USING "#####"; AcctNum;
  PRINT #UBRpt, TAB(36); UBCustRec(1).Addr1
  PRINT #UBRpt, TAB(36); UBCustRec(1).Addr2
  PRINT #UBRpt, TAB(23); PastDueDate$;
  PRINT #UBRpt, TAB(36); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(3); CustMsg$
  PRINT #UBRpt, TAB(3); Message$
  PRINT #UBRpt,
  PRINT #UBRpt, "~"

