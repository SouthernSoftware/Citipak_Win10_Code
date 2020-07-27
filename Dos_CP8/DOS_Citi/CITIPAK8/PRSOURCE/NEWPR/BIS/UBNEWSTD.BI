'New Utility Bill format 10-28-96 BAR CODE PRINTABLE
'MUST SHOW BOTH METERS OR, TOTAL CONSUMPTION ON THIS BILL

    Zip$ = QPTrim$(UBCustRec(1).ZipCode)

    ZDigit$ = GetZipEDigit$(Zip$)

    Zip$ = Zip$ + ZDigit$

    CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)
      
    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    END IF
                                                      ':  M
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

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF

    IF TotalAmt# >= 5 THEN
      'TenPct# = 2
      'TenPct# = 0
      'TenPct# = 10
      'TenPct# = Round#(TotalAmt# * .1)
      TenPct# = Round#(UBBillRec(1).TransAmt * .1)
    ELSE
      TenPct# = 0
    END IF

'NOTE: this must be removed to apply a penalty
    TenPct# = 0
'-=-=-=-=-=-=-=-=-=-=-=-

    AcctNum = UBBillRec(1).CustAcctNo
    Acct$ = QPTrim$(STR$(AcctNum))
    SELECT CASE AcctNum
    CASE IS < 10
      Acct$ = "00" + Acct$
    CASE IS < 100
      Acct$ = "0" + Acct$
    END SELECT
    AcctLen = LEN(Acct$)

    Acct$ = QPTrim$(UBCustRec(1).ZipCode)
    AcctLen = LEN(Acct$)

    PRINT #UBRpt, "~"; TAB(50); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(18); PrevDate$; TAB(33); DateRead$;
     'Only Print Days if Greater than 0
     IF DaysINRead > 0 THEN
       PRINT #UBRpt, "    "; USING "####"; DaysINRead
     ELSE
       PRINT #UBRpt, ""
     END IF

    PRINT #UBRpt,
    PRINT #UBRpt,

    PCnt = 0
    FOR WRevCnt = 1 TO 6
      PCnt = PCnt + 1
      IF UBBillRec(1).RevAmt(WRevCnt) <> 0 THEN
        PRINT #UBRpt, LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 3);
        IF UBBillRec(1).CurRead(WRevCnt) > 0 THEN
          UsageAmt& = UBBillRec(1).CurRead(WRevCnt) - UBBillRec(1).PrevRead(WRevCnt)
          IF UsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WRevCnt))) - 1)
            UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WRevCnt)) + UBBillRec(1).CurRead(WRevCnt)
          END IF
          PRINT #UBRpt, TAB(4); USING "#########"; UBBillRec(1).PrevRead(WRevCnt);
          PRINT #UBRpt, TAB(14); USING "#########"; UBBillRec(1).CurRead(WRevCnt);
          PRINT #UBRpt, TAB(25); USING "#######"; UsageAmt&;
        END IF
        PRINT #UBRpt, TAB(33); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt);
      END IF
      SELECT CASE PCnt
      CASE 1
        PRINT #UBRpt, TAB(44); USING "##########"; UBBillRec(1).CustAcctNo
      CASE 5
        PRINT #UBRpt, TAB(49); LEFT$(UBCustRec(1).ServAddr, 26)
      CASE ELSE
        PRINT #UBRpt,
      END SELECT
    NEXT
    IF TotalTax# > 0 THEN
      PRINT #UBRpt, TAB(14); USING "     TAX:       $$,######.##"; TotalTax#
    ELSE
      PRINT #UBRpt, ""
    END IF
    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(14); USING "Previous:       $$,######.##"; Previous#
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt, TAB(14); USING " Current:       $$,######.##"; UBBillRec(1).TransAmt

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, TAB(14); USING " Deposit:       $$,######.##"; -UBCustRec(1).DepositAmt;
    END IF
    PRINT #UBRpt, TAB(45); Num2Date$(UBBillRec(1).BillDate); TAB(60); Num2Date$(UBBillRec(1).PastDueDate)
    PRINT #UBRpt, TAB(2); Message$
    PRINT #UBRpt, TAB(2); Msg2$
    PRINT #UBRpt, TAB(2); Msg3$
    IF DidADraftFlag THEN
      PRINT #UBRpt, TAB(2); "DRAFT NOTICE DO NOT PAY!!";
    ELSE
      PRINT #UBRpt, TAB(2); Msg4$;
    END IF
    
    IF TotalAmt# < 0 AND FinalFlag THEN
      PRINT #UBRpt, TAB(34); USING "Refund: $$,######.##"; ABS(TotalAmt#)
    ELSE
      PRINT #UBRpt, TAB(34); USING " Total: $$,######.##"; TotalAmt#
    END IF

    PRINT #UBRpt,
    PRINT #UBRpt, TAB(22); LEFT$(UBCustRec(1).CustName, 29)
    PRINT #UBRpt, TAB(22); UBCustRec(1).Addr1
    PRINT #UBRpt, TAB(22); UBCustRec(1).Addr2
    PRINT #UBRpt, USING "##########"; UBBillRec(1).CustAcctNo;
    PRINT #UBRpt, TAB(22); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, USING "#######.##"; TotalAmt#
    PRINT #UBRpt,
    PRINT #UBRpt,
    IF FinalFlag THEN
      PRINT #UBRpt, USING "#######.##"; Round#(TotalAmt#);
    ELSE
      PRINT #UBRpt, USING "#######.##"; Round#(TotalAmt# + TenPct#);
    END IF
    PRINT #UBRpt, TAB(22); CHR$(27); CHR$(16); "C"; CHR$(LEN(Zip$)); Zip$
    PRINT #UBRpt,
    PRINT #UBRpt, "~"

