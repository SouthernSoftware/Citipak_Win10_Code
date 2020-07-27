
    FoundMtr = False
    WCurrRead& = 0
    WPrevRead& = 0
    WUsageAmt& = 0

    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        SELECT CASE UBBillRec(1).MtrTypes(mChk)
        CASE 1, 2, 3
          WCurrRead& = UBBillRec(1).CurRead(mChk)
          WPrevRead& = UBBillRec(1).PrevRead(mChk)
          WUsageAmt& = WCurrRead& - WPrevRead&
          IF WUsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(WPrevRead&)) - 1)
            WUsageAmt& = (MaxMeterAmt& - WPrevRead&) + WCurrRead&
          END IF
          FoundMtr = True
        END SELECT
      END IF
    NEXT

    CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)

    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    END IF

    IF FoundAMtr = False AND BillInfoRec(1).PRDate <= 0 THEN
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
    TenPct# = 0
    AcctNum = UBBillRec(1).CustAcctNo
    Acct$ = QPTrim$(STR$(AcctNum))
    SELECT CASE AcctNum
    CASE IS < 10
      Acct$ = "00" + Acct$
    CASE IS < 100
      Acct$ = "0" + Acct$
    END SELECT
    AcctLen = LEN(Acct$)
    
    PRINT #UBRpt, "~"; TAB(50); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(10); LEFT$(UBCustRec(1).ServAddr, 26); TAB(50); USING "########"; UBBillRec(1).CustAcctNo  'line 5
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, PrevDate$; TAB(12); DateRead$;     'line 8
    IF DaysINRead > 0 THEN
      PRINT #UBRpt, TAB(23); USING "##"; DaysINRead;
    END IF
    PRINT #UBRpt, TAB(26); Num2Date$(UBBillRec(1).BillDate); TAB(37); Num2Date$(UBBillRec(1).PastDueDate); TAB(48); Num2Date$(UBBillRec(1).PastDueDate); TAB(59); Num2Date$(UBBillRec(1).PastDueDate + 8)
    PRINT #UBRpt, 'line 9
    PRINT #UBRpt, TAB(40); Message$ 'line 10
    PRINT #UBRpt, USING "########"; WCurrRead&; TAB(9); WPrevRead&; TAB(18); WUsageAmt&;
    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      PRINT #UBRpt, TAB(27); USING "W #####.##"; UBBillRec(1).RevAmt(1);
    END IF
    PRINT #UBRpt, TAB(40); Msg2$

    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      PRINT #UBRpt, TAB(27); USING "S #####.##"; UBBillRec(1).RevAmt(2);
    END IF
    PRINT #UBRpt, TAB(40); Msg3$

    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      PRINT #UBRpt, TAB(27); USING "  #####.##"; UBBillRec(1).RevAmt(3);
    END IF
    PRINT #UBRpt, TAB(40); Msg4$

    IF UBBillRec(1).RevAmt(4) <> 0 THEN
      PRINT #UBRpt, TAB(27); USING "  #####.##"; UBBillRec(1).RevAmt(4);
    END IF
    PRINT #UBRpt, TAB(40); CustMsg$

    PRINT #UBRpt,
    PRINT #UBRpt,

    IF FinalFlag AND CDeposit# THEN  'line 17
      PRINT #UBRpt, TAB(18); "Deposit:";
      PRINT #UBRpt, TAB(28); USING "######.##"; -UBCustRec(1).DepositAmt
    ELSE
      PRINT #UBRpt,
    END IF


    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, USING "#####.##"; UBBillRec(1).TransAmt; TAB(15); Previous#; TAB(29); TotalAmt#
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(3); USING "######"; UBBillRec(1).CustAcctNo;
    PRINT #UBRpt, TAB(40); LEFT$(UBCustRec(1).CustName, 29)
    PRINT #UBRpt, TAB(40); UBCustRec(1).Addr1
    PRINT #UBRpt, TAB(40); UBCustRec(1).Addr2
    PRINT #UBRpt, TAB(3); USING "#####.##"; TotalAmt#;
    PRINT #UBRpt, TAB(40); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State;
    PRINT #UBRpt, TAB(40); UBCustRec(1).ZipCode
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, "B"

'unrem
'    IF FinalFlag THEN
'      PRINT #UBRpt, TAB(64);
'    ELSE
'      PRINT #UBRpt, TAB(64); USING "#####.##"; Round#(TotalAmt# + TenPct#)
'    END IF
    
    
'    IF TotalTax# > 0 THEN
'      PRINT #UBRpt, "                       TAX:  "; USING "$$,######.##"; TotalTax#
'    ELSE
'      PRINT #UBRpt, ""
'    END IF
'    PRINT #UBRpt, TAB(64); USING "#####.##"; TotalAmt#
    
'    IF Previous# <> 0 THEN
'      PRINT #UBRpt, "                  Previous:  "; USING "$$,######.##"; Previous#;
'    END IF
'    PRINT #UBRpt, TAB(47); "LOC: "; UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb
'    PRINT #UBRpt, "                   Current:  "; USING "$$,######.##"; UBBillRec(1).TransAmt
'    PRINT #UBRpt, "                           --------------"
'
'    IF FinalFlag AND CDeposit# THEN
'      PRINT #UBRpt, "                   Deposit:  "; USING "$$,######.##"; -UBCustRec(1).DepositAmt
'    ELSE
'      PRINT #UBRpt,
'    END IF
'
'    IF TotalAmt# < 0 AND FinalFlag THEN
'      PRINT #UBRpt, "                Refund Due:  "; USING "$$,######.##"; ABS(TotalAmt#)
'    ELSE
'      PRINT #UBRpt, "                     Total:  "; USING "$$,######.##"; TotalAmt#
'    END IF
'
'    PRINT #UBRpt, TAB(3); CustMsg$
'
'    PRINT #UBRpt, TAB(3); Message$;
'    IF DidADraftFlag THEN
'      PRINT #UBRpt, TAB(47); "DRAFT NOTICE DO NOT PAY!!"
'    ELSE
'      PRINT #UBRpt,
'    END IF
'
'    PRINT #UBRpt, TAB(3); Msg2$;
'    IF DidADraftFlag THEN
'      PRINT #UBRpt, TAB(47); "DRAFT DATE: "; Num2Date$(BillInfoRec(1).DrftDate)
'    ELSE
'      PRINT #UBRpt,
'    END IF
'
'    PRINT #UBRpt, TAB(3); Msg3$

