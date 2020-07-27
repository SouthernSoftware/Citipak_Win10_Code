    CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)

    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);  ':  M
      'put printer in     8 lpi             12 cpi  oki mode
    END IF

    FoundAMtr = False

    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        FoundAMtr = True
        EXIT FOR
      END IF
    NEXT

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

    IF TotalAmt# > 0 THEN
      TenPct# = 25
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
    PRINT #UBRpt, TAB(15); LEFT$(UBCustRec(1).ServAddr, 26); TAB(50); USING "########"; UBBillRec(1).CustAcctNo;
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

    PRINT #UBRpt, TAB(50); Num2Date$(UBBillRec(1).PastDueDate);
'    PRINT #UBRpt, TAB(50); QPTrim$(Message$);
'unrem
    IF FinalFlag THEN
      PRINT #UBRpt, TAB(64); USING "#####.##"; Round#(TotalAmt#)
                    'unrem
    ELSE
      PRINT #UBRpt, TAB(64); USING "#####.##"; Round#(TotalAmt# + TenPct#)
                    'unrem
    END IF

    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, 'TAB(52); "Address Service"
    PRINT #UBRpt, 'TAB(52); "   Requested"

'    PRINT #UBRpt, STRING$(50, " "); CHR$(27); CHR$(16); "A";
'    PRINT #UBRpt, CHR$(8);
'    PRINT #UBRpt, CHR$(2); CHR$(0);
'    PRINT #UBRpt, CHR$(0); CHR$(2);
'    PRINT #UBRpt, CHR$(1); CHR$(1);
'    PRINT #UBRpt, CHR$(1); CHR$(2);
'    PRINT #UBRpt, CHR$(27); CHR$(16); "B"; CHR$(AcctLen); Acct$

    PCnt = 0
    FOR WRevCnt = 1 TO 7
      PCnt = PCnt + 1
      IF UBBillRec(1).RevAmt(WRevCnt) <> 0 THEN
        PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 5);
        IF UBBillRec(1).CurRead(WRevCnt) > 0 THEN
          UsageAmt& = UBBillRec(1).CurRead(WRevCnt) - UBBillRec(1).PrevRead(WRevCnt)
          IF UsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WRevCnt))) - 1)
            UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WRevCnt)) + UBBillRec(1).CurRead(WRevCnt)
          END IF
          PRINT #UBRpt, TAB(7); USING "##########"; UBBillRec(1).CurRead(WRevCnt);
          PRINT #UBRpt, TAB(17); USING "##########"; UBBillRec(1).PrevRead(WRevCnt);
          PRINT #UBRpt, TAB(28); USING "#######"; UsageAmt&;
        END IF
        PRINT #UBRpt, TAB(36); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt);
      END IF
      SELECT CASE PCnt

      CASE 4
        PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).CustName, 29)
      CASE 5
        PRINT #UBRpt, TAB(47); UBCustRec(1).Addr1
      CASE 6
        PRINT #UBRpt, TAB(47); UBCustRec(1).Addr2
      CASE 7
        PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
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
      PRINT #UBRpt, "                  Previous:  "; USING "$$,######.##"; Previous#;
    END IF
    PRINT #UBRpt, TAB(47); "LOC: "; UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb
    PRINT #UBRpt, "                   Current:  "; USING "$$,######.##"; UBBillRec(1).TransAmt
    PRINT #UBRpt, "                           --------------"

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, "                   Deposit:  "; USING "$$,######.##"; -UBCustRec(1).DepositAmt;
    END IF

    PRINT #UBRpt, ' TAB(47); "Route: "; UBCustRec(1).POSTRTE

    IF TotalAmt# < 0 AND FinalFlag THEN
      PRINT #UBRpt, "                Refund Due:  "; USING "$$,######.##"; ABS(TotalAmt#)
    ELSE
      PRINT #UBRpt, "                     Total:  "; USING "$$,######.##"; TotalAmt#
    END IF

    PRINT #UBRpt, TAB(3); CustMsg$
    PRINT #UBRpt, TAB(3); Message$;

    IF DidADraftFlag THEN
      PRINT #UBRpt, TAB(47); "DRAFT NOTICE DO NOT PAY!!"
    ELSE
      PRINT #UBRpt,
    END IF

    PRINT #UBRpt, TAB(3); Msg2$;
    IF DidADraftFlag THEN
      PRINT #UBRpt, TAB(47); "DRAFT DATE: "; Num2Date$(BillInfoRec(1).DrftDate)
    ELSE
      PRINT #UBRpt,
    END IF

    PRINT #UBRpt, TAB(3); Msg3$
    PRINT #UBRpt, TAB(3); Msg4$


'    IF UBBillRec(1).CurRead(1) >= 0 AND UBBillRec(1).PrevRead(1) >= 0 THEN
'      UsageAmt& = UBBillRec(1).CurRead(1) - UBBillRec(1).PrevRead(1)
'      IF UsageAmt& < 0 THEN
'        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(1))) - 1)
'        UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(1)) + UBBillRec(1).CurRead(1)
'      END IF
'    ELSE
'      UsageAmt& = 0
'    END IF
'
'    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
'    TotalAmt# = Round(UBBillRec(1).TransAmt + Previous#)
'
'    PRINT #UBRpt, "~"; USING "#####"; FBillNO& + PrintedCnt
'    PRINT #UBRpt,
'    PRINT #UBRpt, TAB(5); CustAcctNo&; TAB(20); DateRead$
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt, TAB(3); USING "#########"; UBBillRec(1).PrevRead(1);
'    PRINT #UBRpt, TAB(13); USING "#########"; UBBillRec(1).CurRead(1);
'    PRINT #UBRpt, TAB(24); USING "########"; UsageAmt&
'    PRINT #UBRpt,
'    IF UBBillRec(1).RevAmt(1) <> 0 THEN
'      PRINT #UBRpt, TAB(4); UBSetUpRec(1).Revenues(1).RevName; TAB(23); USING "######.##"; UBBillRec(1).RevAmt(1);
'    END IF
'    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).CustName, 25)
'
'    IF UBBillRec(1).RevAmt(2) <> 0 THEN
'      PRINT #UBRpt, TAB(4); UBSetUpRec(1).Revenues(2).RevName; TAB(23); USING "######.##"; UBBillRec(1).RevAmt(2);
'    END IF
'    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).Addr1, 25)
'
'    IF UBBillRec(1).RevAmt(5) <> 0 THEN
'      PRINT #UBRpt, TAB(4); UBSetUpRec(1).Revenues(5).RevName; TAB(23); USING "######.##"; UBBillRec(1).RevAmt(5);
'    END IF
'    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).Addr2, 25)
'
'    PRINT #UBRpt, TAB(37); QPTrim$(UBCustRec(1).City); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode
'
'    IF Previous# <> 0 THEN
'      PRINT #UBRpt, TAB(4); "Previous:"; TAB(23); USING "######.##"; Previous#
'    ELSE
'      PRINT #UBRpt, ""
'    END IF
'
'    IF FinalFlag AND CDeposit# THEN
'      PRINT #UBRpt, TAB(4); "Deposit:"; TAB(24); USING "######.##"; -UBCustRec(1).DepositAmt;
'      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
'      PRINT #UBRpt, TAB(37); CustAcctNo&
'    ELSE
'      PRINT #UBRpt, TAB(37); CustAcctNo&
'    END IF
'    PRINT #UBRpt,
'    PRINT #UBRpt, TAB(5); "DUE: "; PastDueDate$; TAB(23); USING "######.##"; TotalAmt#
'    PRINT #UBRpt,
'    PRINT #UBRpt, TAB(38); LEFT$(PastDueDate$, 6); RIGHT$(PastDueDate$, 2); TAB(49); USING "######.##"; TotalAmt#
'    PRINT #UBRpt, TAB(3); Message$
'    PRINT #UBRpt, TAB(3); Msg2$
'    PRINT #UBRpt, TAB(3); Msg3$
'    PRINT #UBRpt, "~"

