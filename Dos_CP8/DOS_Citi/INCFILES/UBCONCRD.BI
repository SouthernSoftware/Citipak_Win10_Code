'New Utility Bill format 10-28-96 BAR CODE PRINTABLE
'MUST SHOW BOTH METERS OR, TOTAL CONSUMPTION ON THIS BILL


    CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)

    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);  ':  M
      'put printer in     8 lpi             12 cpi  oki mode
    END IF

'**************************************
    EFoundMtr = False
    WFoundMtr = False

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
          WFoundMtr = True
        CASE 4, 5
          ECurrRead& = UBBillRec(1).CurRead(mChk)
          EPrevRead& = UBBillRec(1).PrevRead(mChk)
          EUsageAmt& = ECurrRead& - EPrevRead&
          IF EUsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(EPrevRead&)) - 1)
            EUsageAmt& = (MaxMeterAmt& - EPrevRead&) + ECurrRead&
          END IF
          EFoundMtr = True
        END SELECT
      END IF
    NEXT

'**************************************

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
      'TenPct# = 0
      'TenPct# = 2
      'TenPct# = 5
      'TenPct# = 5
      'TenPct# = 8
      'TenPct# = 10
      'TenPct# = 15
      'TenPct# = 25
      TenPct# = Round#(TotalAmt# * .1)
      'TenPct# = Round#(TotalAmt# * .05)
      'TenPct# = Round#(TotalAmt# * .03)     'Taloga
      'TenPct# = Round#(UBBillRec(1).TransAmt * .1)
      'TenPct# = TenPct# + 10
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

'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
    PRINT #UBRpt,

    PCnt = 0
    FOR WRevCnt = 1 TO 7
      PCnt = PCnt + 1
      IF UBBillRec(1).RevAmt(WRevCnt) <> 0 THEN
        PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 5);
        SELECT CASE PCnt
        CASE 1, 2      ', 6 'water/sewer
          IF WFoundMtr THEN
            PRINT #UBRpt, TAB(7); USING "#########"; WCurrRead&;
            PRINT #UBRpt, TAB(17); USING "#########"; WPrevRead&;
            PRINT #UBRpt, TAB(28); USING "#######"; WUsageAmt&;
          END IF
        CASE 3 'electric
          IF EFoundMtr THEN
            PRINT #UBRpt, TAB(7); USING "#########"; ECurrRead&;
            PRINT #UBRpt, TAB(17); USING "#########"; EPrevRead&;
            PRINT #UBRpt, TAB(28); USING "#######"; EUsageAmt&;
          END IF
        END SELECT
        PRINT #UBRpt, TAB(36); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt);
      END IF
      SELECT CASE PCnt
      CASE 4
        IF INSTR(UBCustRec(1).CustName, "DALE WATSON") > 0 THEN
          PRINT #UBRpt, TAB(47); QPTrim$(LEFT$(UBCustRec(1).CustName, 29)); " !"
        ELSE
          PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).CustName, 29)
        END IF
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
      PRINT #UBRpt, "                  Past Due:  "; USING "$$,######.##"; Previous#;
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


'    REM  Concord, NC Card Format

'    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
'    TotalAmt# = Round(UBBillRec(1).TransAmt + Previous#)
'    UsageAmt& = UBBillRec(1).CurRead(1) - UBBillRec(1).PrevRead(1)
'    'STOP
'    IF UsageAmt& < 0 THEN
'       MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(1))) - 1)
'       UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(1)) + UBBillRec(1).CurRead(1)
'    END IF
'
'    PRINT #UBRpt, CHR$(27); CHR$(58);    'CHR$(27); CHR$(58);
'    PRINT #UBRpt, "TOP"
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).CustName, 25)
'    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).Addr1, 25)
'    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).Addr2, 25)
'    PRINT #UBRpt, TAB(2); USING "########"; UBBillRec(1).PrevRead(1); TAB(13); UBBillRec(1).CurRead(1); TAB(23); UsageAmt&;
'    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).CITY, 14); " "; UBCustRec(1).State
'    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).ZipCode, 5); "-"; MID$(UBCustRec(1).ZipCode, 6, 4)
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt, TAB(2); USING "####.##"; UBBillRec(1).RevAmt(1);
'    TheRest# = 0
'    FOR FCnt = 2 TO 10
'      TheRest# = Round#(TheRest# + UBBillRec(1).RevAmt(FCnt))
'    NEXT
'    PRINT #UBRpt, TAB(17); USING "####.##"; TheRest#;
'
'    IF Previous# > 0 THEN
'      PRINT #UBRpt, TAB(25); USING "####.##"; Previous#;
'    END IF
'      PRINT #UBRpt, TAB(37); "Acct#: "; CustAcctNo&
'    PRINT #UBRpt,
'
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt, TAB(3); LEFT$(BILLDATE$, 6) + RIGHT$(BILLDATE$, 2);
'    PRINT #UBRpt, TAB(15); USING "######"; CustAcctNo&;
'    PRINT #UBRpt, TAB(25); USING "$####.##"; TotalAmt#;
'    PRINT #UBRpt, TAB(35); LEFT$(BILLDATE$, 6) + RIGHT$(BILLDATE$, 2)
'    IF CDeposit# > 0 THEN
'      PRINT #UBRpt, TAB(50); "Less Deposit"
'      PRINT #UBRpt, TAB(50); USING "####.##"; CDeposit#
'      PRINT #UBRpt,
'    ELSE
'      PRINT #UBRpt,
'      PRINT #UBRpt,
'      PRINT #UBRpt,
'    END IF
'    PRINT #UBRpt,
'    PRINT #UBRpt, TAB(37); USING "######"; CustAcctNo&;
'    PRINT #UBRpt, TAB(49); USING "$####.##"; TotalAmt#
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt, Message$

