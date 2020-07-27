
    CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)
      
    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
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

    dRead$ = LEFT$(DateRead$, 6) + RIGHT$(DateRead$, 2)
    pRead$ = LEFT$(PrevDate$, 6) + RIGHT$(PrevDate$, 2)

    bDate$ = Num2Date$(UBBillRec(1).BillDate)
    bDate$ = LEFT$(bDate$, 6) + RIGHT$(bDate$, 2)
    pDate$ = Num2Date$(UBBillRec(1).PastDueDate)
    pDate$ = LEFT$(pDate$, 6) + RIGHT$(pDate$, 2)

    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF

    IF TotalAmt# > 0 THEN
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
    
    PRINT #UBRpt, "~"; TAB(50); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, pRead$; "  "; dRead$; "  "; bDate$; "  "; pDate$

    PRINT #UBRpt,
    PRINT #UBRpt,

'CONST MtrWaterOnly = 1
'CONST MtrSewerOnly = 2
'CONST MtrCombined = 3
'CONST MtrElectric = 4
'CONST MtrDemand = 5
'CONST MtrGas = 6

    PCnt = 0
    FOR WRevCnt = 1 TO 7
      PCnt = PCnt + 1
      IF UBBillRec(1).RevAmt(WRevCnt) <> 0 THEN
        SELECT CASE WRevCnt
        CASE 1
          RV$ = "EL"
          LookFor = 4
        CASE 2
          RV$ = "DM"
          LookFor = 5
        CASE 3
          RV$ = "WA"
          LookFor = 3
        CASE 4
          RV$ = "SW"
          LookFor = 3
        CASE 5
          RV$ = "YL"
          LookFor = -1
        CASE 6
          RV$ = "SH"
          LookFor = -1
        CASE 7
          RV$ = "SA"
          LookFor = -1
        CASE ELSE
          LookFor = -1
        END SELECT
        FOR MCnt = 1 TO 7
          IF UBBillRec(1).MtrTypes(MCnt) = LookFor AND UBBillRec(1).MtrTypes(MCnt) > 0 THEN
            UBBillRec(1).MtrTypes(MCnt) = 0
            UsageAmt& = UBBillRec(1).CurRead(MCnt) - UBBillRec(1).PrevRead(MCnt)
            IF UsageAmt& < 0 THEN
              MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(MCnt))) - 1)
              UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MCnt)) + UBBillRec(1).CurRead(MCnt)
            END IF
            PRINT #UBRpt, USING "#######"; UBBillRec(1).CurRead(MCnt);
            PRINT #UBRpt, TAB(10); USING "#######"; UBBillRec(1).PrevRead(MCnt);
            PRINT #UBRpt, TAB(19); USING "#######"; UsageAmt&;
            EXIT FOR
          ELSEIF LookFor = 3 AND UBBillRec(1).MtrTypes(MCnt) <= LookFor AND UBBillRec(1).MtrTypes(MCnt) > 0 THEN
            UBBillRec(1).MtrTypes(MCnt) = 0
            UsageAmt& = UBBillRec(1).CurRead(MCnt) - UBBillRec(1).PrevRead(MCnt)
            IF UsageAmt& < 0 THEN
              MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(MCnt))) - 1)
              UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MCnt)) + UBBillRec(1).CurRead(MCnt)
            END IF
            PRINT #UBRpt, USING "#######"; UBBillRec(1).CurRead(MCnt);
            PRINT #UBRpt, TAB(10); USING "#######"; UBBillRec(1).PrevRead(MCnt);
            PRINT #UBRpt, TAB(19); USING "#######"; UsageAmt&;
            EXIT FOR
          END IF
        NEXT
        PRINT #UBRpt, TAB(29); RV$;   'LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 2);
        PRINT #UBRpt, TAB(32); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt);
      END IF
      SELECT CASE PCnt
      CASE 4
        PRINT #UBRpt, TAB(44); LEFT$(UBCustRec(1).CustName, 29)
      CASE 5
        PRINT #UBRpt, TAB(44); UBCustRec(1).Addr1
      CASE 6
        PRINT #UBRpt, TAB(44); UBCustRec(1).Addr2
      CASE 7
        PRINT #UBRpt, TAB(44); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode
      CASE ELSE
        PRINT #UBRpt,
      END SELECT
    NEXT
    
    IF TotalTax# > 0 THEN
      PRINT #UBRpt, TAB(28); "TAX"; TAB(32); USING "#####.##"; TotalTax#
    ELSE
      PRINT #UBRpt, ""
    END IF

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, TAB(15); "Less Deposit"; TAB(32); USING "#####.##"; -UBCustRec(1).DepositAmt
    ELSE
      PRINT #UBRpt,
    END IF
    
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    
    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(30); USING "#######.##"; Previous#;
    END IF
    PRINT #UBRpt, TAB(44); UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb;
    PRINT #UBRpt, TAB(60); pDate$

    PRINT #UBRpt,
    PRINT #UBRpt, TAB(31); "TOTAL DUE"; TAB(59); "TOTAL DUE"
    PRINT #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb;
    PRINT #UBRpt, USING "######.##"; TAB(20); UBBillRec(1).TransAmt; TAB(31); TotalAmt#;
    PRINT #UBRpt, USING "######.##"; TAB(46); UBBillRec(1).TransAmt; TAB(59); TotalAmt#

    PRINT #UBRpt, TAB(3); Message$
    PRINT #UBRpt, TAB(3); Msg2$
    PRINT #UBRpt, TAB(3); Msg3$
    PRINT #UBRpt, TAB(3); Msg4$
    PRINT #UBRpt, "~"

