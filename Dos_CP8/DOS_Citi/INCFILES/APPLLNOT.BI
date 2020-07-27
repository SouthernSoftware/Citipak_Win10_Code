
    FDate$ = Num2Date(NoticeInfo.NoticeDate)
    TDate$ = Num2Date(NoticeInfo.PayByDate)
    
    Loca$ = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb
    'CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)

    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi
    END IF

    AcctNum = AcctNo&

    Previous# = UBCustRec(1).PrevBalance
    Current# = UBCustRec(1).CurrBalance
    TotalAmt# = Round#(Previous# + Current#)

    AcctNum = AcctNo&

    Acct$ = QPTrim$(STR$(AcctNum))
    SELECT CASE AcctNum
    CASE IS < 10
      Acct$ = "00" + Acct$
    CASE IS < 100
      Acct$ = "0" + Acct$
    END SELECT
    AcctLen = LEN(Acct$)

    FMth$ = LEFT$(FDate$, 2)
    FDay$ = MID$(FDate$, 4, 2)

    TMth$ = LEFT$(TDate$, 2)
    TDay$ = MID$(TDate$, 4, 2)

    DMth$ = LEFT$(PDate$, 2)
    DDay$ = MID$(PDate$, 4, 2)
    DYer$ = RIGHT$(PDate$, 4)

    PRINT #UBRpt, "~"
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, ""  '11
    PRINT #UBRpt, TAB(2); FMth$; TAB(8); FDay$; TAB(14); TMth$; TAB(20); TDay$;
    PRINT #UBRpt, TAB(26); DMth$; TAB(32); DDay$; TAB(38); DYer$; TAB(44); "PRE-SORT"; "  "; Loca$
    PRINT #UBRpt, TAB(54); USING "#########"; AcctNo&
    PRINT #UBRpt, TAB(44); LEFT$(UBCustRec(1).CustName, 23)
    PRINT #UBRpt, TAB(44); LEFT$(UBCustRec(1).Addr1, 23)
    PRINT #UBRpt, TAB(44); LEFT$(UBCustRec(1).Addr2, 23)
    PRINT #UBRpt, TAB(44); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)

    IF UBCustRec(1).CurrRevAmts(1) <> 0 THEN
      PRINT #UBRpt, TAB(29); LEFT$(UBSetUpRec(1).Revenues(1).RevName, 3);
      PRINT #UBRpt, TAB(33); USING "#####.##"; UBCustRec(1).CurrRevAmts(1)
    ELSE
      PRINT #UBRpt, 'line 17
    END IF

    IF UBCustRec(1).CurrRevAmts(2) <> 0 THEN
      IF FoundAMtr THEN
        PRINT #UBRpt, USING "#########"; PRead#;
        PRINT #UBRpt, TAB(10); USING "#########"; CRead#; UsageAmt&;
      END IF
      PRINT #UBRpt, TAB(29); LEFT$(UBSetUpRec(1).Revenues(2).RevName, 3);
      PRINT #UBRpt, TAB(33); USING "#####.##"; UBCustRec(1).CurrRevAmts(2)
    ELSE
      PRINT #UBRpt, 'line 18
    END IF

    FOR RevCnt = 3 TO 7
      IF UBCustRec(1).CurrRevAmts(RevCnt) <> 0 THEN
        PRINT #UBRpt, TAB(29); LEFT$(UBSetUpRec(1).Revenues(RevCnt).RevName, 3);
        PRINT #UBRpt, TAB(33); USING "#####.##"; UBCustRec(1).CurrRevAmts(RevCnt)
      ELSE
        PRINT #UBRpt,
      END IF
    NEXT

    PRINT #UBRpt,

    PRINT #UBRpt,   '25
    PRINT #UBRpt,   '26
    PRINT #UBRpt, TAB(17); Location$; TAB(29); USING "$$,######.##"; TotalAmt#; TAB(44); TotalAmt#
    PRINT #UBRpt, '27
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, "~"

