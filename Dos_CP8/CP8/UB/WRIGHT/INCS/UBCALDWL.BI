    FoundAMtr = False
    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        WCurrRead& = UBBillRec(1).CurRead(mChk)
        WPrevRead& = UBBillRec(1).PrevRead(mChk)
        WUsageAmt& = WCurrRead& - WPrevRead&
        IF WUsageAmt& < 0 THEN
          MaxMeterAmt& = 10& ^ (LEN(STR$(WPrevRead&)) - 1)
          WUsageAmt& = (MaxMeterAmt& - WPrevRead&) + WCurrRead&
        END IF
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

    SRDate$ = LEFT$(DateRead$, 6) + RIGHT$(DateRead$, 2)
    SPDate$ = LEFT$(PrevDate$, 6) + RIGHT$(PrevDate$, 2)
    SBDate$ = LEFT$(BillDate$, 6) + RIGHT$(BillDate$, 2)
    SDDate$ = LEFT$(PastDueDate$, 6) + RIGHT$(PastDueDate$, 2)

    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF

    'IF TotalAmt# > 0 THEN
    '  TenPct# = 10
    'ELSE
    '  TenPct# = 0
    'END IF


     PRINT #UBRpt, "~"; TAB(25); USING "######"; FBillNO& + PrintedCnt
     PRINT #UBRpt,
     PRINT #UBRpt,
     PRINT #UBRpt,
     PRINT #UBRpt, TAB(9); LEFT$(UBCustRec(1).ServAddr, 25)
     PRINT #UBRpt, TAB(4); SRDate$; TAB(23); USING "########"; WCurrRead&
     PRINT #UBRpt, TAB(4); SPDate$; TAB(23); USING "########"; WPrevRead&
     PRINT #UBRpt, TAB(23); USING "########"; WUsageAmt&
     PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).CustName, 25)
     PRINT #UBRpt, TAB(11); USING "#####.##"; Previous#;
     PRINT #UBRpt, TAB(35); UBCustRec(1).Addr1

     IF UBBillRec(1).RevAmt(1) <> 0 THEN
       PRINT #UBRpt, TAB(11); USING "#####.##"; UBBillRec(1).RevAmt(1);
     END IF
     PRINT #UBRpt, TAB(35); UBCustRec(1).Addr2

     IF UBBillRec(1).RevAmt(2) <> 0 THEN
       PRINT #UBRpt, TAB(11); USING "#####.##"; UBBillRec(1).RevAmt(2);
     END IF
     PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
     
     IF UBBillRec(1).RevAmt(3) <> 0 THEN
       PRINT #UBRpt, TAB(11); USING "#####.##"; UBBillRec(1).RevAmt(3)
     ELSE
       PRINT #UBRpt,
     END IF

     IF UBBillRec(1).RevAmt(4) <> 0 THEN
       PRINT #UBRpt, TAB(11); USING "#####.##"; UBBillRec(1).RevAmt(4);
     END IF
     PRINT #UBRpt, TAB(35); Message$

     IF FinalFlag AND CDeposit# THEN
       PRINT #UBRpt, TAB(11); USING "#####.##"; -UBCustRec(1).DepositAmt;
     END IF
     PRINT #UBRpt, TAB(35); Msg2$

     PRINT #UBRpt, TAB(24); SDDate$
     PRINT #UBRpt,
     PRINT #UBRpt,
     PRINT #UBRpt, SDDate$; TAB(12); USING "######"; AcctNum;
     PRINT #UBRpt, TAB(24); USING "#####.##"; TotalAmt#;
     PRINT #UBRpt, TAB(38); USING "######"; AcctNum;
     PRINT #UBRpt, TAB(50); USING "#####.##"; TotalAmt#
     PRINT #UBRpt,
     PRINT #UBRpt, "~"

