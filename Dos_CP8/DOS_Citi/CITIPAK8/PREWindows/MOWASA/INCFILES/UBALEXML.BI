    
    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        CurrRead# = UBBillRec(1).CurRead(mChk)
        PrevRead# = UBBillRec(1).PrevRead(mChk)
        UsageAmt# = CurrRead# - PrevRead#
        IF UsageAmt# < 0 THEN
          MaxMeterAmt# = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(mChk))) - 1)
          UsageAmt# = (MaxMeterAmt# - PrevRead#) + CurrRead#
        END IF
        FoundAMtr = True
        EXIT FOR
      END IF
    NEXT

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF

    PRINT #UBRpt, "~"; TAB(27); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(11); PastDate$; TAB(23); DateRead$
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); USING "########"; PrevRead#; TAB(13); CurrRead#;
    PRINT #UBRpt, TAB(27); USING "########"; UsageAmt#
    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).CustName, 29)
    PRINT #UBRpt, TAB(37); UBCustRec(1).Addr1
    PRINT #UBRpt, TAB(37); UBCustRec(1).Addr2
    PRINT #UBRpt, TAB(37); QPTrim$(LEFT$(UBCustRec(1).CITY, 14)); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
    PRINT #UBRpt, TAB(2); USING "####.##"; UBBillRec(1).RevAmt(1);
    PRINT #UBRpt, TAB(11); USING "####.##"; UBBillRec(1).RevAmt(2);
    PRINT #UBRpt, TAB(19); USING "####.##"; UBBillRec(1).RevAmt(4);
    PRINT #UBRpt, TAB(28); USING "####.##"; Previous#
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,

    PRINT #UBRpt, TAB(2); BillDate$;

    PRINT #UBRpt, TAB(16); USING "#####"; AcctNum;
    PRINT #UBRpt, TAB(27); USING "$####.##"; TotalAmt#;
    PRINT #UBRpt, TAB(36); BillDate$
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); Message$
    PRINT #UBRpt, TAB(2); Msg2$
    PRINT #UBRpt, TAB(2); Msg3$; TAB(37); USING "#####"; AcctNum;
    PRINT #UBRpt, TAB(49); USING "$####.##"; TotalAmt#
    PRINT #UBRpt, "~"

