    
    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    END IF

    FoundAMtr = False

    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        CurrRead& = UBBillRec(1).CurRead(mChk)
        PrevRead& = UBBillRec(1).PrevRead(mChk)
        UsageAmt& = CurrRead& - PrevRead&
        IF UsageAmt& < 0 THEN
          MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(mChk))) - 1)
          UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(mChk)) + UBBillRec(1).CurRead(mChk)
        END IF
        FoundAMtr = True
        EXIT FOR
      END IF
    NEXT
    
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round(UBBillRec(1).TransAmt + Previous# + TaxTotal#)

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF

    IF TotalAmt# > 0 AND UBBillRec(1).TransAmt > 0 THEN
      TenPct# = Round#(UBBillRec(1).TransAmt * .1)
    ELSE
      TenPct# = 0
    END IF
    
    PRINT #UBRpt, "!"; USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(47); USING "ACCT: #####"; CustAcctNo&
    PRINT #UBRpt,

    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(1).RevName, 3);
      PRINT #UBRpt, TAB(5); USING "#########"; CurrRead&;
      PRINT #UBRpt, TAB(14); USING "#########"; PrevRead&;
      PRINT #UBRpt, TAB(23); USING "########"; UsageAmt&;
      PRINT #UBRpt, TAB(35); USING "######.##"; UBBillRec(1).RevAmt(1);
    END IF
    PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).CustName, 25)

    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(2).RevName, 3);
      PRINT #UBRpt, TAB(5); USING "#########"; CurrRead&;
      PRINT #UBRpt, TAB(14); USING "#########"; PrevRead&;
      PRINT #UBRpt, TAB(23); USING "########"; UsageAmt&;
      PRINT #UBRpt, TAB(35); USING "######.##"; UBBillRec(1).RevAmt(2);
    END IF
    PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).Addr1, 25)

    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(3).RevName, 3);
      PRINT #UBRpt, TAB(5); USING "#########"; CurrRead&;
      PRINT #UBRpt, TAB(14); USING "#########"; PrevRead&;
      PRINT #UBRpt, TAB(23); USING "########"; UsageAmt&;
      PRINT #UBRpt, TAB(35); USING "######.##"; UBBillRec(1).RevAmt(3);
    END IF
    PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).Addr2, 25)

    IF UBBillRec(1).RevAmt(4) <> 0 THEN
      PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(4).RevName, 3); TAB(35); USING "######.##"; UBBillRec(1).RevAmt(4);
    END IF
    PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode

    IF UBBillRec(1).RevAmt(5) <> 0 THEN
      PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(5).RevName, 3); TAB(35); USING "######.##"; UBBillRec(1).RevAmt(5);
    ELSE
      PRINT #UBRpt,
    END IF

    IF TotalTax# > 0 THEN
      PRINT #UBRpt, TAB(2); "TAX"; TAB(35); USING "######.##"; TotalTax#
    ELSE
      PRINT #UBRpt,
    END IF

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, TAB(2); "Less Deposit:"; TAB(35); USING "######.##"; -UBCustRec(1).DepositAmt
    ELSE
      PRINT #UBRpt,
    END IF

    'if
    'PRINT #UBRpt,

    PRINT #UBRpt, TAB(48); UBCustRec(1).Book; TAB(53); UBCustRec(1).SeqNumb;
    PRINT #UBRpt, TAB(61); LEFT$(PastDueDate$, 2); TAB(65); MID$(PastDueDate$, 4, 2); TAB(69); RIGHT$(PastDueDate$, 2)
    PRINT #UBRpt,
    
    PRINT #UBRpt, TAB(2); LEFT$(DateRead$, 2); TAB(7); MID$(DateRead$, 4, 2);
    PRINT #UBRpt, TAB(14); USING "#######.##"; UBBillRec(1).TransAmt;
    PRINT #UBRpt, USING "######.##"; TenPct#;
    PRINT #UBRpt, "  "; USING "######.##"; Round#(UBBillRec(1).TransAmt + TenPct#);

    PRINT #UBRpt, TAB(49); USING "#######.##"; UBBillRec(1).TransAmt;
    PRINT #UBRpt, "   "; USING "#######.##"; Round#(UBBillRec(1).TransAmt + TenPct#)

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, TAB(14); USING "#######.##"; -UBCustRec(1).DepositAmt;
      PRINT #UBRpt, TAB(34); USING "#######.##"; -UBCustRec(1).DepositAmt;
      PRINT #UBRpt, TAB(49); USING "#######.##"; -UBCustRec(1).DepositAmt;
      PRINT #UBRpt, "   "; USING "#######.##"; -UBCustRec(1).DepositAmt
    ELSE
      PRINT #UBRpt,
    END IF

    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(14); USING "#######.##"; Previous#;
      PRINT #UBRpt, TAB(34); USING "#######.##"; Previous#;
      PRINT #UBRpt, TAB(49); USING "#######.##"; Previous#;
      PRINT #UBRpt, "   "; USING "#######.##"; Previous#
    ELSE
      PRINT #UBRpt,
    END IF

    PRINT #UBRpt, TAB(14); USING "#######.##"; TotalAmt#;
    PRINT #UBRpt, TAB(34); USING "#######.##"; Round#(TotalAmt# + TenPct#);

    PRINT #UBRpt, TAB(49); USING "#######.##"; TotalAmt#;
    PRINT #UBRpt, "   "; USING "#######.##"; Round#(TotalAmt# + TenPct#)

    PRINT #UBRpt, Message$
    PRINT #UBRpt, "~"

