'111198

    PRINT #UBRpt, CHR$(27); CHR$(58);
    'switch to 12pitch

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
    PRINT #UBRpt, TAB(2); LEFT$(UBCustRec(1).CustName, 30)
    PRINT #UBRpt, TAB(44); LEFT$(UBCustRec(1).CustName, 25)
    PRINT #UBRpt, TAB(44); LEFT$(UBCustRec(1).Addr1, 25)
    PRINT #UBRpt, TAB(44); LEFT$(UBCustRec(1).Addr2, 25)
    PRINT #UBRpt, TAB(2); DateRead$; TAB(16); USING "#####"; UBBillRec(1).CustAcctNo; TAB(23); UBBillRec(1).CustAcctNo;
    PRINT #UBRpt, TAB(31); UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb;
    PRINT #UBRpt, TAB(44); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode
    PRINT #UBRpt,

    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(1).RevName, 2);
      IF UBSetUpRec(1).Revenues(1).UseMtr = "Y" THEN
        PRINT #UBRpt, TAB(4); USING "#########"; CurrRead&;
        PRINT #UBRpt, TAB(13); USING "#########"; PrevRead&;
        PRINT #UBRpt, TAB(23); USING "#######"; UsageAmt&;
      END IF
      PRINT #UBRpt, TAB(31); USING "######.##"; UBBillRec(1).RevAmt(1);
    END IF
    PRINT #UBRpt, TAB(50); USING "#####"; UBBillRec(1).CustAcctNo;
    PRINT #UBRpt, TAB(58); UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb

    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(2).RevName, 2);
      IF UBSetUpRec(1).Revenues(2).UseMtr = "Y" THEN
        PRINT #UBRpt, TAB(4); USING "#########"; CurrRead&;
        PRINT #UBRpt, TAB(13); USING "#########"; PrevRead&;
        PRINT #UBRpt, TAB(23); USING "#######"; UsageAmt&;
      END IF
      PRINT #UBRpt, TAB(31); USING "######.##"; UBBillRec(1).RevAmt(2);
    END IF
    PRINT #UBRpt,

    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(3).RevName, 2);
      IF UBSetUpRec(1).Revenues(3).UseMtr = "Y" THEN
        PRINT #UBRpt, TAB(4); USING "#########"; CurrRead&;
        PRINT #UBRpt, TAB(13); USING "#########"; PrevRead&;
        PRINT #UBRpt, TAB(23); USING "#######"; UsageAmt&;
      END IF
      PRINT #UBRpt, TAB(31); USING "######.##"; UBBillRec(1).RevAmt(3);
    END IF
    PRINT #UBRpt, TAB(62); PastDueDate$

    IF UBBillRec(1).RevAmt(4) <> 0 THEN
      PRINT #UBRpt, TAB(2); LEFT$(UBSetUpRec(1).Revenues(4).RevName, 2);
      PRINT #UBRpt, TAB(31); USING "######.##"; UBBillRec(1).RevAmt(4);
    END IF
    PRINT #UBRpt,
    
    IF TotalTax# > 0 THEN
      PRINT #UBRpt, TAB(2); "TAX"; TAB(31); USING "######.##"; TotalTax#;
    END IF
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(31); USING "######.##"; UBBillRec(1).TransAmt; TAB(43); UBBillRec(1).TransAmt; TenPct#; TAB(63); Round#(UBBillRec(1).TransAmt + TenPct#)
    PRINT #UBRpt,
    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(31); USING "######.##"; Previous#; TAB(43); Previous#; TAB(63); Previous#;
    END IF
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(31); USING "######.##"; TotalAmt#; TAB(43); TotalAmt#; TAB(63); Round#(TotalAmt# + TenPct#)
    PRINT #UBRpt,
    PRINT #UBRpt, Message$

