' Cashion   Laser Bills 3 per page
     
    FFFlag = True

    IF BPrnCnt = 3 THEN
      PRINT #UBRpt, CHR$(12);
      BPrnCnt = 0
    END IF

    WFoundMtr = False
    GFoundMtr = False

    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        SELECT CASE UBBillRec(1).MtrTypes(mChk)
        CASE 1, 2, 3, 7
          WCurrRead& = UBBillRec(1).CurRead(mChk)
          WPrevRead& = UBBillRec(1).PrevRead(mChk)
          WUsageAmt& = WCurrRead& - WPrevRead&
          IF WUsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(WPrevRead&)) - 1)
            WUsageAmt& = (MaxMeterAmt& - WPrevRead&) + WCurrRead&
          END IF
          WFoundMtr = True
        CASE 6
          GCurrRead& = UBBillRec(1).CurRead(mChk)
          GPrevRead& = UBBillRec(1).PrevRead(mChk)
          GUsageAmt& = GCurrRead& - GPrevRead&
          IF EUsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(GPrevRead&)) - 1)
            GUsageAmt& = (MaxMeterAmt& - GPrevRead&) + GCurrRead&
          END IF
          GFoundMtr = True
        END SELECT
      END IF
    NEXT

    'CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)
    
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round(UBBillRec(1).TransAmt + Previous#)
    
    IF BPrnCnt > 0 THEN
      PRINT #UBRpt,
      PRINT #UBRpt,
    ELSE
      PRINT #UBRpt, CHR$(27); "&l3a"; CHR$(27); "&l6D";
    END IF

    PRINT #UBRpt, USING "####"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(1); LEFT$(DateRead$, 2); TAB(6); MID$(DateRead$, 4, 2); TAB(10); RIGHT$(DateRead$, 2);
    PRINT #UBRpt, TAB(14); LEFT$(PastDueDate$, 2); TAB(19); MID$(PastDueDate$, 4, 2); TAB(24); RIGHT$(PastDueDate$, 2); TAB(37); PastDueDate$
    PRINT #UBRpt,
    PRINT #UBRpt,

    IF WFoundMtr THEN
      PRINT #UBRpt, USING "########"; WPrevRead&;
      PRINT #UBRpt, TAB(9); USING "########"; WCurrRead&;
      PRINT #UBRpt, TAB(18); USING "########"; WUsageAmt&;
      PRINT #UBRpt, " W";
    END IF
    
    PRINT #UBRpt, TAB(33); LEFT$(UBCustRec(1).ServAddr, 24)
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(32); LEFT$(UBCustRec(1).CustName, 25)

    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(1).RevName; TAB(18); USING "######.##"; UBBillRec(1).RevAmt(1);
    END IF
    PRINT #UBRpt, TAB(32); LEFT$(UBCustRec(1).Addr1, 25)

    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(2).RevName; TAB(18); USING "######.##"; UBBillRec(1).RevAmt(2);
    END IF
    PRINT #UBRpt, TAB(32); LEFT$(UBCustRec(1).Addr2, 25)
    
    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(3).RevName;
      PRINT #UBRpt, TAB(18); USING "######.##"; Round#(UBBillRec(1).RevAmt(3) + UBBillRec(1).TaxAmt(3));
    END IF
    PRINT #UBRpt, TAB(32); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode

    IF UBBillRec(1).RevAmt(4) <> 0 THEN
      PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(4).RevName;
      PRINT #UBRpt, TAB(18); USING "######.##"; Round#(UBBillRec(1).RevAmt(4) + UBBillRec(1).TaxAmt(4));
    END IF
    PRINT #UBRpt, TAB(32); STRING$(24, "-")
    
    IF UBBillRec(1).RevAmt(5) <> 0 THEN
      PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(5).RevName;
      PRINT #UBRpt, TAB(18); USING "######.##"; UBBillRec(1).RevAmt(5);
    END IF


    IF DidADraftFlag THEN
      PRINT #UBRpt, TAB(32); "Account Will Be Drafted"
    ELSE
      PRINT #UBRpt,
    END IF

'    IF UBBillRec(1).RevAmt(6) <> 0 THEN
'      PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(6).RevName;
'      PRINT #UBRpt, TAB(18); USING "######.##"; UBBillRec(1).RevAmt(6);
'    END IF

    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(2); "Previous:"; TAB(18); USING "######.##"; Previous#
    ELSE
      PRINT #UBRpt, ""
    END IF

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, TAB(2); "Deposit:"; TAB(18); USING "######.##"; -UBCustRec(1).DepositAmt
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    ELSE
      PRINT #UBRpt,
    END IF

    PRINT #UBRpt,

    PRINT #UBRpt, TAB(3); CustAcctNo&; TAB(18); USING "######.##"; TotalAmt#;
    PRINT #UBRpt, TAB(35); CustAcctNo&; TAB(45); USING "######.##"; TotalAmt#
    PRINT #UBRpt, TAB(2); Message$
    PRINT #UBRpt,
    'PRINT #UBRpt,
    BPrnCnt = BPrnCnt + 1


