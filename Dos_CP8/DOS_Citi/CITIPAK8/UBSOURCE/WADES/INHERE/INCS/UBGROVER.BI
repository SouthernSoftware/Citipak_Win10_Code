'Grover's Utility Bill format 03-26-97
    
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    TenPercentAmount# = Round#(TotalAmt# * .1)

    TenPercentAmount# = Round#(UBBillRec(1).TransAmt * .1)

    WRevCnt = 1

    UsageAmt& = UBBillRec(1).CurRead(WRevCnt) - UBBillRec(1).PrevRead(WRevCnt)
    IF UsageAmt& < 0 THEN
      MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WRevCnt))) - 1)
      UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WRevCnt)) + UBBillRec(1).CurRead(WRevCnt)
    END IF

    IF TotalAmt# < 0 THEN
      TenPercentAmount# = 0
    END IF

     PRINT #UBRpt, "~"
     PRINT #UBRpt,
     PRINT #UBRpt,
     PRINT #UBRpt,
     PRINT #UBRpt,
     PRINT #UBRpt, TAB(2); LEFT$(UBCustRec(1).ServAddr, 25)   'line 5 is serv address
     PRINT #UBRpt,
     'PRINT #UBRpt, TAB(2); Num2Date$(UBBillRec(1).BillDate);    ' line 7 is billdate and meter read date
     PRINT #UBRpt, TAB(2); LEFT$(BillDate$, 6); RIGHT$(BillDate$, 2); ' line 7 is billdate and meter read date
     'PRINT #UBRpt, TAB(13); DateRead$; "   "; USING "#####"; UBBillRec(1).CustAcctNo
     PRINT #UBRpt, TAB(11); LEFT$(DateRead$, 6); RIGHT$(DateRead$, 2); USING "#########"; UBBillRec(1).CustAcctNo
     PRINT #UBRpt,
     PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).CustName, 25)
     PRINT #UBRpt, TAB(2); USING "########"; UBBillRec(1).PrevRead(WRevCnt);
     PRINT #UBRpt, TAB(12); USING "#########"; UBBillRec(1).CurRead(WRevCnt);
     PRINT #UBRpt, TAB(22); USING "######"; UsageAmt&;
     PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr1, 25)
     PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr2, 25)
     IF UBBillRec(1).RevAmt(1) <> 0 THEN
        PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 14);
        PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt);
     END IF
        PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
     WRevCnt = 2
     IF UBBillRec(1).RevAmt(2) <> 0 THEN
       PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 14);
       PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt)
     ELSE
       PRINT #UBRpt,
     END IF

     WRevCnt = 3
     IF UBBillRec(1).RevAmt(3) <> 0 THEN
       PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 14);
       PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt)
     ELSE
       PRINT #UBRpt,
     END IF

     WRevCnt = 4
     IF UBBillRec(1).RevAmt(4) <> 0 THEN
       PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 14);
       PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt)
     ELSE
       PRINT #UBRpt,
     END IF

'     WRevCnt = 5
'     IF UBBillRec(1).RevAmt(5) <> 0 THEN
'       PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 14);
'       PRINT #UBRpt, TAB(22); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt);
'     END IF
'     PRINT #UBRpt, TAB(47); USING "#####"; UBBillRec(1).CustAcctNo
'
'    IF Previous# <> 0 THEN
'      PRINT #UBRpt, " Balance Forward"; TAB(22); USING "#####.##"; Previous#
'    ELSE
'      PRINT #UBRpt,
'    END IF

    IF Previous# <> 0 THEN
      PRINT #UBRpt, " Balance Forward     "; ; USING "#####.##"; Previous#;
    END IF
    PRINT #UBRpt, TAB(47); USING "#####"; UBBillRec(1).CustAcctNo

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, "         Deposit     "; USING "#####.##"; -UBCustRec(1).DepositAmt
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    ELSE
      PRINT #UBRpt,
    END IF
    
    PRINT #UBRpt,
    IF TotalAmt# < 0 THEN
      PRINT #UBRpt, TAB(17); USING "####.##"; TotalAmt#;
    ELSE
      PRINT #UBRpt, TAB(17); USING "####.##"; TotalAmt# + TenPercentAmount#;
    END IF
    PRINT #UBRpt, TAB(25); USING "####.##"; TotalAmt#;

    IF TotalAmt# < 0 AND FinalFlag THEN
      PRINT #UBRpt, TAB(42); USING "####.##"; TotalAmt#;
    ELSE
      PRINT #UBRpt, TAB(42); USING "####.##"; TotalAmt# + TenPercentAmount#;
    END IF
    PRINT #UBRpt, TAB(50); USING "####.##"; TotalAmt#
    IF TotalAmt# < 0 AND FinalFlag THEN
      PRINT #UBRpt, "       Refund Due     "; USING "$$####.##"; ABS(TotalAmt#)
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(34); Message$
    PRINT #UBRpt, TAB(34); USING "########"; FBillNO& + PrintedCnt

 
' REFUND NOT ACTIVE ON THIS BILL FOR GROVER AT THIS TIME 4-23-97 NJP

'    IF FinalFlag AND CDeposit# THEN
'      PRINT #UBRpt, "                   Deposit:  "; USING "$$,######.##"; -UBCustRec(1).DepositAmt
'      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
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

