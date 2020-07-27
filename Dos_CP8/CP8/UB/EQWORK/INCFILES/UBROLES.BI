'063097

   LPIFlag = -2
   PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
   ' put printer in     8 lpi      Epson FX Codes     

    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)

    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF

    IF TotalAmt# > 0 THEN
      FiveBucks# = 5
    ELSE
      FiveBucks# = 0
    END IF


    PRINT #UBRpt, TAB(24); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(49); USING "##########"; UBBillRec(1).CustAcctNo
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(49); Num2Date$(UBBillRec(1).BillDate)
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(4); USING "#####.##"; TotalAmt#; TAB(20); TotalAmt# + FiveBucks#;
    PRINT #UBRpt, TAB(30); LEFT$(UBCustRec(1).ServAddr, 25)
    PRINT #UBRpt,
    PRINT #UBRpt,

    PCnt = 0
    FOR WRevCnt = 1 TO 7
      IF UBBillRec(1).RevAmt(WRevCnt) <> 0 THEN
        PCnt = PCnt + 1
        PRINT #UBRpt, LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 3);
        IF UBBillRec(1).CurRead(WRevCnt) > 0 THEN
          UsageAmt& = UBBillRec(1).CurRead(WRevCnt) - UBBillRec(1).PrevRead(WRevCnt)
          IF UsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WRevCnt))) - 1)
            UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(WRevCnt)) + UBBillRec(1).CurRead(WRevCnt)
          END IF
          PRINT #UBRpt, TAB(6); DateRead$;
          PRINT #UBRpt, TAB(19); USING "##########"; UBBillRec(1).CurRead(WRevCnt);
          PRINT #UBRpt, TAB(34); USING "##########"; UBBillRec(1).PrevRead(WRevCnt);
          PRINT #UBRpt, TAB(47); USING "##########"; UsageAmt&;
        END IF
        PRINT #UBRpt, TAB(60); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt)
      END IF
    NEXT

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, "                   Deposit:  "; USING "$$,######.##"; -UBCustRec(1).DepositAmt
'      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    ELSE
      PRINT #UBRpt,
    END IF
         'was 8
    FOR zz = 7 TO PCnt STEP -1
      PRINT #UBRpt,
    NEXT

    PRINT #UBRpt, TAB(58); USING "$$,#####.##"; UBBillRec(1).TransAmt
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(58); USING "$$,#####.##"; Previous#
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(4); USING "#####.##"; TotalAmt#; TAB(20); TotalAmt# + FiveBucks#;
    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).CustName, 29)
    PRINT #UBRpt, TAB(35); UBCustRec(1).Addr1
    PRINT #UBRpt, TAB(35); UBCustRec(1).Addr2
    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)

    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, USING "##########"; AcctNum
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); Message$


'    IF TotalAmt# < 0 AND FinalFlag THEN
'      PRINT #UBRpt, "                Refund Due:  "; USING "$$,######.##"; ABS(TotalAmt#)
'    ELSE
'      PRINT #UBRpt, "                     Total:  "; USING "$$,######.##"; TotalAmt#
'    END IF

'    PRINT #UBRpt, TAB(50); Num2Date$(UBBillRec(1).PastDueDate); TAB(64); USING "#####.##"; ABS(TotalAmt#)
'    PRINT #UBRpt, TAB(3); Num2Date$(UBBillRec(1).BillDate); TAB(15); PrevDate$; TAB(26); DateRead$;
'     'Only Print Days if Greater than 0
'     IF DaysINRead > 0 THEN
'       PRINT #UBRpt, TAB(40); USING "####"; DaysINRead
'     ELSE
'       PRINT #UBRpt, ""
'     END IF
'
'    PRINT #UBRpt, TAB(50); Num2Date$(UBBillRec(1).PastDueDate); TAB(64); USING "#####.##"; TotalAmt#

'    PRINT #UBRpt,
'    PRINT #UBRpt,



'***********************************************************************
'051697 Montross, Mineral Springs
'062397 modified for Rolesville
'
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
'    PRINT #UBRpt, "~"; TAB(34); USING "#####"; FBillNO& + PrintedCnt
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt, TAB(3); UBCustRec(1).ServAddr
'    PRINT #UBRpt,
'    PRINT #UBRpt, TAB(2); LEFT$(BillDate$, 6); RIGHT$(BillDate$, 2); TAB(11); LEFT$(DateRead$, 6); RIGHT$(DateRead$, 2); TAB(20); USING "########"; CustAcctNo&
'    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).CustName, 25)
'    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr1, 25)
'    PRINT #UBRpt, TAB(2); USING "#########"; UBBillRec(1).PrevRead(1);
'    PRINT #UBRpt, TAB(12); USING "#########"; UBBillRec(1).CurRead(1);
'    PRINT #UBRpt, TAB(22); USING "########"; UsageAmt&;
'    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).Addr2, 25)
'    PRINT #UBRpt, TAB(34); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode
'
'    IF UBBillRec(1).RevAmt(1) <> 0 THEN
'      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(1).RevName; TAB(23); USING "#####.##"; UBBillRec(1).RevAmt(1);
'    END IF
'    PRINT #UBRpt,
'
'    IF UBBillRec(1).RevAmt(2) <> 0 THEN
'      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(2).RevName; TAB(23); USING "#####.##"; UBBillRec(1).RevAmt(2);
'    END IF
'    PRINT #UBRpt,
'
'    IF UBBillRec(1).RevAmt(3) <> 0 THEN
'      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(3).RevName; TAB(23); USING "#####.##"; UBBillRec(1).RevAmt(3);
'    END IF
'    PRINT #UBRpt,
'
'    IF UBBillRec(1).RevAmt(4) <> 0 THEN
'      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(4).RevName; TAB(23); USING "#####.##"; UBBillRec(1).RevAmt(4);
'    END IF
'
'    IF UBBillRec(1).RevAmt(5) <> 0 THEN
'      PRINT #UBRpt, TAB(3); UBSetUpRec(1).Revenues(5).RevName; TAB(23); USING "#####.##"; UBBillRec(1).RevAmt(5);
'    END IF
'    PRINT #UBRpt,
'    PRINT #UBRpt, TAB(34); LEFT$(PastDueDate$, 6); RIGHT$(BillDate$, 2); TAB(44); USING "########"; CustAcctNo&
'
'    'insert 6 here
'
'    IF Previous# <> 0 THEN
'      PRINT #UBRpt, TAB(3); "Previous:"; TAB(23); USING "#####.##"; Previous#
'    ELSE
'      PRINT #UBRpt, ""
'    END IF
'
'    IF FinalFlag AND CDeposit# THEN
'      PRINT #UBRpt, TAB(4); "Deposit:"; TAB(20); USING "#####.##"; -UBCustRec(1).DepositAmt
'      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
'    ELSE
'      PRINT #UBRpt,
'    END IF
'
'    PRINT #UBRpt, TAB(2); USING "#####.##"; TotalAmt#; TAB(16); TotalAmt# +FiveBucks#; TAB(24); TotalAmt#;
'    PRINT #UBRpt, USING "#####.##"; TAB(41); TotalAmt# +FiveBucks#; TAB(49); TotalAmt#
'    PRINT #UBRpt,
'    PRINT #UBRpt,
'    PRINT #UBRpt, Message$
'    PRINT #UBRpt, "~"

