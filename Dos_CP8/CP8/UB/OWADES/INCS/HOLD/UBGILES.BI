'New Utility Bill format 10-28-96 BAR CODE PRINTABLE

    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
    END IF

    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    IF UBCustRec(1).BillCycl = 5 THEN
      TenPercentAmount# = Round#(UBBillRec(1).TransAmt * .05)
    ELSE
      TenPercentAmount# = Round#(UBBillRec(1).TransAmt * .1)
    END IF

    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    IF TotalAmt# < 0 THEN
      TenPercentAmount# = 0
    END IF

    PRINT #UBRpt, TAB(50); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, USING "##########"; UBBillRec(1).CustAcctNo;
    PRINT #UBRpt, TAB(15); LEFT$(UBCustRec(1).ServAddr, 19); TAB(50); USING "########"; UBBillRec(1).CustAcctNo;
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

    PRINT #UBRpt, TAB(50); Num2Date$(UBBillRec(1).PastDueDate); TAB(64); USING "#####.##"; Round#(TenPercentAmount# + TotalAmt#)
    
    PRINT #UBRpt,
    PRINT #UBRpt,

    PRINT #UBRpt,
    PCnt = 0

    FOR WRevCnt = 1 TO 15
      PCnt = PCnt + 1
      IF UBBillRec(1).RevAmt(WRevCnt) <> 0 THEN
        PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 5);
         IF WRevCnt = 1 OR WRevCnt = 10 THEN
           FOR MMCnt = 1 TO 7
             MtrType = UBBillRec(1).MtrTypes(MMCnt)
             IF MtrType = 1 OR MtrType = 3 THEN
               IF UBBillRec(1).CurRead(MMCnt) > 0 THEN
                 UsageAmt& = UBBillRec(1).CurRead(MMCnt) - UBBillRec(1).PrevRead(MMCnt)
                 IF UsageAmt& < 0 THEN
                   MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(MMCnt))) - 1)
                   UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MMCnt)) + UBBillRec(1).CurRead(MMCnt)
                 END IF
                 PRINT #UBRpt, TAB(7); USING "#########"; UBBillRec(1).CurRead(MMCnt);
                 PRINT #UBRpt, TAB(18); USING "#########"; UBBillRec(1).PrevRead(MMCnt);
                 PRINT #UBRpt, TAB(27); USING "######"; UsageAmt&;
                 SELECT CASE UBCustRec(1).LocMeters(MMCnt).MTRMulti
                 CASE 10
                   Zero$ = "0"
                 CASE 100
                   Zero$ = "00"
                 CASE 1000
                   Zero$ = "000"
                 CASE ELSE
                   Zero$ = ""
                 END SELECT
                 PRINT #UBRpt, Zero$;
               END IF
             END IF
           NEXT
         END IF
         PRINT #UBRpt, TAB(36); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt);
      END IF
      SELECT CASE PCnt
      CASE 4
        PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).CustName, 29)
      CASE 5
        PRINT #UBRpt, TAB(47); UBCustRec(1).Addr1
      CASE 6
        PRINT #UBRpt, TAB(47); UBCustRec(1).Addr2
      CASE 7
        PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
      CASE 8, 9, 10
        PRINT #UBRpt,
      CASE ELSE
      END SELECT
    NEXT

    PRINT #UBRpt, ""

    IF Previous# <> 0 THEN
      PRINT #UBRpt, "                  Previous:  "; USING "$$,######.##"; Previous#
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt, "                   Current:  "; USING "$$,######.##"; UBBillRec(1).TransAmt
    PRINT #UBRpt, "                           --------------"

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, "                   Deposit:  "; USING "$$,######.##"; -UBCustRec(1).DepositAmt
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    ELSE
      PRINT #UBRpt,
    END IF

    IF TotalAmt# < 0 AND FinalFlag THEN
      PRINT #UBRpt, "                Refund Due:  "; USING "$$,######.##"; ABS(TotalAmt#)
    ELSE
      PRINT #UBRpt, "                     Total:  "; USING "$$,######.##"; TotalAmt#
    END IF

    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, "  "; QPTrim$(Message$)
    'PRINT Message$
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,

