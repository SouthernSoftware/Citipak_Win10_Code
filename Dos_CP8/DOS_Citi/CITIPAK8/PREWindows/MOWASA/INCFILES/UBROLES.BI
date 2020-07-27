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
    ELSE
      PRINT #UBRpt,
    END IF

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

