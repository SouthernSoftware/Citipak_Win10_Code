
    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        MtrAT = mChk
        FoundAMtr = True
        EXIT FOR
      END IF
    NEXT

    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    'FinalFlag = True
    'CDeposit# = -1
    'UBCustRec(1).DepositAmt = 50

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF

    PRINT #UBRpt, "~"; TAB(30); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    
    PCnt = 0
    FOR WRevCnt = 1 TO 7
      PCnt = PCnt + 1
      IF UBBillRec(1).RevAmt(WRevCnt) <> 0 THEN
        IF WRevCnt < 3 THEN
          UsageAmt& = UBBillRec(1).CurRead(MtrAT) - UBBillRec(1).PrevRead(MtrAT)
          IF UsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(MtrAT))) - 1)
            UsageAmt& = (MaxMeterAmt& - UBBillRec(1).PrevRead(MtrAT)) + UBBillRec(1).CurRead(MtrAT)
          END IF
          PRINT #UBRpt, TAB(1); USING "#########"; UBBillRec(1).PrevRead(MtrAT);
          PRINT #UBRpt, TAB(10); USING "#########"; UBBillRec(1).CurRead(MtrAT);
          PRINT #UBRpt, TAB(19); USING "#######"; UsageAmt&;
        END IF
        PRINT #UBRpt, TAB(27); LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 2);
        PRINT #UBRpt, TAB(29); USING "#####.##"; UBBillRec(1).RevAmt(WRevCnt);
      END IF

      SELECT CASE PCnt
      CASE 2
        PRINT #UBRpt, TAB(40); LEFT$(UBCustRec(1).CustName, 29)
      CASE 3
        PRINT #UBRpt, TAB(40); UBCustRec(1).Addr1
      CASE 4
        PRINT #UBRpt, TAB(40); UBCustRec(1).Addr2
      CASE 5
        PRINT #UBRpt, TAB(40); QPTrim$(LEFT$(UBCustRec(1).City, 14)); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
      CASE 6
        IF FinalFlag AND CDeposit# THEN
          PRINT #UBRpt, TAB(19); "Deposit:"; TAB(29); USING "#####.##"; -UBCustRec(1).DepositAmt
        ELSE
          PRINT #UBRpt,
        END IF
      CASE ELSE
        PRINT #UBRpt,
        'PRINT #UBRpt, TAB(40); LEFT$(UBCustRec(1).ServAddr, 29)
      END SELECT
    NEXT

    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(9); DateRead$; TAB(27); Num2Date$(UBBillRec(1).BillDate)
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); USING "#####"; AcctNum;
    PRINT #UBRpt, TAB(10); USING "#####.##"; Previous#; TAB(19); UBBillRec(1).TransAmt; TAB(29); TotalAmt#;
    PRINT #UBRpt, TAB(43); USING "#####"; AcctNum;
    PRINT #UBRpt, TAB(51); USING "#####.##"; TotalAmt#
    PRINT #UBRpt,
    PRINT #UBRpt, Message$




