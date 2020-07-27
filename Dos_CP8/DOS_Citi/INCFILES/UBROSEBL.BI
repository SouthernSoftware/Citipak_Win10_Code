
'***********************
    WFoundMtr = False
    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        SELECT CASE UBBillRec(1).MtrTypes(mChk)
        CASE 1, 2, 3
          CurrRead& = UBBillRec(1).CurRead(mChk)
          PrevRead& = UBBillRec(1).PrevRead(mChk)
          UsageAmt& = CurrRead& - PrevRead&
          IF UsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(PrevRead&)) - 1)
            UsageAmt& = (MaxMeterAmt& - PrevRead&) + CurrRead&
          END IF
          WFoundMtr = True
          EXIT FOR
        END SELECT
      END IF
    NEXT

    'BCnt# = Value#(BillInfoRec(1).MsgLine2, ECode)
    'BCnt# = BCnt# + (PrintedCnt - 1)

     IF WFoundMtr = False THEN
      'if no metered services then adjust read dates to billdate
      'and billdate - 30
      'STOP
      DateRead$ = Num2Date$(UBBillRec(1).BillDate)
      PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
    END IF

    Loc$ = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb

'***********************
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round(UBBillRec(1).TransAmt + Previous#)

    PRINT #UBRpt, "~"; TAB(26); USING "#####"; FBillNO& + PrintedCnt
    PRINT #UBRpt, TAB(15); LEFT$(UBCustRec(1).CustName, 20)
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(15); Num2Date$(UBBillRec(1).BillDate)
    PRINT #UBRpt, TAB(24); USING "########"; CustAcctNo&
    PRINT #UBRpt,
    IF WFoundMtr THEN
      PRINT #UBRpt, TAB(10); USING "#########"; CurrRead&
      PRINT #UBRpt, TAB(10); USING "#########"; PrevRead&
      PRINT #UBRpt, TAB(17); USING "########"; UsageAmt&;
    ELSE
      PRINT #UBRpt,
      PRINT #UBRpt,
    END IF
    IF UBBillRec(1).RevAmt(1) <> 0 THEN
      PRINT #UBRpt, TAB(26); USING "######.##"; UBBillRec(1).RevAmt(1);
    END IF
    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).CustName, 25)

    IF UBBillRec(1).RevAmt(2) <> 0 THEN
      PRINT #UBRpt, TAB(26); USING "######.##"; UBBillRec(1).RevAmt(2);
    END IF
    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).Addr1, 25)
    
    IF UBBillRec(1).RevAmt(3) <> 0 THEN
      PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(3).RevName;
      PRINT #UBRpt, TAB(26); USING "######.##"; UBBillRec(1).RevAmt(3);
    END IF
    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).Addr2, 25)

    IF UBBillRec(1).RevAmt(5) <> 0 THEN
      PRINT #UBRpt, TAB(2); UBSetUpRec(1).Revenues(5).RevName;
      PRINT #UBRpt, TAB(26); USING "######.##"; UBBillRec(1).RevAmt(5);
    END IF

    PRINT #UBRpt, TAB(37); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; UBCustRec(1).ZipCode
    
    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, " Deposit"; TAB(26); USING "######.##"; -UBCustRec(1).DepositAmt;
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
      PRINT #UBRpt, 'TAB(50); USING "######.##"; TotalAmt#
    ELSE
      PRINT #UBRpt, 'TAB(50); USING "######.##"; TotalAmt#
    END IF
    
    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(26); USING "######.##"; Previous#;
    END IF

    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(26); USING "######.##"; TotalAmt#
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(37); USING "AC: #####"; CustAcctNo&;
    PRINT #UBRpt, TAB(2); "Meter Read - "; DateRead$; TAB(37); Loc$;
    PRINT #UBRpt, TAB(51); USING "######.##"; TotalAmt#
    PRINT #UBRpt, "~"

