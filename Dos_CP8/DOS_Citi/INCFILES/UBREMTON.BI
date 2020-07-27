  WCurrRead& = 0
  WPrevRead& = 0
  WUsageAmt& = 0
  WFoundMtr = False

  FOR mChk = 1 TO 7
    IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
      SELECT CASE UBBillRec(1).MtrTypes(mChk)
      CASE 1, 2, 3
        WCurrRead& = UBBillRec(1).CurRead(mChk)
        WPrevRead& = UBBillRec(1).PrevRead(mChk)
        WUsageAmt& = WCurrRead& - WPrevRead&
        IF WUsageAmt& < 0 THEN
          MaxMeterAmt& = 10& ^ (LEN(STR$(WPrevRead&)) - 1)
          WUsageAmt& = (MaxMeterAmt& - WPrevRead&) + WCurrRead&
        END IF
        WFoundMtr = True
      END SELECT
    END IF
  NEXT

  AcctNum = UBBillRec(1).CustAcctNo
  Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)
  IF FinalFlag AND CDeposit# THEN
    TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
  END IF

  PRINT #UBRpt, "~"; TAB(17); USING "########"; FBillNO& + PrintedCnt
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(6); DateRead$;
  PRINT #UBRpt, TAB(19); AcctNum; TAB(44); LEFT$(UBCustRec(1).CustName, 16)
  PRINT #UBRpt, TAB(44); LEFT$(UBCustRec(1).Addr1, 16)
  PRINT #UBRpt, TAB(6); USING "########"; WPrevRead&; TAB(16); WCurrRead&; TAB(26); WUsageAmt&;
  PRINT #UBRpt, TAB(35); USING "####.##"; UBBillRec(1).RevAmt(1);
  PRINT #UBRpt, TAB(44); LEFT$(UBCustRec(1).City, 14)
  PRINT #UBRpt, TAB(35); USING "####.##"; UBBillRec(1).RevAmt(2);
  PRINT #UBRpt, TAB(44); UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
  PRINT #UBRpt, TAB(2); LEFT$(UBCustRec(1).CustName, 25);
  PRINT #UBRpt, TAB(35); USING "####.##"; Previous#
  PRINT #UBRpt, TAB(35); USING "####.##"; 0
  PRINT #UBRpt, TAB(2); Message$; TAB(35); USING "####.##"; UBBillRec(1).RevAmt(6)
  PRINT #UBRpt, TAB(2); Msg2$; TAB(34); USING "#####.##"; TotalAmt#; TAB(51); TotalAmt#

  IF FinalFlag AND CDeposit# THEN
    PRINT #UBRpt, TAB(2); "Less Deposit:"; USING "####.##"; UBCustRec(1).DepositAmt
  ELSE
    PRINT #UBRpt, TAB(2); Msg3$
  END IF

  PRINT #UBRpt,
  PRINT #UBRpt, TAB(32); PastDueDate$; TAB(44); "Acct No: "; AcctNum
  PRINT #UBRpt,
  PRINT #UBRpt, "~"

