
    IF UBCustRec(1).PrevBalance <> 0 OR UBCustRec(1).CurrBalance <> 0 THEN
      PTotalAmt# = Round(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    ELSE
      PTotalAmt# = 0
    END IF
    TotalAmt# = Round(PTotalAmt# + UBBillRec(1).TransAmt)

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF
    
    EFoundMtr = False
    WFoundMtr = False
    WCurrRead& = 0
    WPrevRead& = 0
    WUsageAmt& = 0
    ECurrRead& = 0
    EPrevRead& = 0
    EUsageAmt& = 0


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
        CASE 4, 5
          ECurrRead& = UBBillRec(1).CurRead(mChk)
          EPrevRead& = UBBillRec(1).PrevRead(mChk)
          EUsageAmt& = ECurrRead& - EPrevRead&
          IF EUsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(EPrevRead&)) - 1)
            EUsageAmt& = (MaxMeterAmt& - EPrevRead&) + ECurrRead&
          END IF
          EFoundMtr = True
        END SELECT
      END IF
    NEXT

         
     PRINT #UBRpt, FBillNO& + PrintedCnt
     PRINT #UBRpt, " "
     PRINT #UBRpt, TAB(3); DateRead$;
     PRINT #UBRpt, TAB(19); USING "######"; UBBillRec(1).CustAcctNo
     PRINT #UBRpt, " "
     PRINT #UBRpt, " "

     PRINT #UBRpt, TAB(5); USING "########"; WPrevRead&;
     PRINT #UBRpt, TAB(15); USING "########"; WCurrRead&;
     PRINT #UBRpt, TAB(24); USING "########"; WUsageAmt&


     PRINT #UBRpt, ""
     PRINT #UBRpt, TAB(5); USING "########"; EPrevRead&;
     PRINT #UBRpt, TAB(15); USING "########"; ECurrRead&;
     PRINT #UBRpt, TAB(24); USING "########"; EUsageAmt&;
     PRINT #UBRpt, TAB(36); LEFT$(UBCustRec(1).ServAddr, 23)
     PRINT #UBRpt, " "
     PRINT #UBRpt, " "
     IF UBBillRec(1).RevAmt(1) <> 0 THEN
       PRINT #UBRpt, TAB(3); "WATER"; TAB(24); USING "####,#.##"; UBBillRec(1).RevAmt(1);
     END IF
     PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).CustName, 25)
     IF UBBillRec(1).RevAmt(2) <> 0 THEN
       PRINT #UBRpt, TAB(3); "SEWER"; TAB(24); USING "####,#.##"; UBBillRec(1).RevAmt(2);
     END IF
     PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).Addr1, 25)
     IF UBBillRec(1).RevAmt(3) <> 0 THEN
       PRINT #UBRpt, TAB(3); "ELECTRIC"; TAB(24); USING "####,#.##"; UBBillRec(1).RevAmt(3);
     END IF
     PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).Addr2, 25)

     IF UBBillRec(1).RevAmt(6) <> 0 THEN
       PRINT #UBRpt, TAB(3); "SEC LIGHT"; TAB(24); USING "####,#.##"; UBBillRec(1).RevAmt(6);
     END IF
     PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)

     ECred# = Round#(UBBillRec(1).RevAmt(4) + UBBillRec(1).RevAmt(5))
     IF ECred# <> 0 THEN
       PRINT #UBRpt, TAB(3); "ELEC CREDIT"; TAB(24); USING "####,#.##"; ECred#;
     END IF
     IF FinalFlag AND CDeposit# THEN
       PRINT #UBRpt, TAB(3); "Deposit:  "; TAB(24); USING "####,#.##"; -UBCustRec(1).DepositAmt;
     END IF

     PRINT #UBRpt, TAB(35); STRING$(25, "-")
     PRINT #UBRpt, TAB(35); Message$
     PRINT #UBRpt, TAB(35); "Past Due on "; Num2Date$(UBBillRec(1).PastDueDate)
     PRINT #UBRpt, " "
     PRINT #UBRpt, " "
     PRINT #UBRpt, TAB(5); USING "####,#.##"; PTotalAmt#;
     PRINT #UBRpt, TAB(15); USING "#####.##"; UBBillRec(1).TransAmt;
     PRINT #UBRpt, TAB(24); USING "####,#.##"; TotalAmt#;
     PRINT #UBRpt, TAB(37); USING "#####"; UBBillRec(1).CustAcctNo;
     PRINT #UBRpt, TAB(49); USING "####,#.##"; TotalAmt#
     PRINT #UBRpt, " "

