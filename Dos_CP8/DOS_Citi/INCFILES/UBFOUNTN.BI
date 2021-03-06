
    CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)
      
    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      'put printer in     8 lpi             12 cpi  oki mode
    END IF
                                                      ':  M
    DFoundMtr = False
    EFoundMtr = False
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
        CASE 4
          ECurrRead& = UBBillRec(1).CurRead(mChk)
          EPrevRead& = UBBillRec(1).PrevRead(mChk)
          EUsageAmt& = ECurrRead& - EPrevRead&
          IF EUsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(EPrevRead&)) - 1)
            EUsageAmt& = (MaxMeterAmt& - EPrevRead&) + ECurrRead&
          END IF
          EFoundMtr = True
        CASE 5
          DCurrRead& = UBBillRec(1).CurRead(mChk)
          DPrevRead& = UBBillRec(1).PrevRead(mChk)
          DUsageAmt& = DCurrRead& - DPrevRead&
          IF DUsageAmt& < 0 THEN
            MaxMeterAmt& = 10& ^ (LEN(STR$(DPrevRead&)) - 1)
            DUsageAmt& = (MaxMeterAmt& - DPrevRead&) + DCurrRead&
          END IF
          DFoundMtr = True

        END SELECT
      END IF
    NEXT

    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF

    IF TotalAmt# > 0 THEN
      TenPct# = 10
    ELSE
      TenPct# = 0
    END IF

    AcctNum = UBBillRec(1).CustAcctNo
    Acct$ = QPTrim$(STR$(AcctNum))
    SELECT CASE AcctNum
    CASE IS < 10
      Acct$ = "00" + Acct$
    CASE IS < 100
      Acct$ = "0" + Acct$
    END SELECT
    AcctLen = LEN(Acct$)
    
    PRINT #UBRpt, "~"; TAB(50); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,

    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, USING "##########"; UBBillRec(1).CustAcctNo;
    PRINT #UBRpt, TAB(15); LEFT$(UBCustRec(1).ServAddr, 26); TAB(50); USING "########"; UBBillRec(1).CustAcctNo;
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

    PRINT #UBRpt, TAB(50); Num2Date$(UBBillRec(1).PastDueDate);
    IF FinalFlag THEN
      PRINT #UBRpt, TAB(64); USING "#####.##"; Round#(TotalAmt#)
    ELSE
      PRINT #UBRpt, TAB(64); USING "#####.##"; Round#(TotalAmt# + TenPct#)
    END IF
    
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
'*******************************************
    PCnt = 0
    FOR WRevCnt = 1 TO 7
      PCnt = PCnt + 1
      IF UBBillRec(1).RevAmt(WRevCnt) <> 0 THEN
        PRINT #UBRpt, " "; LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 5);
        SELECT CASE PCnt
        CASE 1, 2  'water/sewer
          IF WFoundMtr THEN
            PRINT #UBRpt, TAB(7); USING "##########"; WCurrRead&;
            PRINT #UBRpt, TAB(17); USING "##########"; WPrevRead&;
            PRINT #UBRpt, TAB(28); USING "#######"; WUsageAmt&;
          END IF
        CASE 3
          IF EFoundMtr THEN    'electric
            PRINT #UBRpt, TAB(7); USING "##########"; ECurrRead&;
            PRINT #UBRpt, TAB(17); USING "##########"; EPrevRead&;
            PRINT #UBRpt, TAB(28); USING "#######"; EUsageAmt&;
          END IF
        CASE 4
          IF DFoundMtr THEN    'electric
            PRINT #UBRpt, TAB(7); USING "##########"; DCurrRead&;
            PRINT #UBRpt, TAB(17); USING "##########"; DPrevRead&;
            PRINT #UBRpt, TAB(28); USING "#######"; DUsageAmt&;
          END IF
        END SELECT
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
      CASE ELSE
        PRINT #UBRpt,
      END SELECT
    NEXT

'******************************************


    IF TotalTax# > 0 THEN
      PRINT #UBRpt, "                       TAX:  "; USING "$$,######.##"; TotalTax#
    ELSE
      PRINT #UBRpt, ""
    END IF
    
    IF Previous# <> 0 THEN
      PRINT #UBRpt, "                  Previous:  "; USING "$$,######.##"; Previous#;
    END IF
    PRINT #UBRpt, ' TAB(47); "LOC: "; UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb

    PRINT #UBRpt, "                   Current:  "; USING "$$,######.##"; UBBillRec(1).TransAmt
    PRINT #UBRpt, "                           --------------"

    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, "                   Deposit:  "; USING "$$,######.##"; -UBCustRec(1).DepositAmt
    ELSE
      PRINT #UBRpt,
    END IF

    IF TotalAmt# < 0 AND FinalFlag THEN
      PRINT #UBRpt, "                Refund Due:  "; USING "$$,######.##"; ABS(TotalAmt#)
    ELSE
      PRINT #UBRpt, "                     Total:  "; USING "$$,######.##"; TotalAmt#
    END IF
    
'    PRINT #UBRpt, TAB(3); CustMsg$

    PRINT #UBRpt, TAB(3); Message$;
    IF DidADraftFlag THEN
      PRINT #UBRpt, TAB(47); "DRAFT NOTICE DO NOT PAY!!"
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt, "~"

