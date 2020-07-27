'New Landis Utility Bill format 02-24-97
'

    DidASenior = False
    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi
    END IF

    'FOR MCnt = 1 TO 7

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
    IF UBCustRec(1).PrevBalance <> 0 OR UBCustRec(1).CurrBalance <> 0 THEN
      PTotalAmt# = Round(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    ELSE
      PTotalAmt# = 0
    END IF
    TotalAmt# = Round(PTotalAmt# + UBBillRec(1).TransAmt)

    IF FinalFlag THEN
      PenAmt# = 0
    ELSE
      PenAmt# = Round#(UBBillRec(1).TransAmt * .05)
    END IF

    PRINT #UBRpt, TAB(50); Num2Date$(UBBillRec(1).BillDate); TAB(64); USING "#####.##"; TotalAmt#
    PRINT #UBRpt, TAB(3); Num2Date$(UBBillRec(1).BillDate); TAB(15); PrevDate$; TAB(26); DateRead$; TAB(40); USING "####"; DaysINRead
    PRINT #UBRpt, TAB(50); Num2Date$(UBBillRec(1).PastDueDate); TAB(64); USING "#####.##"; Round(TotalAmt# + PenAmt#)
    PRINT #UBRpt,
    PRINT #UBRpt,
    
    PRINT #UBRpt,

'*******************
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
          IF EFoundMtr THEN
            PRINT #UBRpt, TAB(7); USING "##########"; ECurrRead&;
            PRINT #UBRpt, TAB(17); USING "##########"; EPrevRead&;
            PRINT #UBRpt, TAB(28); USING "#######"; EUsageAmt&;
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
'***************


    OtherRev# = 0
    FOR ORevs = 9 TO 12
      OtherRev# = Round#(OtherRev# + UBBillRec(1).RevAmt(ORevs))
    NEXT
    IF OtherRev# <> 0 THEN
      PRINT #UBRpt, " MISC."; TAB(37); USING "#####.##"; OtherRev#
    ELSE
      PRINT #UBRpt,
    END IF
    IF FinalFlag AND Round#(UBCustRec(1).DepositAmt) = Round#(CDBL(UBBillRec(1).TaxAmt(15))) THEN
      DepositAmt# = UBCustRec(1).DepositAmt
      PRINT #UBRpt, "                   Deposit:  "; USING "$$,######.##"; UBCustRec(1).DepositAmt
    ELSE
      DepositAmt# = 0
      PRINT #UBRpt,
    END IF

'******** long view
    IF UBCustRec(1).PrevBalance <> 0 OR UBCustRec(1).CurrBalance <> 0 THEN
      'PTotalAmt# = Round(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
      IF PTotalAmt# <> 0 THEN
        PRINT #UBRpt, "                  Previous:  "; USING "$$,######.##"; PTotalAmt#
      ELSE
        PRINT #UBRpt,
      END IF
      PRINT #UBRpt, "                   Current:  "; USING "$$,######.##"; UBBillRec(1).TransAmt
      PRINT #UBRpt, "                           --------------"
    ELSE
      PRINT #UBRpt,
      PRINT #UBRpt, "                   Current:  "; USING "$$,######.##"; UBBillRec(1).TransAmt
      PRINT #UBRpt, "                           --------------"
    END IF

    TotalAmt# = Round((PTotalAmt# + UBBillRec(1).TransAmt) - DepositAmt#)

    IF TotalAmt# < 0 AND FinalFlag THEN
      PRINT #UBRpt, "                Refund Due:  "; USING "$$,######.##"; ABS(TotalAmt#)
    ELSE
      PRINT #UBRpt, "                     Total:  "; USING "$$,######.##"; TotalAmt#
    END IF


'here
    FOR SCnt = 1 TO 15
      SRateCode$ = QPTrim$(UBCustRec(1).Serv(SCnt).RateCode)
      SELECT CASE SRateCode$
      CASE "RS"
         IF UBBillRec(1).RevAmt(ElecRev) > 100.75 THEN
           RevFig# = 100.75
         ELSE
           RevFig# = UBBillRec(1).RevAmt(ElecRev)
         END IF
         DidASenior = True
      CASE "RWS"
         IF UBBillRec(1).RevAmt(ElecRev) > 87.66 THEN
           RevFig# = 87.66
         ELSE
           RevFig# = UBBillRec(1).RevAmt(ElecRev)
         END IF
         DidASenior = True
      CASE "RAS"
         IF UBBillRec(1).RevAmt(ElecRev) > 87.75 THEN
           RevFig# = 87.75
         ELSE
           RevFig# = UBBillRec(1).RevAmt(ElecRev)
         END IF
         DidASenior = True
      END SELECT
        
      IF DidASenior THEN
        RevDif# = (RevFig# * 1.0975) - RevFig#
        PRINT #UBRpt, "*Discount Applied:"; USING "###.##"; RevDif#;
        EXIT FOR
      END IF
    NEXT

    PRINT #UBRpt,
    
    'PRINT #UBRpt, TAB(3); Message$
    'PRINT #UBRpt, TAB(3); Msg2$

    PRINT #UBRpt, TAB(3); Message$;
    IF DidADraftFlag THEN
      PRINT #UBRpt, TAB(47); "DRAFT NOTICE DO NOT PAY!!"
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt, TAB(3); Msg2$;
    IF DidADraftFlag THEN
      PRINT #UBRpt, TAB(47); "DRAFT DATE: "; Num2Date$(BillInfoRec(1).DrftDate)
    ELSE
      PRINT #UBRpt,
    END IF
    
    PRINT #UBRpt, TAB(3); Msg3$
    IF (PTotalAmt# > 0) AND (NOT FinalFlag) THEN
      PRINT #UBRpt, TAB(3); Msg4$
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt, "~"

