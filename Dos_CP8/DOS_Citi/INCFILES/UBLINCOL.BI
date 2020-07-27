'DEFINT A-Z
    
    WFoundMtr = False

    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        SELECT CASE UBBillRec(1).MtrTypes(mChk)
        CASE 1, 2, 3, 7
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
    
    FoundAMtr = False
    FOR mChk = 1 TO 7
      IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
        FoundAMtr = True
        EXIT FOR
      END IF
    NEXT

    IF FoundAMtr = False THEN
      'if no metered services then adjust read dates to billdate
      'and billdate - 30
      DateRead$ = Num2Date$(UBBillRec(1).BillDate)
      PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
    END IF

    AcctNum = UBBillRec(1).CustAcctNo
    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF
    
    IF TotalAmt# > 0 THEN
      TenPct# = Round#(TotalAmt# * .1)
    ELSE
      TenPct# = 0
    END IF

'2800 REM alignment mask
'2810 REM first set to 12 cpi for bills
     REM i have set ibm codes as default, epson codes incl. but rem'd out
     REM LPRINT CHR$(27); CHR$(58); : REM ibm codes for 12 cpi printing
     'LPRINT CHR$(27); CHR$(77); : REM epson codes for 12 cpi printing

     PRINT #UBRpt, CHR$(27); CHR$(77); "~"; TAB(47); USING "########"; FBillNO& + PrintedCnt
     PRINT #UBRpt, "LINCOLN COUNTY WATER LINCOLN COUNTY WATER"
     PRINT #UBRpt, "P.O. BOX 340"; TAB(22); "P.O. BOX 340"
     PRINT #UBRpt, "LINCOLNTON, GA 30817"; TAB(22); "LINCOLNTON, GA 30817"
     PRINT #UBRpt,

     PRINT #UBRpt, USING "WATER     #####.##"; UBBillRec(1).RevAmt(1);
     PRINT #UBRpt, TAB(19); USING "##########"; WPrevRead&;
     PRINT #UBRpt, TAB(29); USING "##########"; WCurrRead&;
     PRINT #UBRpt, TAB(39); USING "#######"; WUsageAmt&;
     PRINT #UBRpt, TAB(52); USING "#####.##  WATER"; UBBillRec(1).RevAmt(1)
     PRINT #UBRpt,

     IF UBBillRec(1).RevAmt(3) <> 0 THEN
       PRINT #UBRpt, USING "MISC      #####.##"; UBBillRec(1).RevAmt(3);
       PRINT #UBRpt, TAB(52); USING "#####.##  MISC"; UBBillRec(1).RevAmt(3)
     ELSE
       PRINT #UBRpt,
     END IF
     IF UBBillRec(1).RevAmt(4) <> 0 THEN
       PRINT #UBRpt, USING "TAP FEE   #####.##"; UBBillRec(1).RevAmt(4);
       PRINT #UBRpt, TAB(52); USING "#####.##  TAP FEE"; UBBillRec(1).RevAmt(4)
     ELSE
       PRINT #UBRpt,
     END IF

     IF UBBillRec(1).RevAmt(6) <> 0 THEN
       PRINT #UBRpt, USING "SEWAGE    #####.##"; UBBillRec(1).RevAmt(6);
       PRINT #UBRpt, TAB(19); USING "##########"; WPrevRead&;
       PRINT #UBRpt, TAB(29); USING "##########"; WCurrRead&;
       PRINT #UBRpt, TAB(39); USING "#######"; WUsageAmt&;
       PRINT #UBRpt, TAB(52); USING "#####.##  SEWAGE"; UBBillRec(1).RevAmt(6)
     ELSE
       PRINT #UBRpt,
     END IF


     PRINT #UBRpt,
     PRINT #UBRpt, TAB(13); USING "#####.##"; Previous#; TAB(53); Previous#
     PRINT #UBRpt, TAB(13); USING "#####.##"; TotalAmt#; TAB(53); TotalAmt#
     PRINT #UBRpt, TAB(5); LEFT$(PastDueDate$, 5); TAB(13); USING "#####.##"; Round#(TotalAmt# + TenPct#);
     PRINT #UBRpt, TAB(27); LEFT$(PastDueDate$, 5); TAB(34); "10%"; TAB(53); USING "#####.##"; Round#(TotalAmt# + TenPct#)
     PRINT #UBRpt,
     PRINT #UBRpt, TAB(8); USING "######"; AcctNum; TAB(29); AcctNum;
     PRINT #UBRpt, TAB(49); UBCustRec(1).ServAddr
     PRINT #UBRpt, TAB(9); BillDate$
     PRINT #UBRpt, TAB(9); PastDueDate$;
     PRINT #UBRpt, TAB(41); LEFT$(UBCustRec(1).CustName, 30)
     PRINT #UBRpt, TAB(41); UBCustRec(1).Addr1
     PRINT #UBRpt, Message$; TAB(41); UBCustRec(1).Addr2
     PRINT #UBRpt, Msg2$; TAB(41); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
     PRINT #UBRpt, Msg3$
     PRINT #UBRpt, Msg4$
     PRINT #UBRpt, "~"

