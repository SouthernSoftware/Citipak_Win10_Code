'Macclesfield NC '111797
'what a piece of shit!!!!

    Line4$ = SPACE$(4)
    Line1$ = SPACE$(30)
    Line7$ = SPACE$(7)

    BD$ = Num2Date$(UBBillRec(1).BillDate)
    DateBill$ = " " + LEFT$(BD$, 2) + MID$(BD$, 4, 2) + RIGHT$(BD$, 2)
    DateFrom$ = " " + LEFT$(PrevDate$, 2) + MID$(PrevDate$, 4, 2) + RIGHT$(PrevDate$, 2)
    DateTo$ = " " + LEFT$(DateRead$, 2) + MID$(DateRead$, 4, 2) + RIGHT$(DateRead$, 2)

    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF

    IF TotalAmt# > 0 THEN
      TenPct# = Round#(TotalAmt# + 10)
    END IF
    
    PRINT #UBRpt, "~"; TAB(30); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); QPTrim$(UBCustRec(1).ServAddr)
    PRINT #UBRpt,
    PRINT #UBRpt,

    LSET Line1$ = UBCustRec(1).Serv(1).RateCode

    IF LEN(QPTrim$(UBCustRec(1).Serv(2).RateCode)) > 0 THEN
      MID$(Line1$, 6, 4) = UBCustRec(1).Serv(2).RateCode
    ELSEIF LEN(QPTrim$(UBCustRec(1).Serv(3).RateCode)) > 0 THEN
      MID$(Line1$, 6, 4) = UBCustRec(1).Serv(3).RateCode
    END IF

    FOR MtrCnt = 1 TO 7
      SELECT CASE UBCustRec(1).LocMeters(MtrCnt).MtrType
      CASE "E", "D"
        RSET Line4$ = QPTrim$(STR$(UBCustRec(1).LocMeters(MtrCnt).MtrMulti))
        MID$(Line1$, 11, 4) = Line4$
        EXIT FOR
      CASE ELSE
      END SELECT
    NEXT
    DayIn$ = SPACE$(4)
    RSET DayIn$ = QPTrim$(STR$(DaysINRead))
    MID$(Line1$, 27, 4) = DayIn$
    
    PRINT #UBRpt, Line1$
    'line 8
    PRINT #UBRpt,
    
    LSET Line1$ = DateBill$ + DateFrom$ + DateTo$
    RSET Line4$ = QPTrim$(STR$(UBBillRec(1).CustAcctNo))
    MID$(Line1$, 27, 4) = Line4$
    PRINT #UBRpt, Line1$
    'line 10
    PRINT #UBRpt, TAB(42); Line4$; TAB(51); LEFT$(BillDate$, 6) + RIGHT$(BillDate$, 2)
    'line 11

    DidAMtr = False
    FOR MtrCnt = 1 TO 7
      SELECT CASE UBBillRec(1).MtrTypes(MtrCnt)
      CASE MtrElectric, MtrDemand
        DidAMtr = True
        Curr# = UBBillRec(1).CurRead(MtrCnt)
        Prev# = UBBillRec(1).PrevRead(MtrCnt)
        UsageAmt& = Curr# - Prev#
        IF UsageAmt& < 0 THEN
          MaxMeterAmt& = 10& ^ (LEN(STR$(Curr#)) - 1)
          UsageAmt& = (MaxMeterAmt& - Prev#) + Curr#
        END IF
        RSET Line7$ = QPTrim$(STR$(Prev#))
        PRINT #UBRpt, Line7$;
        RSET Line7$ = QPTrim$(STR$(Curr#))
        PRINT #UBRpt, Line7$;
        RSET Line7$ = QPTrim$(STR$(UsageAmt&))
        PRINT #UBRpt, Line7$; " ";
        PRINT #UBRpt, USING "E#####.##"; UBBillRec(1).RevAmt(1)
        EXIT FOR
      END SELECT
    NEXT
    IF NOT DidAMtr THEN
      PRINT #UBRpt,
    END IF

    DidAMtr = False
    FOR MtrCnt = 1 TO 7
      SELECT CASE UBBillRec(1).MtrTypes(MtrCnt)
      CASE MtrWaterOnly, MtrCombined
        DidAMtr = True
        Curr# = UBBillRec(1).CurRead(MtrCnt)
        Prev# = UBBillRec(1).PrevRead(MtrCnt)
        UsageAmt& = Curr# - Prev#
        IF UsageAmt& < 0 THEN
          MaxMeterAmt& = 10& ^ (LEN(STR$(Curr#)) - 1)
          UsageAmt& = (MaxMeterAmt& - Prev#) + Curr#
        END IF
        RSET Line7$ = QPTrim$(STR$(Prev#))
        PRINT #UBRpt, Line7$;
        RSET Line7$ = QPTrim$(STR$(Curr#))
        PRINT #UBRpt, Line7$;
        RSET Line7$ = QPTrim$(STR$(UsageAmt&))
        PRINT #UBRpt, Line7$; " ";
        PRINT #UBRpt, USING "W#####.##"; UBBillRec(1).RevAmt(2)
        EXIT FOR
      END SELECT
    NEXT
    IF NOT DidAMtr THEN
      PRINT #UBRpt,
    END IF

    FOR MtrCnt = 1 TO 7
      SELECT CASE UBBillRec(1).MtrTypes(MtrCnt)
      CASE MtrSewerOnly, MtrCombined
        DidAMtr = True
        Curr# = UBBillRec(1).CurRead(MtrCnt)
        Prev# = UBBillRec(1).PrevRead(MtrCnt)
        UsageAmt& = Curr# - Prev#
        IF UsageAmt& < 0 THEN
          MaxMeterAmt& = 10& ^ (LEN(STR$(Curr#)) - 1)
          UsageAmt& = (MaxMeterAmt& - Prev#) + Curr#
        END IF
        RSET Line7$ = QPTrim$(STR$(Prev#))
        PRINT #UBRpt, Line7$;
        RSET Line7$ = QPTrim$(STR$(Curr#))
        PRINT #UBRpt, Line7$;
        RSET Line7$ = QPTrim$(STR$(UsageAmt&))
        PRINT #UBRpt, Line7$; " ";
        PRINT #UBRpt, USING "S#####.##"; UBBillRec(1).RevAmt(3);
        EXIT FOR
      END SELECT
    NEXT
    PRINT #UBRpt, TAB(35); LEFT$(PastDueDate$, 6) + RIGHT$(PastDueDate$, 2); USING "#####.##"; Previous#; TotalAmt#
    PRINT #UBRpt,

    DidOne = False
    IF UBBillRec(1).RevAmt(4) <> 0 THEN
      PRINT #UBRpt, USING "GAR.  #####.##"; UBBillRec(1).RevAmt(4);
    END IF
    IF UBBillRec(1).RevAmt(5) <> 0 THEN
      PRINT #UBRpt, TAB(17); USING "SEC.  #####.##"; UBBillRec(1).RevAmt(5)
      DidOne = True
    END IF
    IF NOT DidOne THEN
      PRINT #UBRpt,
    END IF

    DidOne = False
    IF UBBillRec(1).RevAmt(6) <> 0 THEN
      PRINT #UBRpt, USING "WHC.  #####.##"; UBBillRec(1).RevAmt(6);
    END IF
    IF UBBillRec(1).RevAmt(7) <> 0 THEN
      PRINT #UBRpt, TAB(17); USING "ACC.  #####.##"; UBBillRec(1).RevAmt(7)
      DidOne = True
    END IF
    IF NOT DidOne THEN
      PRINT #UBRpt,
    END IF
    
    IF UBBillRec(1).RevAmt(8) <> 0 THEN
      PRINT #UBRpt, USING "PEN.  #####.##"; UBBillRec(1).RevAmt(8);
    END IF
    IF UBBillRec(1).RevAmt(9) <> 0 THEN
      PRINT #UBRpt, TAB(17); USING "MSC.  #####.##"; UBBillRec(1).RevAmt(9);
    END IF
    PRINT #UBRpt, TAB(33); LEFT$(UBCustRec(1).CustName, 26)
    IF FinalFlag AND CDeposit# THEN
      PRINT #UBRpt, USING "DEP.  #####.##"; -CDeposit#;
    END IF
    PRINT #UBRpt, TAB(33); UBCustRec(1).Addr1
    PRINT #UBRpt, TAB(33); UBCustRec(1).Addr2
    PRINT #UBRpt, USING "#####.##"; Previous#;
    PRINT #UBRpt, TAB(16); USING "####.##"; UBBillRec(1).TransAmt;
    PRINT #UBRpt, TAB(23); USING "######.##"; TotalAmt#;
    PRINT #UBRpt, TAB(33); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
    'line 21
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, "~"; Message$

