'New Utility Bill format 10-28-96 BAR CODE PRINTABLE
'MUST SHOW BOTH METERS OR, TOTAL CONSUMPTION ON THIS BILL

    CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)
      
    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); 'CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    END IF
                                                      ':  M
    FoundAMtr = False
    
    PrevRead& = 0
    CurrRead& = 0
    UsageAmt& = 0

    FOR zz = 1 TO 7
      IF UBCustRec(1).LocMeters(zz).CurRead > 0 THEN
        DateRead$ = Num2Date$(UBCustRec(1).LocMeters(zz).CurDate)
        PrevDate$ = Num2Date$(UBCustRec(1).LocMeters(zz).PastDate)
        PrevRead& = UBCustRec(1).LocMeters(zz).PrevRead
        CurrRead& = UBCustRec(1).LocMeters(zz).CurRead
        UsageAmt& = CurrRead& - PrevRead&
        IF UsageAmt& < 0 THEN
          MaxMeterAmt& = 10& ^ (LEN(STR$(PrevRead&)) - 1)
          UsageAmt& = (MaxMeterAmt& - PrevRead&) + CurrRead&
        END IF
        EXIT FOR
      END IF
    NEXT
    
    Previous# = UBCustRec(1).PrevBalance
    Current# = UBCustRec(1).CurrBalance
    TotalAmt# = Round#(Previous# + Current#)
    
    TenPct# = 0
    
    PRINT #UBRpt, "!"
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    IF UBCustRec(1).CurrRevAmts(1) <> 0 THEN
      PRINT #UBRpt, TAB(22); USING "#####.##"; UBCustRec(1).CurrRevAmts(1)
    ELSE
      PRINT #UBRpt,
    END IF

    IF UBCustRec(1).CurrRevAmts(2) <> 0 THEN
      PRINT #UBRpt, TAB(22); USING "#####.##"; UBCustRec(1).CurrRevAmts(2);
    END IF
    PRINT #UBRpt, TAB(34); PDate$; TAB(46); USING "##########"; AcctNo&;
    PRINT #UBRpt, TAB(60); USING "#####.##"; TotalAmt#
    IF UBCustRec(1).CurrRevAmts(5) <> 0 THEN
      PRINT #UBRpt, TAB(22); USING "#####.##"; UBCustRec(1).CurrRevAmts(5)
    ELSE
      PRINT #UBRpt,
    END IF
    IF UBCustRec(1).CurrRevAmts(3) <> 0 THEN
      PRINT #UBRpt, TAB(22); USING "#####.##"; UBCustRec(1).CurrRevAmts(3);
    END IF
    PRINT #UBRpt, TAB(45); LEFT$(UBCustRec(1).ServAddr, 26)

    IF UBCustRec(1).CurrRevAmts(4) <> 0 THEN
      PRINT #UBRpt, TAB(22); USING "#####.##"; UBCustRec(1).CurrRevAmts(4)
    ELSE
      PRINT #UBRpt,
    END IF
    IF Previous# <> 0 THEN
      PRINT #UBRpt, "     Previous"; TAB(22); USING "#####.##"; Previous#;
    END IF
    PRINT #UBRpt, TAB(45); PrevDate$; TAB(60); DateRead$

    PRINT #UBRpt,
    
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(34); USING "########"; CurrRead&;
    PRINT #UBRpt, TAB(46); USING "########"; PrevRead&;
    PRINT #UBRpt, TAB(61); USING "#######"; UsageAmt&
    PRINT #UBRpt,
    PRINT #UBRpt, NoticeInfo.MsgLine1
    PRINT #UBRpt, NoticeInfo.MsgLine2
    PRINT #UBRpt, NoticeInfo.MsgLine3
    PRINT #UBRpt, NoticeInfo.MsgLine4
    
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(60); USING "##########"; AcctNo&
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(60); USING "######.##"; TotalAmt#
    PRINT #UBRpt,
    
    PRINT #UBRpt, TAB(10); LEFT$(UBCustRec(1).CustName, 29)
    PRINT #UBRpt, TAB(10); UBCustRec(1).Addr1
    PRINT #UBRpt, TAB(10); UBCustRec(1).Addr2
    Zip$ = QPTrim$(UBCustRec(1).ZipCode)
    IF LEN(Zip$) > 5 THEN
      Zip$ = LEFT$(Zip$, 5) + "-" + RIGHT$(Zip$, 4)
    END IF
    PRINT #UBRpt, TAB(10); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; Zip$
    PRINT #UBRpt,
    PRINT #UBRpt,
    
    'IF DidADraftFlag THEN
    '  PRINT #UBRpt, TAB(47); "DRAFT NOTICE DO NOT PAY!!"
    'ELSE
      PRINT #UBRpt, CustMsg$
    'END IF

    'IF DidADraftFlag THEN
    '  PRINT #UBRpt, TAB(47); "DRAFT DATE: "; Num2Date$(BillInfoRec(1).DrftDate)
    'ELSE
      PRINT #UBRpt,
    'END IF

    PRINT #UBRpt, "!"

