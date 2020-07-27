'Hamlet Bill format 01-28-99 BAR CODE PRINTABLE

    Zip$ = QPTrim$(UBCustRec(1).ZipCode)
    PZip$ = Zip$
    PZip$ = LEFT$(PZip$, 5) + "-" + MID$(PZip$, 6)
    ZDigit$ = GetZipEDigit$(Zip$)
    Zip$ = Zip$ + ZDigit$
    CustMsg$ = QPTrim$(UBCustRec(1).BILLCMNT)
      
'    IF NOT LPIFlag THEN
'      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
'    END IF

'    AcctNum = UBCustRec(1).CustAcctNo
    Previous# = UBCustRec(1).PrevBalance
    Current# = UBCustRec(1).CurrBalance
    TotalAmt# = Round#(Previous# + Current#)

    Acct$ = QPTrim$(UBCustRec(1).ZipCode)
    AcctLen = LEN(Acct$)

    PRINT #UBRpt, "~"; TAB(50); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,

    MPCnt = 1
    PCnt = 0
    FOR WRevCnt = 1 TO 6
      PCnt = PCnt + 1
      IF UBCustRec(1).CurrRevAmts(WRevCnt) <> 0 THEN
        PRINT #UBRpt, LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 3);
      END IF
      IF UBCustRec(1).CurrRevAmts(WRevCnt) <> 0 THEN
        PRINT #UBRpt, TAB(33); USING "#####.##"; UBCustRec(1).CurrRevAmts(WRevCnt);
      END IF
      SELECT CASE PCnt
      CASE 1
        PRINT #UBRpt, TAB(44); USING "##########"; AcctNo&
      CASE 5
        PRINT #UBRpt, TAB(49); LEFT$(UBCustRec(1).ServAddr, 26)
      CASE ELSE
        PRINT #UBRpt,
      END SELECT
    NEXT
    
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(14); USING " Past Due:      $$,######.##"; TotalAmt#
    PRINT #UBRpt, TAB(45); NDate$; TAB(60); PDate$
    PRINT #UBRpt, TAB(2); NoticeInfo.MsgLine1
    PRINT #UBRpt, TAB(2); NoticeInfo.MsgLine2
    PRINT #UBRpt, TAB(2); NoticeInfo.MsgLine3
    PRINT #UBRpt, TAB(2); NoticeInfo.MsgLine4; TAB(34); USING "        $$,######.##"; TotalAmt#
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(22); LEFT$(UBCustRec(1).CustName, 29)
    PRINT #UBRpt, TAB(22); UBCustRec(1).Addr1
    PRINT #UBRpt, TAB(22); UBCustRec(1).Addr2
    PRINT #UBRpt, USING "##########"; AcctNo&;
    PRINT #UBRpt, TAB(22); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; PZip$
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, USING "#######.##"; TotalAmt#
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, USING "#######.##"; TotalAmt#;
    PRINT #UBRpt, TAB(22); CHR$(27); CHR$(16); "C"; CHR$(LEN(Zip$)); Zip$
    PRINT #UBRpt,
    PRINT #UBRpt, "~"

