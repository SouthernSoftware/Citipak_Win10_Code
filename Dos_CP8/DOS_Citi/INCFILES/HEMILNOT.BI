
    IF NOT LPIFlag THEN
      LPIFlag = -2
      PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(58);
      'PRINT #UBRpt, CHR$(27); CHR$(48); CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    END IF

    AcctNum = AcctNo&
    Acct$ = QPTrim$(STR$(AcctNum))
    SELECT CASE AcctNum
    CASE IS < 10
      Acct$ = "00" + Acct$
    CASE IS < 100
      Acct$ = "0" + Acct$
    END SELECT
    AcctLen = LEN(Acct$)

    Previous# = UBCustRec(1).PrevBalance
    Current# = UBCustRec(1).CurrBalance
    TotalAmt# = Round#(Previous# + Current#)

    PRINT #UBRpt, "~"  '; TAB(30); USING "########"; FBillNO& + PrintedCnt
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, USING "##########"; AcctNo&;
    PRINT #UBRpt, TAB(15); LEFT$(UBCustRec(1).ServAddr, 19); TAB(50); USING "########"; AcctNum;
    PRINT #UBRpt, TAB(62); NDate$
    PRINT #UBRpt,

    PRINT #UBRpt, TAB(50); PDate$; TAB(64); USING "#####.##"; ABS(TotalAmt#)
    PRINT #UBRpt, TAB(3); NDate$; 'TAB(15); PrevDate$; TAB(26); DateRead$;
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PCnt = 0
    FOR WRevCnt = 1 TO 7
      PCnt = PCnt + 1
      IF UBCustRec(1).CurrRevAmts(WRevCnt) <> 0 THEN
        PRINT #UBRpt, " "; QPTrim$(UBSetUpRec(1).Revenues(WRevCnt).RevName); ' LEFT$(UBSetUpRec(1).Revenues(WRevCnt).RevName, 10);
        PRINT #UBRpt, TAB(36); USING "#####.##"; UBCustRec(1).CurrRevAmts(WRevCnt);
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

    'Zip$ = QPTrim$(UBCustRec(1).ZipCode) + "@"
    'Ziplen = LEN(Zip$)
    'PRINT #UBRpt, STRING$(47, " "); CHR$(27); CHR$(10); CHR$(67); CHR$(10); Zip$
     PRINT #UBRpt,

    'Zip$ = QPTrim$(LEFT$(UBCustRec(1).ZipCode, 5))
    'Ziplen = LEN(Zip$)
    'PRINT #UBRpt, STRING$(47, " "); CHR$(27); CHR$(16); "A";
    'PRINT #UBRpt, CHR$(8);
    'PRINT #UBRpt, CHR$(2); CHR$(0);           '
    'PRINT #UBRpt, CHR$(0); CHR$(2);           'Line 12
    'PRINT #UBRpt, CHR$(1); CHR$(1);           '
    'PRINT #UBRpt, CHR$(1); CHR$(2);
    'PRINT #UBRpt, CHR$(27); CHR$(16); "B"; CHR$(Ziplen); Zip$

    IF Previous# <> 0 THEN
      PRINT #UBRpt, "                  Previous:  "; USING "$$,######.##"; Previous#
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt, "                   Current:  "; USING "$$,######.##"; Current#
    PRINT #UBRpt, "                           --------------"
    PRINT #UBRpt,
    PRINT #UBRpt, "                     Total:  "; USING "$$,######.##"; TotalAmt#

    PRINT #UBRpt, "  "; NoticeInfo.MsgLine1
    PRINT #UBRpt, "  "; NoticeInfo.MsgLine2
    PRINT #UBRpt, "  "; NoticeInfo.MsgLine3
    PRINT #UBRpt, "  "; NoticeInfo.MsgLine4
    PRINT #UBRpt, "~"


'  PRINT #UBRpt, ""
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, "                               TOWN OF HEMINGWAY"
'  PRINT #UBRpt, "                                  P.O. Box 968"
'  PRINT #UBRpt, "                               HEMINGWAY SC 29554"
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, TAB(5); LongNDate$
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, "    "; UBCustRec(1).CustName; TAB(60); "ACCOUNT #"; Acct$
'  PRINT #UBRpt, "    "; UBCustRec(1).ADDR1
'  PRINT #UBRpt, "    "; UBCustRec(1).ADDR2
'  PRINT #UBRpt, "    "; QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, "                                REMINDER NOTICE"
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, "    Our records indicate that your account is past due. Your past due"
'  PRINT #UBRpt, USING "    amount is $$#####,#.## If your bill is not paid by 5:00 p.m. on "; CustBal#
'  PRINT #UBRpt, "    "; LongPDate$; ", your water service will be discontinued, and"
'  PRINT #UBRpt, "    a fee of $25.00 will be added to your account regardless of whether"
'  PRINT #UBRpt, "    service is discontinued or not."
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, "    The past due bill and $25.00 fee must be paid in full before service"
'  PRINT #UBRpt, "    will be reconnected. No exceptions will be made."
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, "    A drop box is available for your convenience."
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt, "    Thank you,"
'  PRINT #UBRpt, "    Hemingway Town Hall"
'  PRINT #UBRpt, CHR$(12)

