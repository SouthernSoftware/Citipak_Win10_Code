'North Lenoir

    PrevRead& = 0
    CurrRead& = 0
    UsageAmt& = 0

    FOR zz = 1 TO 7
      IF UBCustRec(1).LocMeters(zz).CurRead > 0 THEN
        DateRead$ = Num2Date$(UBCustRec(1).LocMeters(zz).CurDate)
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

    PRINT #UBRpt, CHR$(27); CHR$(48); "~"
    PRINT #UBRpt, ""
    PRINT #UBRpt, ""
    PRINT #UBRpt, ""
    PRINT #UBRpt, ""
    PRINT #UBRpt, ""
    PRINT #UBRpt, TAB(2); LEFT$(UBCustRec(1).ServAddr, 28)
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); LEFT$(NDate$, 6); RIGHT$(NDate$, 2); TAB(11); LEFT$(DateRead$, 6); RIGHT$(DateRead$, 2); " "; LEFT$(MtrNum$, 6);
    PRINT #UBRpt, TAB(26); USING "######"; AcctNo&;
    PRINT #UBRpt, TAB(36); LEFT$(PDate$, 6); RIGHT$(PDate$, 2); TAB(49); USING "######"; AcctNo&

    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(2); USING "########"; PrevRead&;
    PRINT #UBRpt, TAB(12); USING "########"; CurrRead&;
    PRINT #UBRpt, TAB(22); USING "#######"; UsageAmt&;
    PRINT #UBRpt, TAB(43); USING "#####.##"; TotalAmt#; TAB(51); TotalAmt#

    PRINT #UBRpt,
    PRINT #UBRpt,

    FOR WRevCnt = 1 TO 4
      IF UBCustRec(1).CurrRevAmts(WRevCnt) <> 0 THEN
        PRINT #UBRpt, " "; QPTrim$(UBSetUpRec(1).Revenues(WRevCnt).RevName);
        PRINT #UBRpt, TAB(22); USING "#####.##"; UBCustRec(1).CurrRevAmts(WRevCnt)
      ELSE
        PRINT #UBRpt,
      END IF
    NEXT
    PRINT #UBRpt,

    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).CustName, 29)  'Line 23
    PRINT #UBRpt, TAB(35); UBCustRec(1).Addr1
    PRINT #UBRpt, TAB(35); UBCustRec(1).Addr2
    PRINT #UBRpt, TAB(35); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)

    PRINT #UBRpt, TAB(2); LEFT$(PDate$, 6); RIGHT$(PDate$, 2);
    PRINT #UBRpt, TAB(14); USING "#####.##"; TotalAmt#; TAB(23); TotalAmt#

    PRINT #UBRpt, TAB(35); "LOC# "; UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, " "; NoticeInfo.MsgLine1
    PRINT #UBRpt, " "; NoticeInfo.MsgLine2
    PRINT #UBRpt, " "; NoticeInfo.MsgLine3
    PRINT #UBRpt, "~"









'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt, "     North Lenoir Water Corp."
'  PRINT #UBRpt, "     P.O. Box 1564"
'  PRINT #UBRpt, "     Kinston, NC 28503"
'  PRINT #UBRpt,
'  PRINT #UBRpt, "     Forwarding & address correction requested."
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt, "     "; UBCustRec(1).CustName
'  PRINT #UBRpt, "     "; UBCustRec(1).ADDR1
'  PRINT #UBRpt, "     "; UBCustRec(1).ADDR2
'  PRINT #UBRpt, "     "; QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
'  PRINT #UBRpt,
'  PRINT #UBRpt, ""
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt, "-"; TAB(78); "-"
'  PRINT #UBRpt, TAB(5); LongNDate$
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, "    ACCOUNT: "; Acct$; TAB(40); "LOC: "; UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, ""
'  PRINT #UBRpt, "    YOUR ACCOUNT IS DUE FOR TWO MONTHS.  YOU MUST PAY THE TOTAL AMOUNT"
'  PRINT #UBRpt, "    DUE, AND A LATE FEE OF $15.00.  IF NOT PAID BY 10:00 AM ON THE DATE"
'  PRINT #UBRpt, "    BELOW, YOU WILL OWE AN ADDITIONAL CHARGE OF $20.00 AND YOUR SERVICE"
'  PRINT #UBRpt, "    WILL BE DISCONTINUED."
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt, TAB(32); "        DUE DATE:  "; PDate$
'  PRINT #UBRpt,
'  PRINT #UBRpt, TAB(32); "BEFORE 10:00 PAY:"; USING "#####.##"; TotalBal#
'  PRINT #UBRpt,
'  PRINT #UBRpt, TAB(32); " AFTER 10:00 PAY:"; USING "#####.##"; Round#(TotalBal# + 20)
'  PRINT #UBRpt,
'  PRINT #UBRpt,
'  PRINT #UBRpt, "    FOR YOUR CONVENIENCE, YOU MAY USE THE PAYMENT DROP LOCATED AT THE"
'  PRINT #UBRpt, "-"; TAB(78); "-"
'  PRINT #UBRpt, "    FRONT OF OUR OFFICE."
'  PRINT #UBRpt, CHR$(12)

