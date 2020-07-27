'New Utility Bill format 03-31-97 BAR CODE PRINTABLE
'
    IF NOT LPIFlag THEN
      LPIFlag = -2
      'PRINT #UBRpt, CHR$(27); CHR$(48);
      'PRINT #UBRpt, CHR$(27); CHR$(58);
      PRINT #UBRpt, CHR$(27); CHR$(48);
      'PRINT #UBRpt, CHR$(27); CHR$(77);
      ' put printer in     8 lpi             12 cpi  oki mode
    END IF

   Ucode$ = qptrim$(UBCustRec(1).UserCode1)
   IF LEN(Ucode$) > 0 THEN
    Ucode = VAL(Ucode$)
   END IF
   
       
    AcctNum = UBBillRec(1).CustAcctNo
    Acct$ = qptrim$(STR$(AcctNum))
    SELECT CASE AcctNum
    CASE IS < 10
      Acct$ = "00" + Acct$
    CASE IS < 100
      Acct$ = "0" + Acct$
    END SELECT
    AcctLen = LEN(Acct$)

    Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
    TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

    IF FinalFlag AND CDeposit# THEN
      TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
    END IF

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
    PRINT #UBRpt, STRING$(50, " "); CHR$(27); CHR$(16); "A";
    PRINT #UBRpt, CHR$(8);
    PRINT #UBRpt, CHR$(2); CHR$(0);           '
    PRINT #UBRpt, CHR$(0); CHR$(2);           'Line 12
    PRINT #UBRpt, CHR$(1); CHR$(1);           '
    PRINT #UBRpt, CHR$(1); CHR$(2);
    PRINT #UBRpt, CHR$(27); CHR$(16); "B"; CHR$(AcctLen); Acct$

    IF Ucode > 0 AND Ucode < 3 THEN
     IF Ucode = 1 THEN
      PRINT #UBRpt, TAB(6); LEFT$(BillInfoRec(1).MsgLine4, 15); TAB(24); USING "#####.##"; Round#(UBBillRec(1).TransAmt)
    END IF
    IF Ucode = 2 THEN
     PRINT #UBRpt, TAB(6); LEFT$(BillInfoRec(1).MsgLine3, 15); TAB(24); USING "#####.##"; Round#(UBBillRec(1).TransAmt / 2)
     PRINT #UBRpt, TAB(6); LEFT$(BillInfoRec(1).MsgLine4, 15); TAB(24); USING "#####.##"; Round#(UBBillRec(1).TransAmt / 2)
    END IF
    ELSE
    PRINT #UBRpt, TAB(6); LEFT$(BillInfoRec(1).MsgLine2, 15); TAB(24); USING "#####.##"; Round#(UBBillRec(1).TransAmt / 3)
    PRINT #UBRpt, TAB(6); LEFT$(BillInfoRec(1).MsgLine3, 15); TAB(24); USING "#####.##"; Round#(UBBillRec(1).TransAmt / 3)
    PRINT #UBRpt, TAB(6); LEFT$(BillInfoRec(1).MsgLine4, 15); TAB(24); USING "#####.##"; Round#(UBBillRec(1).TransAmt / 3)
    END IF
    PRINT #UBRpt,
    IF Previous# <> 0 THEN
      PRINT #UBRpt, TAB(6); "Previous:     "; TAB(23); USING "######.##"; Previous#
    ELSE
      PRINT #UBRpt,
    END IF
    PRINT #UBRpt, TAB(47); USING "##########"; UBBillRec(1).CustAcctNo;
    PRINT #UBRpt, TAB(64); USING "#####.##"; TotalAmt#
    PRINT #UBRpt, TAB(6); "Current:      "; TAB(23); USING "######.##"; UBBillRec(1).TransAmt
    PRINT #UBRpt,

    PRINT #UBRpt,
    PRINT #UBRpt, TAB(58); Num2Date$(UBBillRec(1).PastDueDate)
    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, TAB(21); USING "$$,######.##"; TotalAmt#;
    PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).CustName, 29)
    PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).Addr1, 30)
    PRINT #UBRpt, TAB(47); LEFT$(UBCustRec(1).City, 14); " "; UBCustRec(1).State; " "; LEFT$(UBCustRec(1).ZipCode, 5)
    PRINT #UBRpt, STRING$(47, " ");

'    PRINT #UBRpt, CHR$(27); "[;;1}";      'Human readable on
'    PRINT #UBRpt, CHR$(27); "[50t"
'
'    PRINT #UBRpt, CHR$(2); CHR$(0);           '
'    PRINT #UBRpt, CHR$(0); CHR$(2);           'Line 12
'    PRINT #UBRpt, CHR$(1); CHR$(1);           '
'    PRINT #UBRpt, CHR$(1); CHR$(2);
'    Zip$ = QPTrim$(UBCustRec(1).ZipCode)
'    Ziplen = LEN(Zip$)
'    PRINT #UBRpt, CHR$(27); CHR$(16); "B"; CHR$(Ziplen); Zip$
'**************
'    PRINT #UBRpt, CHR$(27); CHR$(16); "A";
'    PRINT #UBRpt, CHR$(8);
'    PRINT #UBRpt, CHR$(2); CHR$(0);           '
'    PRINT #UBRpt, CHR$(0); CHR$(2);           'Line 12
'    PRINT #UBRpt, CHR$(1); CHR$(1);           '
'    PRINT #UBRpt, CHR$(1); CHR$(2);
'    Zip$ = QPTrim$(UBCustRec(1).ZipCode)
'    Ziplen = LEN(Zip$)
'    PRINT #UBRpt, CHR$(27); CHR$(16); "B"; CHR$(Ziplen); Zip$
'**************

    PRINT #UBRpt,
    PRINT #UBRpt,
    PRINT #UBRpt, ""

'    IF FinalFlag AND CDeposit# THEN
'      PRINT #UBRpt, "                   Deposit:  "; USING "$$,######.##"; -UBCustRec(1).DepositAmt
'    ELSE
'      PRINT #UBRpt,
'    END IF

'    IF TotalAmt# < 0 AND FinalFlag THEN
'      PRINT #UBRpt, "                Refund Due:  "; USING "$$,######.##"; TotalAmt#
'    ELSE
'      PRINT #UBRpt, "                     Total:  "; USING "$$,######.##"; TotalAmt#
'    END IF

    PRINT #UBRpt,
    PRINT #UBRpt,
'    STOP
    PRINT #UBRpt, "  "; Message$
    PRINT #UBRpt, "~"

