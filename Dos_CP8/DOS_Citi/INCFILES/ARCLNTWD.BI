    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(60); "Date: "; DATE$
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, "NOTICE FOR RENEWAL OF BUSINESS LICENSE FOR PERIOD ENDING: DECEMBER, "; Form$(2, 0)
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(5); "Business Account # "; ARCustIdxRec(1).IDXRECORD
    PRINT #RptHandle, TAB(7); ARCustRec(1).BILLNAME
    PRINT #RptHandle, TAB(7); ARCustRec(1).Address1
    PRINT #RptHandle, TAB(7); ARCustRec(1).Address2
    PRINT #RptHandle, TAB(7); RTRIM$(ARCustRec(1).City); " " + RTRIM$(ARCustRec(1).State) + " " + RTRIM$(ARCustRec(1).ZipCode)
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, STRING$(79, "-")
    PRINT #RptHandle, TAB(2); "Code"; TAB(10); "Type of License"
    PRINT #RptHandle, STRING$(79, "-")

    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT1)
    GOSUB GetCode

    PRINT #RptHandle, TAB(2); ARCustRec(1).BILLCAT1;
    PRINT #RptHandle, TAB(7); ARCustRec(1).Desc1;
    PRINT #RptHandle, TAB(50); "BASIS AMOUNT"; TAB(66); "LICENSE AMOUNT"
    IF CodeType$ = "S" THEN
      PRINT #RptHandle, "Minimum Amount Due: "; USING "$$#####.##"; BaseAmt1#
      PRINT #RptHandle, "For Receipts Up To: "; USING "$$######,#.##"; Revenue1#
      PRINT #RptHandle, "Plus "; USING "##.###%  "; Percent1#;
      PRINT #RptHandle, "  Of Receipts Over "; USING "######"; Maximum1#
    ELSE
      PRINT #RptHandle,
      PRINT #RptHandle,
    END IF
    IF CodeType$ = "F" THEN
      PRINT #RptHandle, TAB(50); "Flat Fee: "; TAB(66); USING "######,#.##"; Amt#
    ELSE
      PRINT #RptHandle, TAB(50); "____________"; TAB(66); "______________"
    END IF
    PRINT #RptHandle, STRING$(79, "-")
    lp = 26

    IF LEN(QPTrim$(ARCustRec(1).BILLCAT2)) = 0 THEN GOTO EndAtmore
    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT2)
    GOSUB GetCode

    PRINT #RptHandle, TAB(2); ARCustRec(1).BILLCAT2;
    PRINT #RptHandle, TAB(7); ARCustRec(1).Desc2;
    PRINT #RptHandle, TAB(50); "BASIS AMOUNT"; TAB(66); "LICENSE AMOUNT"
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(50); "____________"; TAB(66); "______________"
    PRINT #RptHandle, STRING$(79, "-")

    lp = lp + 5

    IF LEN(QPTrim$(ARCustRec(1).BILLCAT3)) = 0 THEN GOTO EndAtmore
    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT3)
    GOSUB GetCode

    PRINT #RptHandle, TAB(2); ARCustRec(1).BILLCAT3;
    PRINT #RptHandle, TAB(7); ARCustRec(1).Desc3;
    PRINT #RptHandle, TAB(50); "BASIS AMOUNT"; TAB(66); "LICENSE AMOUNT"
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(50); "____________"; TAB(66); "______________"
    PRINT #RptHandle, STRING$(79, "-")

    lp = lp + 5

    IF LEN(QPTrim$(ARCustRec(1).BILLCAT4)) = 0 THEN GOTO EndAtmore
    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT4)
    GOSUB GetCode

    PRINT #RptHandle, TAB(2); ARCustRec(1).BILLCAT4;
    PRINT #RptHandle, TAB(7); ARCustRec(1).Desc4;
    PRINT #RptHandle, TAB(50); "BASIS AMOUNT"; TAB(66); "LICENSE AMOUNT"
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(50); "____________"; TAB(66); "______________"
    PRINT #RptHandle, STRING$(79, "-")
    lp = lp + 5
    IF LEN(QPTrim$(ARCustRec(1).BILLCAT5)) = 0 THEN GOTO EndAtmore
    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT5)
    GOSUB GetCode
    PRINT #RptHandle, TAB(2); ARCustRec(1).BILLCAT5;
    PRINT #RptHandle, TAB(7); ARCustRec(1).Desc5;
    PRINT #RptHandle, TAB(50); "BASIS AMOUNT"; TAB(66); "LICENSE AMOUNT"
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(50); "____________"; TAB(66); "______________"
    PRINT #RptHandle, STRING$(79, "-")

    lp = lp + 5

EndAtmore:

    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle,

    PRINT #RptHandle, TAB(5); "Make Checks Payable To:"; TAB(45); "License Total: _________________"
    PRINT #RptHandle, TAB(5); "TOWN OF CLINTWOOD"; TAB(45); "Penalty:       _________________"
    PRINT #RptHandle, TAB(5); "P.O. BOX 456"; TAB(45); "Interest:      _________________"
    PRINT #RptHandle, TAB(5); "CLINTWOOD VA, 24228"'; TAB(45); "Issue Fee:                 5.00"
    PRINT #RptHandle, TAB(5); ""; TAB(45); "               -----------------"
    PRINT #RptHandle, TAB(5); ""; TAB(45); "Total Due:     _________________"
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(5); "TO AVOID A 10% LATE PENALTY CHARGE, LICENSE RENEWALS"
    PRINT #RptHandle, TAB(5); "MUST BE MADE BEFORE MARCH 1ST."
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(5); "RENEWALS THAT DO NOT CONTAIN SIGNATURE AND GROSS RECEIPTS"
    PRINT #RptHandle, TAB(5); "(WHERE REQUIRED) WILL NOT BE PROCESSED."
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(5); "I CERTIFY THAT THE ABOVE INFORMATION IS CORRECT"
    PRINT #RptHandle, TAB(5); "NAME ________________________________ TITLE ________________________"
    PRINT #RptHandle, TAB(5); "SUBSCRIBED AND SWORN TO BEFORE ME THIS ______ DAY OF ______, ______."
    PRINT #RptHandle, CHR$(12);
    TotalCust = TotalCust + 1

    'PRINT #RptHandle, TAB(5);
    'PRINT #RptHandle,
    'PRINT #RptHandle,
    '
    'PRINT #RptHandle, CHR$(12);
    'TotalCust = TotalCust + 1

