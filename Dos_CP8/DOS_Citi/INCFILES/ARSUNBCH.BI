    LicTotal# = 0
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(60); "Date: "; DATE$
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, "                            ANNUAL BUSINESS LICENSE"
    PRINT #RptHandle, "                                 RENEWAL NOTICE"
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(7); ARCustRec(1).BILLNAME
    PRINT #RptHandle, TAB(7); ARCustRec(1).Address1
    PRINT #RptHandle, TAB(7); ARCustRec(1).Address2
    PRINT #RptHandle, TAB(7); RTRIM$(ARCustRec(1).City); " " + RTRIM$(ARCustRec(1).State) + " " + RTRIM$(ARCustRec(1).ZipCode)
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "Account: "; ARCustIdxRec(1).IDXRECORD
    PRINT #RptHandle,
    PRINT #RptHandle, STRING$(79, "-")
    PRINT #RptHandle, TAB(2); "Code"; TAB(10); "Type of License"; TAB(50); "BASIS AMOUNT"; TAB(66); "LICENSE AMOUNT"
    PRINT #RptHandle, STRING$(79, "-")

    Lp = 19

    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT1)
    'PRINT #RptHandle, TAB(50); "BASIS AMOUNT"; TAB(66); "LICENSE AMOUNT"
    GOSUB GetCode
    GOSUB PrintThisCode
    LicTotal# = Round#(LicTotal# + Amt#)
    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT2)
    IF LEN(CATCODE$) = 0 THEN
      GOTO EndAtmore
    END IF
    
    GOSUB GetCode
    GOSUB PrintThisCode
    LicTotal# = Round#(LicTotal# + Amt#)

    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT3)
    IF LEN(CATCODE$) = 0 THEN
      GOTO EndAtmore
    END IF
    
    GOSUB GetCode
    GOSUB PrintThisCode

    LicTotal# = Round#(LicTotal# + Amt#)

    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT4)
    IF LEN(CATCODE$) = 0 THEN
      GOTO EndAtmore
    END IF
    
    GOSUB GetCode
    GOSUB PrintThisCode

    CATCODE$ = QPTrim$(ARCustRec(1).BILLCAT5)
    IF LEN(CATCODE$) = 0 THEN
      GOTO EndAtmore
    END IF

    GOSUB GetCode
    GOSUB PrintThisCode
    LicTotal# = Round#(LicTotal# + Amt#)
EndAtmore:

    IF Lp < 30 THEN
      FOR ZZCnt = Lp TO 30
        PRINT #RptHandle,
      NEXT
    END IF

    PRINT #RptHandle,
    PRINT #RptHandle, TAB(45); "License Total: "; USING "####.##"; LicTotal#
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(45); "----------------------------"
    PRINT #RptHandle, TAB(45); "    Total Due: "; USING "####.##"; LicTotal#
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(5); "MAKE CHECKS PAYABLE TO 'TOWN OF SUNSET BEACH'"
    PRINT #RptHandle, TAB(5); "ESTABLISHMENTS NOT PURCHASING A LICENSE BY "; QPTrim$(Form$(4, 0)); " WILL BE"
    PRINT #RptHandle, TAB(5); "REPORTED TO THE ABC COMMISSION."
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(5); "RENEWED LICENSE VALID FROM "; QPTrim$(Form$(5, 0)); " TO "; QPTrim$(Form$(6, 0))
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(5); "700 Sunset Boulevard North"
    PRINT #RptHandle, TAB(5); "Sunset Beach, NC 28468"
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(5); "Telephone: 910-579-6297 or 910-579-3808"
    PRINT #RptHandle, TAB(5); "      Fax: 910-579-1840"

    PRINT #RptHandle, CHR$(12);
    TotalCust = TotalCust + 1

GOTO DoneHere:

PrintThisCode:

    PRINT #RptHandle, TAB(2); CATCODE$;
    PRINT #RptHandle, TAB(7); CodeDesc$;
    SELECT CASE CodeType$
    CASE "S"
      PRINT #RptHandle, "Minimum Amount Due: "; USING "$$#####.##"; BaseAmt1#
      PRINT #RptHandle, "For Receipts Up To: "; USING "$$######,#.##"; Revenue1#
      PRINT #RptHandle, "Plus "; USING "##.###%  "; Percent1#;
      PRINT #RptHandle, "  Of Receipts Over "; USING "######"; Maximum1#
      Lp = Lp + 5
    CASE "F"
      PRINT #RptHandle, TAB(50); "Flat Fee: "; TAB(66); USING "######,#.##"; Amt#
      Lp = Lp + 2
    END SELECT
    PRINT #RptHandle, STRING$(79, "-")

RETURN

DoneHere:

