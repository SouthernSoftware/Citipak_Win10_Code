    PRINT #RptHandle,
    PRINT #RptHandle, 'TAB(34); "CITY OF ATMORE"
    PRINT #RptHandle, 'TAB(35); "Revenue Dept"
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

    IF lp >= 45 THEN

      PRINT #RptHandle,
      PRINT #RptHandle, TAB(5); "PO Drawer 1297"; TAB(55); "201 E. Louisville Ave"
      PRINT #RptHandle, TAB(5); "Atmore, AL 36504"; TAB(33); "(334) 368-0854"; TAB(55); "  Atmore, AL 36052"
      PRINT #RptHandle, CHR$(12);

      PRINT #RptHandle,
      PRINT #RptHandle, TAB(34); "CITY OF ATMORE"
      PRINT #RptHandle, TAB(35); "Revenue Dept"
      PRINT #RptHandle,
    END IF


    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle,

    PRINT #RptHandle, TAB(5); "Make Checks Payable To:"; TAB(45); "License Total: _________________"
    PRINT #RptHandle, TAB(5); "CITY OF ATMORE,ALABAMA"; TAB(45); "Penalty:       _________________"
    PRINT #RptHandle, TAB(5); ""; TAB(45); "Interest:      _________________"
    PRINT #RptHandle, TAB(5); ""; TAB(45); "Issue Fee:                 5.00"
    PRINT #RptHandle, TAB(5); ""; TAB(45); "               -----------------"
    PRINT #RptHandle, TAB(5); ""; TAB(45); "Total Due:     _________________"
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(5); "License renewals are due January 1st and delinquent after"
    PRINT #RptHandle, TAB(5); "February 28, at which time a 10% penalty will be charged."
    PRINT #RptHandle, TAB(5); "Renewals after March 31 will be charged a 15% penalty.  If "
    PRINT #RptHandle, TAB(5); "you have any questions regarding this notice, please call"
    PRINT #RptHandle, TAB(5); "the Revenue Department"
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(10); "RENEWALS THAT DO NOT CONTAIN SIGNATURE AND GROSS RECEIPTS"
    PRINT #RptHandle, TAB(10); "(WHERE REQUIRED) WILL NOT BE PROCESSED."
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "I CERTIFY THAT THE ABOVE INFORMATION IS CORRECT"
    PRINT #RptHandle, TAB(2); "NAME ________________________________ TITLE ________________________"
    PRINT #RptHandle, TAB(2); "SUBSCRIBED AND SWORN TO BEFORE ME THIS ______ DAY OF ______, ______."
    PRINT #RptHandle, TAB(2); "NOTARY PUBLIC ____________________________________________"
    PRINT #RptHandle,
    PRINT #RptHandle,
    'PRINT #RptHandle, TAB(5); "PO Drawer 1297"; TAB(55); "201 E. Louisville Ave"
    'PRINT #RptHandle, TAB(5); "Atmore, AL 36504"; TAB(33); "(334) 368-0854"; TAB(55); "  Atmore, AL 36052"

    PRINT #RptHandle, CHR$(12);
    TotalCust = TotalCust + 1
'

