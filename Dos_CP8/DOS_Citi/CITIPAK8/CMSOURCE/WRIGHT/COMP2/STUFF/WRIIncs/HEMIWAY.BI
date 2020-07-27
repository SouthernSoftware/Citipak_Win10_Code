PRINT #UBRpt, "~"; TAB(45); "Town of Hemingway"; TAB(68); "~"
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, TAB(19); Acct$    'line 4
PRINT #UBRpt, TAB(10); "**PAST DUE NOTICE**"
PRINT #UBRpt,
PRINT #UBRpt, TAB(5); "NOTICE DATE "; NDate$
PRINT #UBRpt, TAB(5); UBCustRec(1).ServAddr
PRINT #UBRpt, TAB(5); "YOUR ACCOUNT IS PAST DUE. IF BILL"; TAB(42); QPTrim$(UBCustRec(1).CustName)
PRINT #UBRpt, TAB(5); "IS NOT PAID BY 5:00 ON"; TAB(42); QPTrim$(UBCustRec(1).Addr1)
PRINT #UBRpt, TAB(5); PDate$; " YOUR SERVICE WILL BE"; TAB(42); QPTrim$(UBCustRec(1).Addr2)
PRINT #UBRpt, TAB(5); "DISCONTINUED, AND A FEE OF $10.00"; TAB(42); QPTrim$(UBCustRec(1).City); " "; UBCustRec(1).State
PRINT #UBRpt, TAB(5); "WILL BE ADDED TO YOUR ACCOUNT."; TAB(51); UBCustRec(1).ZipCode
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, ""
PRINT #UBRpt, TAB(39); FUsing$(STR$(TotalBal#), "$$,######.##")           'line 22
PRINT #UBRpt, ""
PRINT #UBRpt, "~"'line 24

