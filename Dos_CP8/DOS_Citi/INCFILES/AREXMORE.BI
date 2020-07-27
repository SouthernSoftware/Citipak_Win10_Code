    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(33); "TOWN OF EXMORE"
    PRINT #RptHandle, TAB(16); "BUSINESS, PROFESSIONAL AND OCCUPATIONAL LICENSE"
    PRINT #RptHandle, TAB(31); "For Year: "; Form$(2, 0); TAB(70); "PAGE 1"
    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(2); "Dear Business Owner:"
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "     For the purpose of computing Business, Professional and Occupational"
    PRINT #RptHandle, TAB(2); "License (BPOL) Tax promulgated by Virginia Code Section 58.1-3700 et seq."
    PRINT #RptHandle, TAB(2); "and Exmore Town Ordinance #129 Uniform B.P.O.L. Policy adopted 12/1/96, "
    PRINT #RptHandle, TAB(2); "please complete and return this form with the required information no later"
    PRINT #RptHandle, TAB(2); "than March 1."
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "Respectfully,"
    PRINT #RptHandle, TAB(2); "Exmore Town Council"
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); STRING$(76, "-")
    PRINT #RptHandle, TAB(2);
    PRINT #RptHandle, TAB(2);
    PRINT #RptHandle, TAB(33); "TOWN OF EXMORE"
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(24); "Application for Town Licenses"
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "For period beginning January 1, "; Year$; " (or start of business in "; Year$; " ) and"
    PRINT #RptHandle, TAB(2); "ending December 31, "; Year$
    PRINT #RptHandle, TAB(2);
    PRINT #RptHandle, TAB(2); "NAME OF APPLICANT: "; ARCustRec(1).BILLNAME
    PRINT #RptHandle, TAB(2); "      TRADING AS : "; ARCustRec(1).CustName
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(2); "BUSINESS ADDRESS:"; TAB(40); "HOME ADDRESS"
    PRINT #RptHandle, TAB(2); "MAIL: "; ARCustRec(1).Address1; TAB(40); "Mail: ________________________________"
    PRINT #RptHandle, TAB(8); ARCustRec(1).Address2; TAB(40); "      ________________________________"
    PRINT #RptHandle, TAB(8); RTRIM$(ARCustRec(1).City); " " + RTRIM$(ARCustRec(1).State) + " " + RTRIM$(ARCustRec(1).ZipCode)
    PRINT #RptHandle, TAB(2); "911:  ______________________________"
    PRINT #RptHandle, TAB(2);
    PRINT #RptHandle, TAB(8); "______________________________"
    PRINT #RptHandle, TAB(2);
    PRINT #RptHandle, TAB(2); "Phone: _________________________"; TAB(40); "PHONE: ______________________________"
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "A SEPERATE LICENSE WILL BE ISSUED FOR EACH TYPE OF BUSINESS PERFORMED, AS"
    PRINT #RptHandle, TAB(2); "REQUIRED PER THE UNIFORM BPOL ORDINANCE.  THIS WILL NOT RESULT IN ANY"
    PRINT #RptHandle, TAB(2); "ADDITONAL COST TO BUSINESSES.  PLEASE REPORT GROSS RECEIPTS FOR EACH"
    PRINT #RptHandle, TAB(2); "CLASSIFICATION THAT APPLIES TO YOUR BUSINESS."
    PRINT #RptHandle, CHR$(12);
    PRINT #RptHandle, TAB(33); "TOWN OF EXMORE"
    PRINT #RptHandle, TAB(16); "BUSINESS, PROFESSIONAL AND OCCUPATIONAL LICENSE"
    PRINT #RptHandle, TAB(31); "For Year: "; Form$(2, 0); TAB(70); "PAGE 1"
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(2); "WHOLESALE MERCHANT:"
    PRINT #RptHandle, TAB(2); "Gross Purchases through 12-31-____ as shown by applicants records"
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(2); TAB(60); "$_______________"
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(2); "RETAIL MERCHANT:"
    PRINT #RptHandle, TAB(2); "Gross Receipts through 12-31-____ as shown by applicants records"
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(2); TAB(60); "$_______________"
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(2); "FINANCIAL, REAL ESTATE AND PROFESSIONAL:"
    PRINT #RptHandle, TAB(2); "Gross Receipts through 12-31-____ as shown by applicants records"
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(2); TAB(60); "$_______________"
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(2); "CONTRACTING:"
    PRINT #RptHandle, TAB(2); "Gross Receipts through 12-31-____ as shown by applicants records"
    PRINT #RptHandle, TAB(2); "Subject to Virginia Code Sec 58.1-3715)"; TAB(60); "$_______________"
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(2); "REPAIR, PERSONAL or BUSINESS SERVICES:"
    PRINT #RptHandle, TAB(2); "Gross Receipts through 12-31-____ as shown by applicants records"
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(2); TAB(60); "$_______________"
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(2); "If uncertain of your business classification(s), please call the Town"
    PRINT #RptHandle, TAB(2); "Office at (757)442-3114 for assistance."
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(2); "I do affirm that the foregoing figures are true, complete and accurate"
    PRINT #RptHandle, TAB(2); "to the best of my knowledge."
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(35); "Signature ___________________________________"
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(34); "Print Name ___________________________________"
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(33); "*** IMPORTANT ***"
    PRINT #RptHandle, TAB(2); ""
    PRINT #RptHandle, TAB(2); "APPLICATION MUST BE RETURNED PRIOR TO MARCH 1 OF EACH YEAR TO AVOID PENALTY."
    PRINT #RptHandle, TAB(2); "LICENSE FEES DUE PRIOR TO APRIL 1 OF EACH YEAR TO AVOID PENALTY AND INTEREST."
    PRINT #RptHandle, TAB(2); "INTENTIONALLY PROVIDING INSUFFICIENT OR INACCURATE INFORMATION MAY RESULT"
    PRINT #RptHandle, TAB(2); "IN LEGAL RECOURSE BY THE TOWN OF EXMORE AS SET FORTH BY VIRGINIA CODE."
    PRINT #RptHandle, TAB(2); CHR$(12);

