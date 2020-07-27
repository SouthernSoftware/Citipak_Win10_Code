    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(26); "APPLICATION FOR TOWN LICENSE"
    PRINT #RptHandle, TAB(33); "TOWN OF STUART"
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "For period beginning May 1, "; Year$; ", and ending April 30,"; STR$(VAL(Year$) + 1)
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); " APPLICANT: "; ARCustRec(1).BILLNAME
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "TRADING AS: "; ARCustRec(1).CustName
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "   ADDRESS: "; ARCustRec(1).Address1
    PRINT #RptHandle, TAB(13); QPTrim$(ARCustRec(1).Address2); " "; QPTrim$(ARCustRec(1).City); " " + QPTrim$(ARCustRec(1).State) + " " + QPTrim$(ARCustRec(1).ZipCode)
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(33); "GROSS RECEIPTS"
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "  Gross receipts means the gross sales of merchandise and the gross receipts"
    PRINT #RptHandle, TAB(2); "of the business, occupation or professional from all earnings, fees,"
    PRINT #RptHandle, TAB(2); "commissions, and from all income whatsoever, arising from or growing out of"
    PRINT #RptHandle, TAB(2); "the business, occupation or profession. Gross receipts shall be for the most"
    PRINT #RptHandle, TAB(2); "recent preceding fiscal or calendar year used by the applicant for state and"
    PRINT #RptHandle, TAB(2); "federal income tax perposes. Town License are due on or before May 31."
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "A. For applications having at least one year's prior business experience to"
    PRINT #RptHandle, TAB(2); "   report"
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "    1.  Fiscal year/Calendar year from ________________ to _________________"
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "    2.  Gross receipts______________________________________________________"
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "    3.  Gross receipts tax____________;(due on or before May 31st,) Shall be"
    PRINT #RptHandle, TAB(2); "        calculated by dividing the gross receipts figure by $100.00 and"
    PRINT #RptHandle, TAB(2); "        multiplying by the appropriate tax rate below, i.e. $50,000.00 "
    PRINT #RptHandle, TAB(2); "        divided by $100.00 = $500.00 x .10 (if retail sales) = $50.00"
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "B. For applicants having no prior year's business experience:"
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "    1.  Date business began_________________________________________________"
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "    2.  Estimated gross receipts for first fiscal year/calender year________"
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(2); "    3.  Gross receipts tax__________________________________________________"
    PRINT #RptHandle, TAB(2); CHR$(12);

