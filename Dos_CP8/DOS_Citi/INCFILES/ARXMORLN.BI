  'GOSUB MakeCatLine
  Year$ = RIGHT$(DATE$, 4)
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(32); "TOWN OF EXMORE"
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(LNDTab); LNotDate$
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(7); ARCustRec(1).BILLNAME
  PRINT #RptHandle, TAB(7); ARCustRec(1).Address1
  PRINT #RptHandle, TAB(7); ARCustRec(1).Address2
  PRINT #RptHandle, TAB(7); QPTrim$(ARCustRec(1).City); ", "; ARCustRec(1).State; " "; ARCustRec(1).ZipCode
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(5); "Dear Business Owner:"
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(5); "  According to our records, your APPLICATION FOR TOWN LICENSE(S) has not"
  PRINT #RptHandle, TAB(5); "yet been submitted for processing your "; Year$; " BUSINESS, PROFESSIONAL and"
  PRINT #RptHandle, TAB(5); "OCCUPATIONAL LICENSE(BPOL) TAX.  This application is required prior to"
  PRINT #RptHandle, TAB(5); "issuance of any Exmore Business License.  The application form specifies"
  PRINT #RptHandle, TAB(5); "a MARCH 1 deadline for filing, and states that a penalty of "; PenPct$; " may be"
  PRINT #RptHandle, TAB(5); "assessed on delinquent applications. The application also states a"
  PRINT #RptHandle, TAB(5); "deadline of APRIL 1 for payment of applicable BPOL Tax, as stated in Code"
  PRINT #RptHandle, TAB(5); "of Virginia "; CHR$(21); "58.1-3703.1-Uniform Ordinance Provisions.  To avoid further"
  PRINT #RptHandle, TAB(5); "action, please complete and return the APPLICATION FOR TOWN LICENSE(S) "
  PRINT #RptHandle, TAB(5); "immediately.  Failure to comply may result in legal action including,"
  PRINT #RptHandle, TAB(5); "but not limited to, and audit of business records, as permited in Code"
  PRINT #RptHandle, TAB(5); "of Virginia "; CHR$(21); "58.1-3110 and "; CHR$(21); "58.1-3939.1"
  PRINT #RptHandle, TAB(5); ""
  PRINT #RptHandle, TAB(5); ""
  PRINT #RptHandle, TAB(5); ""
  PRINT #RptHandle, TAB(5); "  If there are any questions or if assistance is needed in completing the"
  PRINT #RptHandle, TAB(5); "form, please call Katrina Hickman at 442-3114, Monday - Friday 8:00 A.M."
  PRINT #RptHandle, TAB(5); "to 4:00 P.M."
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(5); "Cordially,"
  PRINT #RptHandle, TAB(5); "Mayor and Council"
  PRINT #RptHandle, TAB(5); "Town Of Exmore"
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(5); "Donald P. Kellam, Sr."
  PRINT #RptHandle, TAB(5); "BPOL Commissioner"
  PRINT #RptHandle, CHR$(12);

