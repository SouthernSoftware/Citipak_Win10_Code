  ThisYear = (VAL(RIGHT$(DATE$, 4)) - 1)
  TYear$ = STR$(ThisYear + 1)
  LYear$ = STR$(ThisYear)

  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(35); "TOWN OF STUART"
  PRINT #RptHandle, TAB(36); "P.O. BOX 422"
  PRINT #RptHandle, TAB(34); "STUART, VA 24171"
  PRINT #RptHandle, TAB(35); "(276) 694-3811"
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(5); "Acct:"; DCCustIdxRec(1).IDXRECORD
  PRINT #RptHandle, TAB(5); DCCustRec(1).BILLNAME
  PRINT #RptHandle, TAB(5); DCCustRec(1).ADDRESS1
  PRINT #RptHandle, TAB(5); DCCustRec(1).ADDRESS2
  PRINT #RptHandle, TAB(5); RTRIM$(DCCustRec(1).city); " "; RTRIM$(DCCustRec(1).STATE); "  "; RTRIM$(DCCustRec(1).ZIPCODE)
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(5); "Dear resident:"
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(5); "The town will be selling the"; TYear$; " decal beginning March 1,"; TYear$; " and"
  PRINT #RptHandle, TAB(5); "the expiration date of the"; LYear$; " decal is April 30,"; TYear$; "."
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(5); "According to our records you own the following vehicle garaged in town:"
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
'  PRINT #RptHandle, TAB(2); LEFT$(DCVRec(1).Notes, 4);
  PRINT #RptHandle, TAB(5); "Vehicle make & model: "; QPTrim$(DCVRec(1).makemodel)
  PRINT #RptHandle, TAB(5); "      State license#: "; QPTrim$(DCVRec(1).StateTag)
  PRINT #RptHandle, TAB(5); "                VIN#: "; QPTrim$(DCVRec(1).Desc)
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(5); "Please verify this information, or correct and send renewal for"
  PRINT #RptHandle, TAB(5); "each to the Town Office or stop by between 9:00 to 1:00 p.m."
  PRINT #RptHandle, TAB(5); "or 2:00 to 5:00 p.m to renew in person."
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(5); "All personal property taxes must be paid in order to renew decals."
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(5); "Thank You,"
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(5); "Town Clerk"
  PRINT #RptHandle, ff$

