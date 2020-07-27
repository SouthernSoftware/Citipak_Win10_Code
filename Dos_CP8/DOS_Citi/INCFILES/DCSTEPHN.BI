  ThisYear = VAL(RIGHT$(DATE$, 4))
  'TYear$ = STR$(ThisYear + 1)
  TYear$ = STR$(ThisYear)
  LYear$ = STR$(ThisYear - 1)

  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  'PRINT #RptHandle, TAB(31); "TOWN OF STEPHENS CITY"
  'PRINT #RptHandle, TAB(31); " Post Office Box 250"
  'PRINT #RptHandle, TAB(27); "STEPHENS CITY, VIRGINIA 22655"
  'PRINT #RptHandle, TAB(34); "(540) 869-3087"

  PRINT #RptHandle, TAB(31); "TOWN OF INDEPENDENCE"
  PRINT #RptHandle, TAB(31); "    P.O. Box 99"
  PRINT #RptHandle, TAB(27); "INDEPENDENCE, VIRGINIA 24348"
  PRINT #RptHandle, TAB(34); "(276) 773-3703"

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
  PRINT #RptHandle, TAB(5); "The town will be selling the"; TYear$; " decal after March 15,"; TYear$; " and"
  PRINT #RptHandle, TAB(5); "the expiration date of the"; LYear$; " decal is April 15,"; TYear$; "."
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
  PRINT #RptHandle, TAB(5); "Please verify this information, or correct and send $15.00 for each"
  PRINT #RptHandle, TAB(5); "automobile renewal, or $7.50 for motorcycle renewals, to the"
  PRINT #RptHandle, TAB(5); "Town Office or stop by between 8:30 a.m. and 5:00 p.m."
  PRINT #RptHandle, TAB(5); "to renew in person."
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(5); "All personal property taxes must be paid in order to renew decals."
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(5); "Thank You,"
  PRINT #RptHandle, ff$
  'PRINT #RptHandle,
  'PRINT #RptHandle,
  'PRINT #RptHandle,
  'PRINT #RptHandle,
  'PRINT #RptHandle,
  'PRINT #RptHandle, TAB(5); "Joyce E. Zalar"
  'PRINT #RptHandle, TAB(5); "Town Clerk"
  

