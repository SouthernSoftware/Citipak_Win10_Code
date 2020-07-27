   ExpDate$ = "06-30-0" + QPTrim$(STR$(VAL(RIGHT$(DATE$, 1)) + 1))
   IsuDate$ = "05-15-0" + QPTrim$(STR$(VAL(RIGHT$(DATE$, 1))))

   PRINT #RptHandle, STR$(DCCustIdxRec(1).IDXRECORD); TAB(35); DCCustRec(1).SoSec
   PRINT #RptHandle,
   PRINT #RptHandle, TAB(65); IsuDate$
   PRINT #RptHandle, TAB(65); ExpDate$
   PRINT #RptHandle,
   PRINT #RptHandle,

  'PRINT #RptHandle, QPTrim$(LEFT$(DCVRec(1).Notes, 4));
  PRINT #RptHandle, TAB(2); LEFT$(DCVRec(1).makemodel, 19);
  PRINT #RptHandle, TAB(22); LEFT$(DCVRec(1).Desc, 30)
  'PRINT #RptHandle,
  'PRINT #RptHandle, TAB(70); USING "#######"; ControlNumber!

  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(69); "15.00"
  PRINT #RptHandle, TAB(69); " 8.00"
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(69); " 1.00"
  PRINT #RptHandle, ""
  PRINT #RptHandle, TAB(44); RIGHT$(IsuDate$, 2) + "-" + RIGHT$(ExpDate$, 2)
  PRINT #RptHandle, TAB(8); DCCustRec(1).BILLNAME; TAB(64); LEFT$(DCVRec(1).StateTag, 12)
  PRINT #RptHandle, TAB(8); DCCustRec(1).ADDRESS1'; TAB(64); "1146"
  PRINT #RptHandle, TAB(8); DCCustRec(1).ADDRESS2
  PRINT #RptHandle, TAB(8); RTRIM$(DCCustRec(1).city); " "; RTRIM$(DCCustRec(1).STATE); "  "; RTRIM$(DCCustRec(1).ZIPCODE)
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, "!"

