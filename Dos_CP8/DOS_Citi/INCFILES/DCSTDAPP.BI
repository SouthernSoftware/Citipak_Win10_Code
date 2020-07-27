  PRINT #RptHandle, TAB(2); LEFT$(DCVRec(1).Notes, 4);
  PRINT #RptHandle, TAB(13); LEFT$(DCVRec(1).makemodel, 11);
  PRINT #RptHandle, TAB(28); LEFT$(DCVRec(1).Desc, 11);
  PRINT #RptHandle, TAB(43); LEFT$(DCVRec(1).StateTag, 12);
  PRINT #RptHandle, TAB(70); USING "#######"; ControlNumber!

  FOR LCnt = 1 TO 11
    PRINT #RptHandle,
  NEXT LCnt
  PRINT #RptHandle, TAB(8); DCCustRec(1).BILLNAME
  PRINT #RptHandle, TAB(8); DCCustRec(1).ADDRESS1
  PRINT #RptHandle, TAB(8); DCCustRec(1).ADDRESS2
  PRINT #RptHandle, TAB(8); RTRIM$(DCCustRec(1).city); " "; RTRIM$(DCCustRec(1).STATE); "  "; RTRIM$(DCCustRec(1).ZIPCODE)
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(5); DCCustRec(1).SOSEC;
  PRINT #RptHandle, TAB(27); DCCustRec(1).SocSec1
  FOR LCnt = 1 TO 4
    PRINT #RptHandle, ""
  NEXT LCnt

