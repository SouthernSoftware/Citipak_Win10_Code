  GOSUB MakeCatLine
  LNYear$ = STR$(Year + 1) + "/" + QPTrim$(STR$(Year + 2))
  FOR zz = 1 TO 16
    PRINT #RptHandle,
  NEXT
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(28); "-----------------------"
  PRINT #RptHandle, TAB(28); "***DELINQUENT NOTICE***"
  PRINT #RptHandle, TAB(28); "-----------------------"
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(LNDTab); LNotDate$
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(52); USING "Taxpayer ID:######"; CustomerNumber
  PRINT #RptHandle, TAB(7); ARCustRec(1).BILLNAME
  PRINT #RptHandle, TAB(7); ARCustRec(1).Address1
  PRINT #RptHandle, TAB(7); ARCustRec(1).Address2
  PRINT #RptHandle, TAB(7); QPTrim$(ARCustRec(1).City); ", "; ARCustRec(1).State; " "; ARCustRec(1).ZipCode
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(7); "Dear Taxpayer:"
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(7); "According to our records, your"; LNYear$; " Business License has not"
  PRINT #RptHandle, TAB(7); "been purchased as of today. All licenses are now subject to a "; PenPct$
  PRINT #RptHandle, TAB(7); "penalty and will NOT be issued unless the penalty amount is included"
  PRINT #RptHandle, TAB(7); "with your payment. We realize that you are very busy, but we would"
  PRINT #RptHandle, TAB(7); "like for you to take the time to purchase this license."
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(16); "We show your account delinquent for the following"
  PRINT #RptHandle, TAB(16); "license code(s):  "; CatLine$
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(7); "Please remit your payment (including "; PenPct$; " penalty) to this office"
  PRINT #RptHandle, TAB(7); "NO later than: "; LPayDate$; ", "; "otherwise your name will be"
  PRINT #RptHandle, TAB(7); "included on a delinquent list that will be supplied to the local media"
  PRINT #RptHandle, TAB(7); "for publishing. If you have questions regarding your license, please"
  PRINT #RptHandle, TAB(7); "feel free to contact our office."
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(7); "If payment has been made prior to receiving this notice, please"
  PRINT #RptHandle, TAB(7); "call our office."
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(52); "Sincerely,"
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(52); "Tina M. Rowe"
  PRINT #RptHandle, TAB(52); "Treasurer"
  PRINT #RptHandle, CHR$(12);

