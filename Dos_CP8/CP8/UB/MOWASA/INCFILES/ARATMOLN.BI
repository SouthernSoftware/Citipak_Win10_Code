  GOSUB MakeCatLine
  FOR zz = 1 TO 7
    PRINT #RptHandle,
  NEXT
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(28); "-----------------------"
  PRINT #RptHandle, TAB(28); "***DELINQUENT NOTICE***"
  PRINT #RptHandle, TAB(28); "-----------------------"
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(LNDTab); LNotDate$
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
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
  PRINT #RptHandle, TAB(7); "According to our records, your"; Year + 1; "Business License has not been"
  PRINT #RptHandle, TAB(7); "purchased as of today. All licenses are now subject to a "; PenPct$; " penalty"
  PRINT #RptHandle, TAB(7); "and will NOT be issued unless the penalty amount is included with"
  PRINT #RptHandle, TAB(7); "your payment. We realize that you are very busy, but we would like"
  PRINT #RptHandle, TAB(7); "for you to take the time to purchase this license."
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(16); "We show your account delinquent for the following"
  PRINT #RptHandle, TAB(16); "license code(s):  "; CatLine$
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(7); "Please remit your payment (including "; PenPct$; " penalty) to this office"
  PRINT #RptHandle, TAB(7); "NO later than: "; LPayDate$; "."; " If you have questions"
  PRINT #RptHandle, TAB(7); "regarding your license, please feel free to contact our office."
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(7); "If payment has been made prior to receiving this notice, please"
  PRINT #RptHandle, TAB(7); "disregard this notice."
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(52); "Sincerely,"
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(52); "Mary Carter"
  PRINT #RptHandle, TAB(52); "Accounting Clerk"
  PRINT #RptHandle, CHR$(12);

