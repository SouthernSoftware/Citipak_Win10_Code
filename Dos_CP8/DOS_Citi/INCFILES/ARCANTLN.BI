  GOSUB MakeCatLine

  FOR ZZ = 1 TO 4
    PRINT #RptHandle,
  NEXT

  IF CustomerNumber MOD 4 THEN
    ChkFlag = 0
  ELSE
    ChkFlag = -1
  END IF
  QQ& = 0
  FOR ZZ& = 1 TO 500000
    QQ& = QQ& + 1
  NEXT

  PRINT #RptHandle,
  PRINT #RptHandle, TAB(28); "-----------------------"
  PRINT #RptHandle, TAB(28); "***DELINQUENT NOTICE***"
  PRINT #RptHandle, TAB(28); "-----------------------"
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(31 + ChkFlag); "  TOWN OF CANTON"
  PRINT #RptHandle, TAB(31); " 58 PARK STREET"
  PRINT #RptHandle, TAB(31); "CANTON, NC 28716"
  PRINT #RptHandle, TAB(31); "  828-648-2363"
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(LNDTab); LNotDate$
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
  PRINT #RptHandle, TAB(7); "purchased as of today. We realize that you are very busy, but we"
  PRINT #RptHandle, TAB(7); "would like for you to take the time to purchase this license."
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(16); "We show your account delinquent for the following"
  PRINT #RptHandle, TAB(16); "license code(s):  "; CatLine$
  PRINT #RptHandle, TAB(16); "Current Balance:  "; USING "####.##"; ARCustRec(1).AcctBal
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(7); "Please remit your payment to this office no later than"
  PRINT #RptHandle, TAB(7 + ChkFlag); LPayDate$; "."; " If you have questions regarding your"
  PRINT #RptHandle, TAB(7); "license, please feel free to contact our office."
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
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(52); "Town Clerk"
  PRINT #RptHandle, CHR$(12);

