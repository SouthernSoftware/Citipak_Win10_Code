  GOSUB MakeCatLine
  FOR zz = 1 TO 3
    PRINT #RptHandle,
  NEXT
  PRINT #RptHandle, TAB(28); "-----------------------"
  PRINT #RptHandle, TAB(28); "***DELINQUENT NOTICE***"
  PRINT #RptHandle, TAB(28); "-----------------------"
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(LNDTab); LNotDate$
  PRINT #RptHandle,

'  PRINT #RptHandle, TAB(31); "TOWN OF LELAND"
'  PRINT #RptHandle, TAB(31); "102 TOWN HALL DRIVE"
'  PRINT #RptHandle, TAB(31); "LELAND, NC 28451"
'  PRINT #RptHandle, TAB(31); "TEL 910-371-0148"

  PRINT #RptHandle, TAB(31); "TOWN OF MAXTON"
  PRINT #RptHandle, TAB(31); "201 MCCASKILL AVE."
  PRINT #RptHandle, TAB(31); "MAXTON, NC 28364"
  PRINT #RptHandle, TAB(31); "TEL 910-844-5231"

  'PRINT #RptHandle, TAB(31); "CITY OF HAMLET"
  'PRINT #RptHandle, TAB(31); "P.O. BOX 1229"
  'PRINT #RptHandle, TAB(31); "HAMLET, NC 28345"
  'PRINT #RptHandle, TAB(31); "TEL 910-582-7983"


  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(52); USING "Account ID:######"; CustomerNumber
  PRINT #RptHandle, TAB(7); QPTrim$(ARCustRec(1).BILLNAME)
  PRINT #RptHandle, TAB(7); QPTrim$(ARCustRec(1).Address1)
  PRINT #RptHandle, TAB(7); QPTrim$(ARCustRec(1).Address2)
  PRINT #RptHandle, TAB(7); QPTrim$(ARCustRec(1).City); ", "; ARCustRec(1).State; " "; ARCustRec(1).ZipCode
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(7); "According to our records, your"; Year + 1; "Business License has not been"
  PRINT #RptHandle, TAB(7); "purchased as of today. All licenses are now subject to a "; PenPct$; " penalty"
  PRINT #RptHandle, TAB(7); "and will NOT be issued unless the penalty amount is included with"
  PRINT #RptHandle, TAB(7); "your payment. We realize that you are very busy, but we would like"
  PRINT #RptHandle, TAB(7); "for you to take the time to purchase this license."
  PRINT #RptHandle,
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(16); "We show your account delinquent for the following"
  PRINT #RptHandle, "                license code(s):"
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(16); ARCustRec(1).BILLCAT1; " "; ARCustRec(1).Desc1
  PRINT #RptHandle, TAB(16); ARCustRec(1).BILLCAT2; " "; ARCustRec(1).Desc2
  PRINT #RptHandle, TAB(16); ARCustRec(1).BILLCAT3; " "; ARCustRec(1).Desc3
  PRINT #RptHandle, TAB(16); ARCustRec(1).BILLCAT4; " "; ARCustRec(1).Desc4
  PRINT #RptHandle, TAB(16); ARCustRec(1).BILLCAT5; " "; ARCustRec(1).Desc5
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
  PRINT #RptHandle, TAB(51); "Sincerely,"
  PRINT #RptHandle,
  PRINT #RptHandle, TAB(52); "Jackie Johnson"
  PRINT #RptHandle, TAB(52); "Business License Clerk"
  PRINT #RptHandle, CHR$(12);

