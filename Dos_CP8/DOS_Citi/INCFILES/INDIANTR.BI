'This is indian trails
  FOR ZCnt = 1 TO 25
    PRINT #UBRpt, ""
  NEXT
  PRINT #UBRpt, TAB(10); UBCustRec(1).CustName
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR1
  PRINT #UBRpt, TAB(10); UBCustRec(1).ADDR2
  PRINT #UBRpt, TAB(10); UBCustRec(1).City; "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
  PRINT #UBRpt, TAB(65); NDate$
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(10); "We have not received your remittance in the amount of "; FUsing$(STR$(CustBal#), ",######.##");
  PRINT #UBRpt, TAB(10); "for the Garbage and Landfill Supplement which is past due."
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(10); "Please mail your check by "; pdate$; " so that your account will be"
  PRINT #UBRpt, TAB(10); "current."
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(10); "Thanking you in advance for your cooperation in this matter."
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(10); "TOWN OF INDIAN TRAIL"
  PRINT #UBRpt, TAB(10); "P.O. BOX 2430"
  PRINT #UBRpt, TAB(10); "INDIAN TRAIL, NC 28079"
  PRINT #UBRpt, TAB(10); "PHONE: (704) 821-8114"
  PRINT #UBRpt, CHR$(12);

