'  Four Oaks
  IF DidCnt > 0 THEN
    PRINT #UBRpt,
    PRINT #UBRpt, ""
    PRINT #UBRpt, ""
  END IF
  
  PRINT #UBRpt, "                               TOWN OF FOUR OAKS"
  PRINT #UBRpt, "                               304 N MAIN STREET"
  PRINT #UBRpt, "                                  P.O. Box 610"
  PRINT #UBRpt, "                            FOUR OAKS, NC 27524-0610"
  PRINT #UBRpt, ""
  PRINT #UBRpt, ""
  PRINT #UBRpt, TAB(5); LongNDate$
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    "; UBCustRec(1).CustName
  PRINT #UBRpt, "    "; UBCustRec(1).ADDR1
  PRINT #UBRpt, "    "; UBCustRec(1).ServAddr
  PRINT #UBRpt, "    "; QPTrim$(UBCustRec(1).City); "  "; UBCustRec(1).State; "   "; UBCustRec(1).ZipCode
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    Dear water customer:"
  PRINT #UBRpt, ""
  PRINT #UBRpt, "      We have not received payment for your water bill this month.  To"
  PRINT #UBRpt, "    prevent your water from being cut-off, we must receive payment before"
  PRINT #UBRpt, "    the cut-off date of "; LongPDate$; ".  If your services are"
  PRINT #UBRpt, "    cut-off, there will be a reconnect fee of $25.00 added to your water"
  PRINT #UBRpt, "    bill.  This amount will have to be paid before your services can be"
  PRINT #UBRpt, "    reconnected.  If you have any questions, call for a hearing before"
  PRINT #UBRpt, "    the cut-off date to correct your bill."
  PRINT #UBRpt,
  PRINT #UBRpt, "          ACCOUNT: "; Acct$; TAB(40); USING "AMOUNT DUE:$$#####,#.##"; CustBal#
  PRINT #UBRpt, "         LOCATION: "; UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb
  PRINT #UBRpt,
  PRINT #UBRpt, "    Thank you,"
  PRINT #UBRpt, "    Town of Four Oaks"
  DidCnt = DidCnt + 1
  IF DidCnt > 1 THEN
    DidCnt = 0
    PRINT #UBRpt, CHR$(12)
  END IF

