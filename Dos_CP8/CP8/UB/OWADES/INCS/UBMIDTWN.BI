'Middletown Va

  Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
  TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

  FOR WMtrCnt = 1 TO 7
    SELECT CASE UBBillRec(1).MtrTypes(WMtrCnt)
    CASE MtrWaterOnly, MtrSewerOnly, MtrCombined, MtrTouchRead
      IF UBBillRec(1).PrevRead(WMtrCnt) < 0 THEN
        UBBillRec(1).PrevRead(WMtrCnt) = 0
      END IF
      IF UBBillRec(1).CurRead(WMtrCnt) < 0 THEN
        UBBillRec(1).CurRead(WMtrCnt) = 0
      END IF
      Current# = UBBillRec(1).CurRead(WMtrCnt)
      PrevRead# = UBBillRec(1).PrevRead(WMtrCnt)
      UsageAmt# = UBBillRec(1).CurRead(WMtrCnt) - UBBillRec(1).PrevRead(WMtrCnt)
      IF UsageAmt# < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBBillRec(1).PrevRead(WMtrCnt))) - 1)
        UsageAmt# = (MaxMeterAmt& - UBBillRec(1).PrevRead(WMtrCnt)) + UBBillRec(1).CurRead(WMtrCnt)
      END IF
      EXIT FOR
    END SELECT
  NEXT

   AcctNum = UBBillRec(1).CustAcctNo

     PRINT #UBRpt, "Date "; Num2Date$(UBBillRec(1).BillDate); TAB(51); "Town of Middletown"
     PRINT #UBRpt, "Present Reading"; TAB(30); USING "#######"; Current#
     PRINT #UBRpt, "Previous Reading"; TAB(30); USING "#######"; PrevRead#;
     PRINT #UBRpt, TAB(55); "P.O. Box 96"
     PRINT #UBRpt, "Consumed (Thousands of Gal.) "; TAB(30); USING "#######"; UsageAmt#
     PRINT #UBRpt, "Water Charge"; TAB(32); USING "#####,#.##"; UBBillRec(1).RevAmt(1);
     PRINT #UBRpt, TAB(50); "Middletown, VA 22645"
     PRINT #UBRpt, "Sewer Charge"; TAB(32); USING "#####,#.##"; UBBillRec(1).RevAmt(2)
     PRINT #UBRpt, "State Surcharge"; TAB(32); USING "#####,#.##"; UBBillRec(1).RevAmt(3);
     PRINT #UBRpt, TAB(54); "(540-869-2226)"
     PRINT #UBRpt, "Misc"; TAB(32); USING "#####,#.##"; UBBillRec(1).RevAmt(5)
     PRINT #UBRpt, "Pastdue Balance Due: "; TAB(32); USING "#####,#.##"; Previous#;
     PRINT #UBRpt, TAB(50); "Billing Cycle: "; QPTrim$(Message$)
     PRINT #UBRpt, "Total Due"; TAB(30); USING "$$#####,#.##"; TotalAmt#;
     PRINT #UBRpt,
     PRINT #UBRpt, TAB(50); "Due Date: "; Num2Date$(UBBillRec(1).pastduedate)
     PRINT #UBRpt,
     PRINT #UBRpt,
     PRINT #UBRpt, USING "####"; AcctNum;
     PRINT #UBRpt, TAB(10); UBCustRec(1).CustName
     PRINT #UBRpt, TAB(10); UBCustRec(1).Addr1
     PRINT #UBRpt, TAB(10); UBCustRec(1).Addr2
     PRINT #UBRpt, TAB(10); LEFT$(UBCustRec(1).CITY, 18); " "; UBCustRec(1).STATE; " "; LEFT$(UBCustRec(1).ZIPCODE, 5)
     PRINT #UBRpt,
     REM condense here lines 19 & 20
     PRINT #UBRpt, CHR$(15);
     PRINT #UBRpt, "All bills due upon presentation. A penalty of 10% will be added 30 days after billing date. Services will be discontinued if not"
     PRINT #UBRpt, "paid within 35 days from billing date.  Receipt will not be returned unless requested.  A $10.00 charge on all returned checks."
     PRINT #UBRpt, CHR$(18);
     REM uncondense here
     PRINT #UBRpt, "~"

'  PRINT #UBRpt, "  "
'  PRINT #UBRpt,
'  PRINT #UBRpt, "~"
'

