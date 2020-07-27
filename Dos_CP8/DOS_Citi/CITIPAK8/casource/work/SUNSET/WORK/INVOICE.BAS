
    'IF NOT (UBCust(1).Deleted) THEN

'      IF lineCnt >= MaxLines THEN
'        PRINT #UBRpt, FF$
'        GOSUB PrintBalanceCustomerRptHeader
'      END IF

      'Determine if Customer HAS A BALANCE & PRINT IF THEY DO!
      CustBalance# = 0
      FOR RVCnt = 1 TO 6
        CustBalance# = Round#(CustBalance# + UBBillRec(1).RevAmt(RVCnt))
      NEXT

      IF CustBalance# <= 0 THEN GOTO CUSTSKIP

      RunPrinTotal# = 0
      RunIntTotal# = 0

      PRINT #UBRpt, TAB(29); "TOWN OF SUNSET BEACH"
      PRINT #UBRpt, TAB(26); "FINANCE DEPT - COLLECTION DIV"
      PRINT #UBRpt, TAB(30); "700 SUNSET BLVD. N."
      PRINT #UBRpt, TAB(28); "SUNSET BEACH, NC 28468"
      PRINT #UBRpt,
      PRINT #UBRpt,
      PRINT #UBRpt, TAB(60); "DATE: "; DATE$
      PRINT #UBRpt,

      PRINT #UBRpt, TAB(3); LEFT$(RTRIM$(UBCustRec(1).CustName), 30);
      PRINT #UBRpt, TAB(40); "    ACCT # "; CustAcctNo&
      PRINT #UBRpt, TAB(3); LEFT$(RTRIM$(UBCustRec(1).Addr1), 30)
      PRINT #UBRpt, TAB(3); QPTrim$(UBCustRec(1).Addr2)
      PRINT #UBRpt, TAB(3); QPTrim$(UBCustRec(1).City); " "; QPTrim$(UBCustRec(1).State); " "; QPTrim$(UBCustRec(1).ZipCode);
      PRINT #UBRpt, TAB(40); "LEGAL ADDR "; QPTrim$(UBCustRec(1).ServAddr)

      PRINT #UBRpt, ""
      PRINT #UBRpt, ""
      PRINT #UBRpt, ""
      PRINT #UBRpt, TAB(3); "THIS IS AN INVOICE FOR ASSESSMENTS DUE TO THE TOWN OF SUNSET BEACH."
      PRINT #UBRpt, TAB(3); "THE STATUS OF THIS ACCOUNT IS SHOWN BELOW."
      PRINT #UBRpt, ""
      PRINT #UBRpt, ""
      'PRINT #UBRpt, TAB(11); "- ORIG ASSESSMENT -"; TAB(31); "INTEREST"; TAB(41); "-- PRIOR PAYMENTS --"; TAB(62); "  LAST  "; TAB(70); "PRINCIPLE"
      PRINT #UBRpt, TAB(3); "ASSESSMENT         AMOUNT"; TAB(65); " BALANCE"

      FOR RVCnt = 1 TO 6
      IF UBBillRec(1).RevAmt(RVCnt) <> 0 THEN
        PRINT #UBRpt, QPTrim$(UBSetUpRec(1).Revenues(RVCnt).RevName);
        PRINT #UBRpt, TAB(23); USING "#####.##"; UBBillRec(1).RevAmt(RVCnt)
        RunPrinTotal# = Round#(RunPrinTotal# + UBBillRec(1).RevAmt(RVCnt))
        lineCnt = lineCnt + 1
      END IF
      NEXT
      PRINT #UBRpt, ""
      PRINT #UBRpt, ""
      PRINT #UBRpt, TAB(5); "Assessment payments may be paid in full or in 4 equal annual"
      PRINT #UBRpt, TAB(5); "installments, based on the confirmation of the assessment roll."
      PRINT #UBRpt, TAB(5); "Interest is accrued monthly (using the simple interest method)"
      PRINT #UBRpt, TAB(5); "on any outstanding principle balance."
      PRINT #UBRpt, ""
      PRINT #UBRpt, ""
      PRINT #UBRpt, "Principle Due ... "; USING "$$#####,#.##"; RunPrinTotal#
      PRINT #UBRpt, "Interest Due .... "; USING "$$#####,#.##"; RunIntTotal#
      PRINT #UBRpt, "                 -------------"
      PRINT #UBRpt, "Total Amount Due  "; USING "$$#####,#.##"; Round#(RunPrinTotal# + RunIntTotal#)
      PRINT #UBRpt, FF$
    'END IF
    CustCnt = CustCnt + 1
CUSTSKIP:
'  NEXT Cnt

  PRINT #UBRpt, CHR$(18);   ' oki 320 10 cpi

 ' CLOSE         'Close all open files now


 ' PrintRptFile Header$, ReportFile$, LptPort%, RetCode%, EntryPoint

 ' KILL ReportFile$



'PrintBalanceCustomerRptEnding:
'  PRINT #UBRpt, Dash80$
'  PRINT #UBRpt, "Total Customers Printed: "; USING "#####"; CustCnt
'  PRINT #UBRpt,
'  PRINT #UBRpt, "Assessment Totals Breakdown"
'  PRINT #UBRpt,
'  PRINT #UBRpt, "Assessment"; TAB(19); "Prin Balance"; TAB(35); "Int Balance"; TAB(52); "Total Due"
'  PRINT #UBRpt, "Street"; TAB(20); USING "######,#.##"; RunPrinStreet#;
'  PRINT #UBRpt, TAB(35); USING "######,#.##"; RunIntStreet#;
'  PRINT #UBRpt, TAB(50); USING "######,#.##"; RunPrinStreet# + RunIntStreet#
'  PRINT #UBRpt, "SideWalks"; TAB(20); USING "######,#.##"; RunPrinSide#;
'  PRINT #UBRpt, TAB(35); USING "######,#.##"; RunIntSide#;
'  PRINT #UBRpt, TAB(50); USING "######,#.##"; RunPrinSide# + RunIntSide#
'  PRINT #UBRpt, "DriveWays"; TAB(20); USING "######,#.##"; RunPrinDrive#;
'  PRINT #UBRpt, TAB(35); USING "######,#.##"; RunIntDrive#;
'  PRINT #UBRpt, TAB(50); USING "######,#.##"; RunPrinDrive# + RunIntDrive#
'  PRINT #UBRpt, "Dredging"; TAB(20); USING "######,#.##"; RunPrinDemo#;
'  PRINT #UBRpt, TAB(35); USING "######,#.##"; RunIntDemo#;
'  PRINT #UBRpt, TAB(50); USING "######,#.##"; RunPrinDemo# + RunIntDemo#
'  PRINT #UBRpt, "Undg Utilities"; TAB(20); USING "######,#.##"; RunPrinWater#;
'  PRINT #UBRpt, TAB(35); USING "######,#.##"; RunIntWater#;
'  PRINT #UBRpt, TAB(50); USING "######,#.##"; RunPrinWater# + RunIntWater#
'  PRINT #UBRpt, "Sewer"; TAB(20); USING "######,#.##"; RunPrinSewer#;
'  PRINT #UBRpt, TAB(35); USING "######,#.##"; RunIntSewer#;
'  PRINT #UBRpt, TAB(50); USING "######,#.##"; RunPrinSewer# + RunIntSewer#
'  PrinTotal# = Round#(RunPrinStreet# + RunPrinSide# + RunPrinDrive# + RunPrinDemo# + RunPrinWater# + RunPrinSewer#)
'  IntTotal# = Round#(RunIntStreet# + RunIntSide# + RunIntDrive# + RunIntDemo# + RunIntWater# + RunIntSewer#)
'  PRINT #UBRpt, STRING$(62, "=")
'  PRINT #UBRpt, "Totals"; TAB(20); USING "######,#.##"; PrinTotal#;
'  PRINT #UBRpt, TAB(35); USING "######,#.##"; IntTotal#;
'  PRINT #UBRpt, TAB(50); USING "######,#.##"; PrinTotal# + IntTotal#
'
'
'
'  PRINT #UBRpt, FF$
'  RETURN

