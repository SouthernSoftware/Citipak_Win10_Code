'Pocahontas

IF DoneHeader = False THEN
  PLineCnt = 0
  DoneHeader = True
  PRINT #RptFile,
  PRINT #RptFile, TAB(66); DATE$

  PRINT #RptFile, TAB(t1); TownName$
  PRINT #RptFile, TAB(s1); ADD1$
  PRINT #RptFile, TAB(s2); ADD2$
  PRINT #RptFile, TAB(s3); ADD3$
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(4);
  PRINT #RptFile, TAB(10); CustName$
  PRINT #RptFile, TAB(10); LEFT$(TaxCust(1).Addr1, 25); TAB(60); "Acct: "; CustAcct&
  PRINT #RptFile, TAB(10); LEFT$(TaxCust(1).Addr2, 25); TAB(60); "Tele: "; TaxCust(1).HPHONE
  PRINT #RptFile, TAB(10); RTRIM$(TaxCust(1).City) + ", "; TaxCust(1).State + " " + LEFT$(TaxCust(1).Zip, 5) + "-" + MID$(TaxCust(1).Zip, 6, 4)
  PRINT #RptFile,
  PRINT #RptFile,
  IF RealFlag THEN
      PRINT #RptFile, "REAL ESTATE LATE NOTICE"
  ELSE
      PRINT #RptFile, "PERSONAL PROPERTY LATE NOTICE"
  END IF
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, STRING$(80, "-")
  PRINT #RptFile, "Description"; TAB(32); "Principal"; TAB(53); "P/I"; TAB(69); "Balance"
  PRINT #RptFile, STRING$(80, "-")
END IF

IF NOT DoFooter THEN
  'PRINT #RptFile, TAB(19); TaxTrans(1).TaxYear; TAB(40); USING "$$#######.##"; Balance#
  PLineCnt = PLineCnt + 1
  'PRINT #RptFile, QPTrim$(TaxTrans(1).Description); TAB(30); USING "######,#.##"; Round#(Balance# - IntBal#); TAB(45); TxIntC#; TAB(65); Balance#
  PRINT #RptFile, TAB(2); USING "####"; TaxTrans(1).TaxYear;
  PRINT #RptFile, " "; LEFT$(QPTrim$(TaxTrans(1).Description), 22); TAB(30); USING "######,#.##"; Round#(Balance# - InterestDue#); TAB(45); InterestDue#; TAB(65); Balance#
END IF
IF DoFooter THEN
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(10); "Total taxes & int. due"; TAB(40); USING "$$#######.##"; TBalance#
  IF PLineCnt < 8 THEN
    FOR CCnt = PLineCnt TO 8
      PRINT #RptFile, ""
    NEXT
  END IF
  PRINT #RptFile,
  IF RealFlag THEN
    PRINT #RptFile, ""
    PRINT #RptFile, ""
  ELSE
    PRINT #RptFile, TAB(8); "Please note that all personal property tax must be paid prior to"
    PRINT #RptFile, TAB(8); "purchasing a car decal. "
  END IF
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(8); "If you have any questions, Please contact the Town Hall as soon"
  PRINT #RptFile, TAB(8); "as possible."
  PRINT #RptFile,
  PRINT #RptFile, STRING$(80, "-")
  PRINT #RptFile, TAB(50); USING "AMOUNT NOW DUE ######,#.##"; TBalance#
  PRINT #RptFile, STRING$(80, "-")
  PRINT #RptFile,
  PRINT #RptFile, "NOTE: P/I = PENALTIES OR INTEREST"
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(8); Mess1$
  PRINT #RptFile, TAB(8); Mess2$
  PRINT #RptFile, TAB(8); Mess3$
  PRINT #RptFile, TAB(8); Mess4$
  PRINT #RptFile, TAB(8); Mess5$
  PRINT #RptFile, CHR$(12)
END IF


