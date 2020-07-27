  'Canton
  IF DoneHeader = False THEN
    PLineCnt = 0
    DoneHeader = True
    PRINT #RptFile,
    PRINT #RptFile, TAB(66); DATE$
    PRINT #RptFile, TAB(34); QPTrim$(TaxSetUpRec.Name)
    PRINT #RptFile, TAB(34); "Tax Statement"
    PRINT #RptFile, TAB(34); QPTrim$(TaxSetUpRec.Add1)
    PRINT #RptFile, TAB(34); QPTrim$(TaxSetUpRec.Add2)
    PRINT #RptFile, TAB(34); QPTrim$(TaxSetUpRec.Add3)
    PRINT #RptFile, 'TAB(34); "PHONE: 828-765-3000"
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile, TAB(4);
    PRINT #RptFile, TAB(10); CustName$
    '; TAB(60); " SSN: "; TaxCustRec.CSSN
    PRINT #RptFile, TAB(10); LEFT$(TaxCustRec.Addr1, 25); TAB(60); "Acct: "; CustAcct&
    PRINT #RptFile, TAB(10); LEFT$(TaxCustRec.Addr2, 25) '; TAB(60); "Tele: "; TaxCustRec.HPHONE
    PRINT #RptFile, TAB(10); RTRIM$(TaxCustRec.City) + ", "; TaxCustRec.State + " " + LEFT$(TaxCustRec.Zip, 5) + "-" + MID$(TaxCustRec.Zip, 6, 4)
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile, STRING$(80, "-")
    PRINT #RptFile, "Description"; TAB(32); "Principle"; TAB(53); "Int."; TAB(69); "Balance"
    PRINT #RptFile, STRING$(80, "-")
  END IF
  IF NOT DoFooter THEN
    'PRINT #RptFile, TAB(19); TaxTransRec.TaxYear; TAB(40); USING "$$#######.##"; Balance#
    PLineCnt = PLineCnt + 1
    'PRINT #RptFile, QPTrim$(TaxTransRec.Description); TAB(30); USING "######,#.##"; Round#(Balance# - IntBal#); TAB(45); TxIntC#; TAB(65); Balance#
    PRINT #RptFile, TAB(2); USING "####"; TaxTransRec.TaxYear;
    PRINT #RptFile, " "; LEFT$(QPTrim$(TaxTransRec.Description), 22); TAB(30); USING "######,#.##"; Round#(Balance# - IntBal#); TAB(45); IntBal#; TAB(65); Balance#
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
'
    PRINT #RptFile, TAB(8); "This is your 2005 Tax bill from the Town of"
    PRINT #RptFile, TAB(8); "Columbia. Taxes are due January 1 2006."
    PRINT #RptFile, TAB(8); "Interest 2% after 1-5-06. 3/4% each month after."
    PRINT #RptFile, TAB(8); "Be advised on "; ADate$; " the Town of Columbia"
    PRINT #RptFile, TAB(8); "will begin publicly advertising ALL DELINQUENT TAXES."
    PRINT #RptFile,


'    PRINT #RptFile, TAB(8); "NOTICE: THE TOWN MAY BEGIN GARNISHING WAGES AND"
'    PRINT #RptFile, TAB(8); "ATTACHING BANK ACCOUNTS FOR DELINQUENT TAXES."
'    PRINT #RptFile, TAB(8); "INTEREST GOOD THROUGH "; PDate$

    PRINT #RptFile, STRING$(80, "-")
    PRINT #RptFile, TAB(50); USING "AMOUNT NOW DUE ######,#.##"; TBalance#
    PRINT #RptFile, STRING$(80, "-")
    PRINT #RptFile, CHR$(12)
  END IF


