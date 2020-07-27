  IF DoneHeader = False THEN
    PLineCnt = 0
    DoneHeader = True
    PRINT #RptFile, ""
    PRINT #RptFile, ""
    PRINT #RptFile, TAB(32); QPTrim$(TaxSetUpRec.Name)
    PRINT #RptFile, TAB(36); QPTrim$(TaxSetUpRec.Add1)
    PRINT #RptFile, TAB(31); QPTrim$(TaxSetUpRec.Add2)
    PRINT #RptFile, TAB(30); QPTrim$(TaxSetUpRec.Add3)
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile, TAB(5); CustName$; TAB(50); MakeLongDate$(DATE$)
    PRINT #RptFile, TAB(5); LEFT$(TaxCustRec.Addr1, 25)
    PRINT #RptFile, TAB(5); LEFT$(TaxCustRec.Addr2, 25)
    PRINT #RptFile, TAB(5); RTRIM$(TaxCustRec.City) + " "; TaxCustRec.State + " " + LEFT$(TaxCustRec.Zip, 5) + "-" + MID$(TaxCustRec.Zip, 6, 4); TAB(50); USING "Account: #####"; CustAcct&
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile, TAB(5); "Dear Customer,"
    PRINT #RptFile,
    PRINT #RptFile, TAB(5); "Our records show that your property taxes for the years listed below are"
    PRINT #RptFile, TAB(5); "past due and unpaid. I have been instructed by the Board of Aldermen"
    PRINT #RptFile, TAB(5); "to collect them without delay, or to institute attachment and garnishment"
    PRINT #RptFile, TAB(5); "proceedings."
    PRINT #RptFile, ""
    PRINT #RptFile, ""
    PRINT #RptFile, TAB(20); "YEAR"; TAB(37); "Amount (Including Interest)"
  END IF
  IF NOT DoFooter THEN
    PRINT #RptFile, TAB(19); TaxTransRec.TaxYear; TAB(40); USING "$$#######.##"; Balance#
    PLineCnt = PLineCnt + 1
    'PRINT #RptFile, TAB(45); TaxTransRec.Description
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
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile, TAB(5); "Thank You,"
    PRINT #RptFile, ""
    PRINT #RptFile, ""
    PRINT #RptFile, TAB(5); "Town of Maggie Valley"
    PRINT #RptFile, TAB(5); "Tax Department"
    PRINT #RptFile,
    PRINT #RptFile, CHR$(12)
  END IF

