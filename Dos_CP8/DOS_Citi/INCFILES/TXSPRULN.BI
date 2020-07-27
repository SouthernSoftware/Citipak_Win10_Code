
  IF DoneHeader = False THEN
    PLineCnt = 0
    DoneHeader = True
    PRINT #RptFile, ""
    PRINT #RptFile, ""
    PRINT #RptFile, TAB(34); QPTrim$(TaxSetUpRec.Name)
    PRINT #RptFile, TAB(34); QPTrim$(TaxSetUpRec.Add1)
    PRINT #RptFile, TAB(34); QPTrim$(TaxSetUpRec.Add2)
    PRINT #RptFile, TAB(34); QPTrim$(TaxSetUpRec.Add3)
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile, TAB(5); CustName$; TAB(50); MakeLongDate$(DATE$)
    PRINT #RptFile, TAB(5); LEFT$(TaxCustRec.Addr1, 25)
    PRINT #RptFile, TAB(5); LEFT$(TaxCustRec.Addr2, 25)
    PRINT #RptFile, TAB(5); RTRIM$(TaxCustRec.City) + " "; TaxCustRec.State + " " + LEFT$(TaxCustRec.Zip, 5) + "-" + MID$(TaxCustRec.Zip, 6, 4); TAB(50); USING "Account: #####"; CustAcct&
    PRINT #RptFile,
    PRINT #RptFile,
    PRINT #RptFile, TAB(5); "Dear Taxpayer:"
    PRINT #RptFile,
    PRINT #RptFile, TAB(5); "Our records show that the following taxes are past due."
    PRINT #RptFile, ""
    PRINT #RptFile, ""
    PRINT #RptFile, ""
'***********************
'    PRINT #RptFile, TAB(5); "Our records show that your property taxes are past due. The Town"
'    PRINT #RptFile, TAB(5); "will advertise delinquent taxes on "; ADate$; ". If your"
'    PRINT #RptFile, TAB(5); "taxes remain unpaid by "; PDate$; ", wage garnishement"
'    PRINT #RptFile, TAB(5); "and/or bank accounts will be attached."
'    PRINT #RptFile, TAB(5); ""
'    PRINT #RptFile, TAB(5); ""
'    PRINT #RptFile, TAB(20); "YEAR"; TAB(37); "Amount (Including Interest)"
'****************************
  END IF

  IF NOT DoFooter THEN
    'PRINT #RptFile, TAB(9); TaxTransRec.TaxYear;
    'PRINT #RptFile, TAB(20); TaxTransRec.Description
    'PRINT #RptFile, TAB(55); USING "$$#######.##"; Balance#

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
    PRINT #RptFile, TAB(5); "We would appreciate your prompt attention to this matter. Tax and Interest"
    PRINT #RptFile, TAB(5); "amounts are good thru "; PDate$
    PRINT #RptFile,
    PRINT #RptFile, TAB(5); "If you have any questions please call Monday through Friday"
    PRINT #RptFile, TAB(5); "8:00am to 4:30pm.   PHONE: (828)765-3000"
    PRINT #RptFile, TAB(5); ""
    PRINT #RptFile, TAB(5); "Thank You,"
    PRINT #RptFile, ""
    PRINT #RptFile, TAB(5); "Tax Department"
    PRINT #RptFile, CHR$(12)
  END IF

