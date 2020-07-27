  IF DoneHeader = False THEN
    PLineCnt = 0
    DoneHeader = True
    PRINT #RptFile, ""
    PRINT #RptFile, ""
    PRINT #RptFile, TAB(34); " TOWN OF MURPHY"
    PRINT #RptFile, TAB(34); " P. O. BOX 130"
    PRINT #RptFile, TAB(34); "MURPHY, NC 28906"
    PRINT #RptFile,
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
    PRINT #RptFile, TAB(5); "past due and unpaid. I have been instructed by the Board of Commissioners"
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
    PRINT #RptFile, TAB(5); "Additional interest will be added as provided by the law if payment is"
    PRINT #RptFile, TAB(5); "not made before the first day of the next month."
    PRINT #RptFile,
    PRINT #RptFile, TAB(5); "Unless your taxes are paid, or a satisfactory arrangement is made with"
    PRINT #RptFile, TAB(5); "regard to payment within ten (10) days of the date of this letter, I am"
    PRINT #RptFile, TAB(5); "instructed to institute attachment and/or garnishment proceeding without"
    PRINT #RptFile, TAB(5); "further notice. Please take care of this matter within the time indicated"
    PRINT #RptFile, TAB(5); "to avoid additional costs, or having your property sold."
    PRINT #RptFile,
    PRINT #RptFile, TAB(5); "If you no longer own this property, Please contact us with the name of"
    PRINT #RptFile, TAB(5); "the responsible party."
    PRINT #RptFile, ""
    PRINT #RptFile, ""
    PRINT #RptFile, TAB(5); "SINCERELY,"
    PRINT #RptFile, ""
    PRINT #RptFile, ""
    PRINT #RptFile, TAB(5); "Mary George Moore"
    PRINT #RptFile, TAB(5); "Tax Collector"
    PRINT #RptFile, TAB(5); "(828)837-2510"
    PRINT #RptFile, CHR$(12)
  END IF

