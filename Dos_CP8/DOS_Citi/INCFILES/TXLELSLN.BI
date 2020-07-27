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
    PRINT #RptFile, TAB(5); "Our records show that your property taxes to the Town Of Leland, are past"
    PRINT #RptFile, TAB(5); "due and unpaid. By authority of section 105-360 of the North Carolina"
    PRINT #RptFile, TAB(5); "general statutes pursuant to an order by the Leland town council. I plan"
    PRINT #RptFile, TAB(5); "to advertise liens of real estate on "; ADate$; ". The amount"
    PRINT #RptFile, TAB(5); "of your taxes will be increased by a prorated share of the advertisement"
    PRINT #RptFile, TAB(5); "cost. If the taxes remain unpaid, the lien will be foreclosed"
    PRINT #RptFile, TAB(5); "by the Town and the property sold to satisfy the claim for taxes."
    PRINT #RptFile, TAB(5); ""
    PRINT #RptFile, TAB(5); ""
    PRINT #RptFile, TAB(20); "YEAR"; TAB(37); "Amount (Including Interest)"
  END IF

  IF NOT DoFooter THEN
    PRINT #RptFile, TAB(9); TaxTransRec.TaxYear;
    PRINT #RptFile, TAB(20); TaxTransRec.Description
    PRINT #RptFile, TAB(55); USING "$$#######.##"; Balance#

    'PRINT #RptFile, TAB(19); TaxTransRec.TaxYear; TAB(40); USING "$$#######.##"; Balance#
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
    PRINT #RptFile, TAB(5); "not made before the first day of the next month. However, If you pay"
    PRINT #RptFile, TAB(5); "the full amount by "; PDate$; ", Your name will be"
    PRINT #RptFile, TAB(5); "removed from the list to be advertised."
    PRINT #RptFile,
    PRINT #RptFile, TAB(5); "Please call if you have any questions about this notice."
    PRINT #RptFile, ""
    PRINT #RptFile,
    PRINT #RptFile, TAB(5); "Thank You,"
    PRINT #RptFile, ""
    PRINT #RptFile, TAB(5); "Tax Department"
    PRINT #RptFile, CHR$(12)
  END IF

