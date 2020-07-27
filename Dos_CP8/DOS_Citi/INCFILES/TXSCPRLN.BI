'  CustName$ = RTRIM$(TaxCust(1).FNAME) + " " + LTRIM$(TaxCust(1).LName)
'  CustName$ = LTRIM$(CustName$)
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  IF RealFlag THEN
    PRINT #RptFile, TAB(21); "NOTICE OF DELINQUENT REAL ESTATE TAXES"
  ELSE
    PRINT #RptFile, TAB(18); "NOTICE OF DELINQUENT PERSONAL PROPERTY TAXES"
  END IF
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); LNDate$
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); TownName$
  PRINT #RptFile, TAB(5); Add1$
  PRINT #RptFile, TAB(5); ADD2$
  PRINT #RptFile, TAB(5); ADD3$
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); CustName$
  PRINT #RptFile, TAB(5); LEFT$(TaxCust(1).Addr1, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxCust(1).Addr2, 25)
  PRINT #RptFile, TAB(5); RTRIM$(TaxCust(1).City) + " "; TaxCust(1).State + " " + LEFT$(TaxCust(1).Zip, 5) + "-" + MID$(TaxCust(1).Zip, 6, 4)
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); USING "Account _#     ##########"; CustAcct&
  PRINT #RptFile, TAB(5); USING "Total Due:    #######.##"; Balance#
  PRINT #RptFile, TAB(5); "Due Date:     "; InterestDate$
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); "The taxes for the above account are now delinquent. The total amount shown"
  PRINT #RptFile, TAB(5); "includes penalties and any interest due. Please make payment to the"
  PRINT #RptFile, TAB(5); TownName$; " as soon as possible to bring your account current. If"
  PRINT #RptFile, TAB(5); "you have any questions regarding your account, please contact the town"
  PRINT #RptFile, TAB(5); "office."
  PRINT #RptFile, CHR$(12);

