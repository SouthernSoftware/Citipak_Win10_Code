  PRINT #RptFile, "Town of Troy                               First Class"
  PRINT #RptFile, "Tax Department                             Mail - U.S."
  PRINT #RptFile, "315 N. Main St.                            Troy, NC 27371"
  PRINT #RptFile, "Troy NC 27371                              Permit No. 37"
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(20); CustName$
  PRINT #RptFile, TAB(20); LEFT$(TaxCustRec.Addr1, 25)
  PRINT #RptFile, TAB(20); LEFT$(TaxCustRec.Addr2, 25)
  PRINT #RptFile, TAB(20); RTRIM$(TaxCustRec.City) + " "; TaxCustRec.State + " " + LEFT$(TaxCustRec.Zip, 5) + "-" + MID$(TaxCustRec.Zip, 6, 4)
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, "         Account: "; CustAcct&
  PRINT #RptFile, "     Bill Number: "; ParseBillNum$(TaxTransRec.Description)
  PRINT #RptFile, USING "  Tax Due (Less Interest): ######,#.##"; Balance#
  PRINT #RptFile,
  PRINT #RptFile, "Taxes will advertise on "; AdvDate$; ". To AVOID advertising"
  PRINT #RptFile, "cost, taxes must be paid in FULL by "; PayDate$; "."
  PRINT #RptFile, "~"

