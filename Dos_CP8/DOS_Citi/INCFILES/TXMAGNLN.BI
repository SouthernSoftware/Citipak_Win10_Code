  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(31); "TOWN OF MAGNOLIA"
  PRINT #RptFile, TAB(33); "P.O. BOX 459"
  PRINT #RptFile, TAB(30); "MAGNOLIA, NC 28453"
  PRINT #RptFile, ""
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); MakeLongDate$(DATE$)
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); CustName$
  PRINT #RptFile, TAB(5); LEFT$(TaxCustRec.Addr1, 25)
  PRINT #RptFile, TAB(5); LEFT$(TaxCustRec.Addr2, 25)
  PRINT #RptFile, TAB(5); RTRIM$(TaxCustRec.City) + " "; TaxCustRec.State + " " + LEFT$(TaxCustRec.Zip, 5) + "-" + MID$(TaxCustRec.Zip, 6, 4)

  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, TAB(5); "The Magnolia Board of Commissioners has directed me to advertise"
  PRINT #RptFile, TAB(5); "delinquent taxes in the newspaper. Your name, tax parcel and amount"
  PRINT #RptFile, TAB(5); "owed will be published in the Wallace Enterprise on "; ADate$; "."
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(5); "To avoid advertising cost, you must pay before "; PDate$; "."
  PRINT #RptFile, TAB(5); "The Town Hall is open Monday through Friday, 8 A.M. TO 5 P.M."
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(5); "SINCERELY,"
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(5); "Katrina Miller"
  PRINT #RptFile, TAB(5); "Tax Collector"
  PRINT #RptFile, CHR$(12)

