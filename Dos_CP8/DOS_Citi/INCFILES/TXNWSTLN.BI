  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, TAB(30); "CITY OF NORTHWEST"
  PRINT #RptFile, TAB(32); "P.O. BOX 1509"
  PRINT #RptFile, TAB(30); "LELAND, NC 28451"
  PRINT #RptFile, TAB(32); "(910)655-3110"
  PRINT #RptFile, ""
  PRINT #RptFile,
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
  PRINT #RptFile, " THE NORTHWEST CITY COUNCIL HAS DIRECTED ME TO ADVERTISE DELINQUENT TAXES IN"
  PRINT #RptFile, " THE NEWSPAPER. YOUR NAME, TAX PARCEL AND AMOUNT OWED WILL BE ADVERTISED"
  PRINT #RptFile, " IN THE BRUNSWICK BEACON ON "; ADate$; "."
  PRINT #RptFile, ""
  PRINT #RptFile, " TO AVOID ADVERTISING COST, YOU MUST PAY BEFORE "; PDate$; "."
  PRINT #RptFile, " INTEREST ACCRUES ON BALANCE DUE THE FIRST OF EACH MONTH."
  PRINT #RptFile, " THE OFFICE IS OPEN MONDAY THROUGH FRIDAY, 9 A.M. TO 5 P.M."
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, USING "  Balance: #####.##"; Balance#;
  PRINT #RptFile, "      Bill Year: "; BillYear
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, " SINCERELY,"
  PRINT #RptFile, ""
  PRINT #RptFile, ""
  PRINT #RptFile, " GWEN SPICER "
  PRINT #RptFile, " TAX COLLECTOR"
  PRINT #RptFile, CHR$(12)

