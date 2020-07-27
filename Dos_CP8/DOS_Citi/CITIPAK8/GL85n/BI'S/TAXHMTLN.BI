  PRINT #RptFile, "~"
  PRINT #RptFile, TAB(14); USING "######"; CustAcct&
  PRINT #RptFile,
  PRINT #RptFile, TAB(22); PDesc$
  PRINT #RptFile,
  PRINT #RptFile,
  PRINT #RptFile,
  FOR Pnt = 1 TO 7
  IF TaxAmt#(Pnt) > 0 OR OthAmt#(Pnt) > 0 THEN
   PRINT #RptFile, TAB(47); TaxYear$(Pnt);
    PRINT #RptFile, TAB(52); LEFT$(TaxRecpt$(Pnt), 5);
    PRINT #RptFile, TAB(58); USING "#####,#.##"; TaxAmt#(Pnt);
    PRINT #RptFile, TAB(70); USING "###,#.##"; OthAmt#(Pnt);
    PRINT #RptFile, TAB(79); USING "$$#####,#.##"; (TaxAmt#(Pnt) + OthAmt#(Pnt))
    GTotal# = GTotal# + (TaxAmt#(Pnt) + OthAmt#(Pnt))
   ELSE
   PRINT #RptFile, ""
  END IF

  NEXT Pnt
  PRINT #RptFile, TAB(10); LEFT$(CustName$, 35);
   IF TaxAmt#(8) > 0 THEN
   PRINT #RptFile, TAB(47); TaxYear$(Pnt);
    PRINT #RptFile, TAB(58); USING "#####,#.##"; TaxAmt#(Pnt);
    PRINT #RptFile, TAB(70); USING "###,#.##"; OthAmt#(Pnt);
    PRINT #RptFile, TAB(79); USING "$$#####,#.##"; (TaxAmt#(Pnt) + OthAmt#(Pnt))
    GTotal# = GTotal# + (TaxAmt#(Pnt) + OthAmt#(Pnt))
   END IF
  PRINT #RptFile, TAB(10); LEFT$(TaxCustRec.Addr1, 35);
   IF TaxAmt#(9) > 0 THEN
   PRINT #RptFile, TAB(47); TaxYear$(Pnt);
    PRINT #RptFile, TAB(58); USING "#####,#.##"; TaxAmt#(Pnt);
    PRINT #RptFile, TAB(70); USING "###,#.##"; OthAmt#(Pnt);
    PRINT #RptFile, TAB(79); USING "$$#####,#.##"; (TaxAmt#(Pnt) + OthAmt#(Pnt))
    GTotal# = GTotal# + (TaxAmt#(Pnt) + OthAmt#(Pnt))
   END IF
  PRINT #RptFile, TAB(10); LEFT$(TaxCustRec.Addr2, 35);
   IF TaxAmt#(10) > 0 THEN
   PRINT #RptFile, TAB(47); TaxYear$(Pnt);
    PRINT #RptFile, TAB(58); USING "#####,#.##"; TaxAmt#(Pnt);
    PRINT #RptFile, TAB(70); USING "###,#.##"; OthAmt#(Pnt);
    PRINT #RptFile, TAB(79); USING "$$#####,#.##"; (TaxAmt#(Pnt) + OthAmt#(Pnt))
    GTotal# = GTotal# + (TaxAmt#(Pnt) + OthAmt#(Pnt))
   END IF
  PRINT #RptFile, TAB(10); RTRIM$(TaxCustRec.City) + " "; TaxCustRec.State + " " + LEFT$(TaxCustRec.Zip, 5) + "-" + MID$(TaxCustRec.Zip, 6, 4);
   PRINT #RptFile, TAB(68); "TOTAL DUE: "; USING "$$#####,#.##"; GTotal#
   GTotal# = 0
  PRINT #RptFile, TAB(48); STRING$(40, "-")
  PRINT #RptFile, TAB(48); Mess1$
  PRINT #RptFile, TAB(48); Mess2$
  PRINT #RptFile, "~"




