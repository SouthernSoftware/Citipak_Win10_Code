    Dash$ = STRING$(30, "_")
    LessBase$ = Dash$
    TLen = LEN(TownName$)
    IF TLen < 2 THEN
      PRINT "error": END
    ELSE
      MID$(LessBase$, ((LEN(LessBase$) / 2) - (TLen / 2))) = TownName$
    END IF

    BaseFee$ = Dash$
    TT$ = QPTrim$(Form$(5, 0))
    TLen = LEN(TT$)
    IF TLen < 2 THEN
      PRINT "error": END
    ELSE
      MID$(BaseFee$, ((LEN(BaseFee$) / 2) - (TLen / 2))) = TT$
    END IF

    MultiBY$ = QPTrim$(Form$(6, 0))

    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(32); "TOWN OF ELLOREE"
    PRINT #RptHandle, TAB(34); "P.O. BOX 28"
    PRINT #RptHandle, TAB(30); "ELLOREE, S.C. 29047"
    PRINT #RptHandle, TAB(20); "APPLICATION FOR BUSINESS LICENSE - "; Form$(2, 0)
    PRINT #RptHandle, ""
    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(5); ARCustRec(1).BILLNAME
    PRINT #RptHandle, TAB(5); ARCustRec(1).Address1
    PRINT #RptHandle, TAB(5); ARCustRec(1).Address2
    PRINT #RptHandle, TAB(5); QPTrim$(ARCustRec(1).City); ", "; ARCustRec(1).State; " "; ARCustRec(1).ZipCode
    PRINT #RptHandle, ""
    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(5); "To engage in business or profession, make a separate application"
    PRINT #RptHandle, TAB(5); "for each business and each location.  Send fee with application to"
    PRINT #RptHandle, TAB(5); "The Town of Elloree:"
    PRINT #RptHandle, ""
    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(5); "         Owners Name:______________________________________________"
    PRINT #RptHandle, TAB(5); "Business Description:______________________________________________"
    PRINT #RptHandle, TAB(5); "      Business Phone:______________________________________________"
    PRINT #RptHandle, TAB(5); "   Federal ID Number:______________________________________________"
    PRINT #RptHandle, TAB(5); "     State ID Number:______________________________________________"
    PRINT #RptHandle, TAB(5); "___________________________________________________________________"
    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(5); "To calculate your TOWN OF ELLOREE Business License Fee, Use the"
    PRINT #RptHandle, TAB(5); "formula bellow."
    PRINT #RptHandle, ""
    PRINT #RptHandle, TAB(5); "1.  Gross Sales"; TAB(40); Dash$
    PRINT #RptHandle, TAB(5); "2.  Less Base Amount"; TAB(40); LessBase$
    PRINT #RptHandle, TAB(5); "3.  Excess Gross"; TAB(40); Dash$
    PRINT #RptHandle, TAB(5); "4.  Base Rate Fee"; TAB(40); BaseFee$
    PRINT #RptHandle, TAB(5); "5.  If No. 3 is Greater than"
    PRINT #RptHandle, TAB(5); "    Zero, divide No. 3 by 1,000"
    PRINT #RptHandle, TAB(5); "    and Round UP"; TAB(40); Dash$
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(5); "6.  Multiply #5 by "; MultiBY$; TAB(40); Dash$
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(5); "7.  Total License Fee"
    PRINT #RptHandle, TAB(5); "    Number 4 plus number 6"; TAB(40); Dash$
    PRINT #RptHandle, TAB(5); "8.  Add  penalty ($10.00 Collector's"
    PRINT #RptHandle, TAB(5); "    Fee and 5% per month after"
    PRINT #RptHandle, TAB(5); "    April 30th)"; TAB(40); Dash$
    PRINT #RptHandle, TAB(5); "9.  TOTAL DUE (No. 7 + No. 8)"; TAB(40); Dash$
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(5); "This is to certify that the amount of total gross for the business"
    PRINT #RptHandle, TAB(5); "transacted at or through the above location for the calendar year"
    PRINT #RptHandle, TAB(5); "ending December 31, or the last complete fiscal year is true and"
    PRINT #RptHandle, TAB(5); "correct, and that this report corresponds with the amount that was"
    PRINT #RptHandle, TAB(5); "reported to the SC Tax Commission or Insurance Commission and with"
    PRINT #RptHandle, TAB(5); "the Internal Revenue Service."
    PRINT #RptHandle, TAB(5); ""
    PRINT #RptHandle, TAB(5); ""
    PRINT #RptHandle, TAB(5); Dash$; TAB(40); Dash$
    PRINT #RptHandle, TAB(5); "Firm Name/ Individual Signature"; TAB(40); "By:"
    PRINT #RptHandle, CHR$(12);
    TotalCust = TotalCust + 1

