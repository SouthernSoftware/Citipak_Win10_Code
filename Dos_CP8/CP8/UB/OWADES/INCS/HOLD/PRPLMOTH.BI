'This is "Plymouth" check format 7/08/94

     Fmt$ = "#####.##"
     Ten$ = SPACE$(10)

     PRINT #1, "TOP"
     PRINT #1,
     PRINT #1, "  "; LTRIM$(Check(1).EMPNO); TAB(10); LEFT$(Check(1).EmpName, 26); Check(1).EmpSSN;
     PRINT #1, " "; Num2Date(Check(1).PayEndDate); " "; Num2Date(Check(1).CheckDate);
     PRINT #1, "  "; Check(1).CheckNum
     PRINT #1,
     PRINT #1,
     PRINT #1, " "; USING "#####.##"; Check(1).RegHrsWork; TAB(11); Check(1).BaseRate;
     PRINT #1, "  "; USING "#####.##"; Check(1).TotRegWage; Check(1).OTHrsPaid; TAB(45); Check(1).TotOTWage; TAB(60); Check(1).TotAdditEarn; TAB(72); Check(1).GrossPay
     PRINT #1,

'----
'     PRINT #1, "  "; USING "COMP USE ####.##"; Check(1).CompUsed;
'     PRINT #1, USING "  VACT USE ####.##"; Check(1).VacUsed;
'     PRINT #1, USING "  SICK USE ####.##"; Check(1).SickUsed
     PRINT #1,

'     PRINT #1, "  "; USING "COMP BAL ####.##"; Check(1).CompBal;
'     PRINT #1, USING "  VACT BAL ####.##"; Check(1).VactBal;
'     PRINT #1, USING "  SICK BAL ####.##"; Check(1).SickBal
     PRINT #1,

     PRINT #1,
     PRINT #1, USING "#####.##"; Check(1).RetireAmt; TAB(11); Check(1).FedTaxAmt; Round#(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
     PRINT #1, " "; USING "#####.##"; Check(1).StaTaxAmt;

     PRINT #1, TAB(38);
     FOR Ded = 1 TO 3
       LSET Ten$ = QPTrim$(STR$(Ded))
       MID$(Ten$, 3) = FUsing$(STR$(Check(1).CDED(Ded).DAmt), Fmt$)
       PRINT #1, Ten$; " ";
     NEXT

     PRINT #1, TAB(72); USING "#####.##"; Check(1).NetPay
     PRINT #1, TAB(38);
     FOR Ded = 4 TO 6
       LSET Ten$ = QPTrim$(STR$(Ded))
       MID$(Ten$, 3) = FUsing$(STR$(Check(1).CDED(Ded).DAmt), Fmt$)
       PRINT #1, Ten$; " ";
     NEXT

     PRINT #1,
     PRINT #1, TAB(38);
     FOR Ded = 7 TO 9
       LSET Ten$ = QPTrim$(STR$(Ded))
       MID$(Ten$, 3) = FUsing$(STR$(Check(1).CDED(Ded).DAmt), Fmt$)
       PRINT #1, Ten$; " ";
     NEXT

     PRINT #1,
     PRINT #1, TAB(38);
     FOR Ded = 10 TO 12
       LSET Ten$ = QPTrim$(STR$(Ded))
       MID$(Ten$, 3) = FUsing$(STR$(Check(1).CDED(Ded).DAmt), Fmt$)
       PRINT #1, Ten$; " ";
     NEXT
     PRINT #1,

     PRINT #1,
     PRINT #1,
     PRINT #1, USING "#####.##"; TAB(5); Check(1).YTDFederal; TAB(15); Round(Check(1).YTDSocial + Check(1).YTDMedicare); Check(1).YTDState;

     PRINT #1, TAB(38);
     FOR Ded = 1 TO 3
       LSET Ten$ = QPTrim$(STR$(Ded))
       MID$(Ten$, 3) = FUsing$(STR$(Check(1).CDED(Ded).YTDDAmt), Fmt$)
       PRINT #1, Ten$; " ";
     NEXT
     PRINT #1, USING "#####.##"; TAB(72); Check(1).YTDGrossPay

     PRINT #1, TAB(38);
     FOR Ded = 4 TO 6
       LSET Ten$ = QPTrim$(STR$(Ded))
       MID$(Ten$, 3) = FUsing$(STR$(Check(1).CDED(Ded).YTDDAmt), Fmt$)
       PRINT #1, Ten$; " ";
     NEXT

     PRINT #1,
     PRINT #1, TAB(38);
     FOR Ded = 7 TO 9
       LSET Ten$ = QPTrim$(STR$(Ded))
       MID$(Ten$, 3) = FUsing$(STR$(Check(1).CDED(Ded).YTDDAmt), Fmt$)
       PRINT #1, Ten$; " ";
     NEXT

     PRINT #1,
     PRINT #1, TAB(38);
     FOR Ded = 10 TO 12
       LSET Ten$ = QPTrim$(STR$(Ded))
       MID$(Ten$, 3) = FUsing$(STR$(Check(1).CDED(Ded).YTDDAmt), Fmt$)
       PRINT #1, Ten$; " ";
     NEXT
     PRINT #1,

     'PRINT #1,
     'PRINT #1,
     'PRINT #1,

     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(49); Num2Date(Check(1).CheckDate); TAB(61); Check(1).CheckNum; TAB(72); USING "####.##"; Check(1).NetPay
     PRINT #1, TAB(6); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##"))  'Print the whole number part
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(10); Check(1).EmpName
     PRINT #1, TAB(10); Check(1).EmpAddr1
     PRINT #1, TAB(10); Check(1).EmpCity; " "; Check(1).EmpState; " "; Check(1).EmpZip
'    GOTO skip
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, "BOTTOM"
'-=-=-=-=-=-=-=-=-=-=-*

