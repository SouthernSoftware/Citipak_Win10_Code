'This is "Whitakers" check format 04/18/96
     TempEmp$ = LEFT$(Check(1).EmpName, 33)
     PRINT #1, '"Top line                                                       Top line"
     PRINT #1, TAB(8); TempEmp$; Num2Date$(Check(1).CheckDate); "  "; USING "Rate:####.##"; Check(1).BaseRate
     PRINT #1,
     PRINT #1, TAB(62); "Tax Frng"; TAB(76); "EIC"

     TempEmp$ = SPACE$(6)
     RSET TempEmp$ = LTRIM$(Check(1).EMPNO)
     PRINT #1, TempEmp$;

     PRINT #1, TAB(9); USING "###.##"; Check(1).OTHrsPaid;
     PRINT #1, TAB(16); USING "###.##"; Check(1).HolUsed;
     PRINT #1, TAB(24); USING "###.##"; Check(1).CompUsed;
     PRINT #1, TAB(34); USING "####.##"; Check(1).TotRegWage;
     PRINT #1, TAB(61); USING "#####.##"; Check(1).TaxFring;
     PRINT #1, TAB(73); USING "####.##"; Check(1).EICAmt

     TempEmp$ = SPACE$(10)
     RSET TempEmp$ = QPTrim$(Check(1).AEarn(1).DCode)
     PRINT #1, TAB(50); TempEmp$;

     RSET TempEmp$ = QPTrim$(Check(1).AEarn(2).DCode)
     PRINT #1, TAB(60); TempEmp$;

     RSET TempEmp$ = QPTrim$(Check(1).AEarn(3).DCode)
     PRINT #1, TAB(70); TempEmp$

     PRINT #1, TAB(3); USING "###.##"; Check(1).RegHrsWork;
     PRINT #1, TAB(9); USING "###.##"; Check(1).VacUsed;
     PRINT #1, TAB(16); USING "###.##"; Check(1).SickUsed;
     PRINT #1, TAB(24); USING "###.##"; Check(1).RegHrsPaid;
     PRINT #1, TAB(34); USING "####.##"; Check(1).TotOTWage;

     PRINT #1, TAB(52); USING "####.##"; Check(1).AEarn(1).DAmt;
     PRINT #1, TAB(62); USING "####.##"; Check(1).AEarn(2).DAmt;
     PRINT #1, TAB(73); USING "####.##"; Check(1).AEarn(3).DAmt
     PRINT #1,
     PRINT #1,

     PRINT #1, TAB(27); LEFT$(Check(1).CDED(1).DCode, 6);
     PRINT #1, TAB(38); LEFT$(Check(1).CDED(2).DCode, 6);
     PRINT #1, TAB(49); LEFT$(Check(1).CDED(3).DCode, 6);
     PRINT #1, TAB(60); LEFT$(Check(1).CDED(4).DCode, 6);
     PRINT #1, TAB(73); USING "####.##"; Check(1).GrossPay

     PRINT #1, TAB(8); USING "####.##"; Round#(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
     PRINT #1, TAB(19); USING "####.##"; Check(1).StaTaxAmt;

     PRINT #1, TAB(29); USING "####.##"; Check(1).CDED(1).DAmt;
     PRINT #1, TAB(40); USING "####.##"; Check(1).CDED(2).DAmt;
     PRINT #1, TAB(51); USING "####.##"; Check(1).CDED(3).DAmt;
     PRINT #1, TAB(62); USING "####.##"; Check(1).CDED(4).DAmt

     PRINT #1, TAB(27); LEFT$(Check(1).CDED(5).DCode, 6);
     PRINT #1, TAB(38); LEFT$(Check(1).CDED(6).DCode, 6);
     PRINT #1, TAB(49); LEFT$(Check(1).CDED(7).DCode, 6);
     PRINT #1, TAB(60); LEFT$(Check(1).CDED(8).DCode, 6);
     PRINT #1, TAB(73); USING "####.##"; Round#(Check(1).MedTaxAmt + Check(1).SocTaxAmt + Check(1).FedTaxAmt + Check(1).StaTaxAmt)

     PRINT #1, TAB(8); USING "####.##"; Check(1).FedTaxAmt;
     PRINT #1, TAB(19); USING "####.##"; Check(1).RetireAmt;

     PRINT #1, TAB(29); USING "####.##"; Check(1).CDED(5).DAmt;
     PRINT #1, TAB(40); USING "####.##"; Check(1).CDED(6).DAmt;
     PRINT #1, TAB(51); USING "####.##"; Check(1).CDED(7).DAmt;
     PRINT #1, TAB(62); USING "####.##"; Check(1).CDED(8).DAmt

     PRINT #1, TAB(5); "YTD Ret";
     PRINT #1, TAB(27); LEFT$(Check(1).CDED(9).DCode, 6);
     PRINT #1, TAB(38); LEFT$(Check(1).CDED(10).DCode, 6);
     PRINT #1, TAB(49); LEFT$(Check(1).CDED(11).DCode, 6);
     PRINT #1, TAB(60); LEFT$(Check(1).CDED(12).DCode, 6)

     PRINT #1, TAB(8); USING "####.##"; Round(Check(1).YTDRetire + Check(1).RetireAmt);
     PRINT #1, TAB(29); USING "####.##"; Check(1).CDED(9).DAmt;
     PRINT #1, TAB(40); USING "####.##"; Check(1).CDED(10).DAmt;
     PRINT #1, TAB(51); USING "####.##"; Check(1).CDED(11).DAmt;
     PRINT #1, TAB(62); USING "####.##"; Check(1).CDED(12).DAmt

     PRINT #1, TAB(73); USING "####.##"; Check(1).TotDedAmt
     PRINT #1, TAB(17); USING "######.##"; Check(1).YTDGrossPay;
     PRINT #1, TAB(40); USING "#####.##"; Check(1).YTDFederal

     PRINT #1, TAB(18); USING "#####.##"; Round(Check(1).YTDSocial + Check(1).YTDMedicare);
     PRINT #1, TAB(40); USING "#####.##"; Check(1).YTDState;
     'PRINT #1, TAB(61); Num2Date$(Check(1).CheckDate);
     PRINT #1, TAB(72); USING "#####.##"; Check(1).NetPay

     PRINT #1, TAB(4); "Vac. Due: "; USING "####.###"; Check(1).VactBal;
     PRINT #1, TAB(30); "Sick Due: "; USING "####.###"; Check(1).SickBal;
     PRINT #1, TAB(56); "Comp Due: "; USING "####.###"; Check(1).CompBal
     PRINT #1,
     PRINT #1,

     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(59); Num2Date$(Check(1).CheckDate);
     PRINT #1, TAB(70); Check(1).CheckNum
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(68); USING "$$#####.##"; Check(1).NetPay
     PRINT #1, TAB(12); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##"));
     'Print the whole number part
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(12); Check(1).EmpName
     PRINT #1, TAB(12); Check(1).EmpAddr1
     PRINT #1, TAB(12); Check(1).EmpCity; " "; Check(1).EmpState; " "; Check(1).EmpZip
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, '"Bottom Line                                                 Bottom Line"
'-=-=-=-=-=-=-=-=-=-=-*

