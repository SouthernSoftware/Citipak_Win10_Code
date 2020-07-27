'This is the "ROLESVILLE" check format 02/16/94
     TempEmp$ = LEFT$(Check(1).EmpName, 33)
     'PRINT #1, CHR$(27) + CHR$(58);'sets compressed mode   12 dpi
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(8); QPTrim$(TempEmp$); TAB(38); USING "Rate: ####.##"; Check(1).BaseRate;
     PRINT #1, TAB(54); "SSN: "; Check(1).EmpSSN
     PRINT #1,
     PRINT #1, TAB(65); "EIC"; TAB(70); QPTrim$(Check(1).AEarn(1).DCode)
     
     PRINT #1, TAB(1); LEFT$(LTRIM$(Check(1).EMPNO), 6);
     PRINT #1, TAB(8); USING "###.##"; Check(1).OTHrsPaid;
     PRINT #1, TAB(15); USING "###.##"; Check(1).HolUsed;
                                         
     PRINT #1, TAB(32); USING "####.##"; Check(1).TotRegWage; TAB(41); Round(Check(1).BaseRate * Check(1).VacUsed);
     PRINT #1, TAB(50); USING "####.##"; Round(Check(1).BaseRate * Check(1).SickUsed);

     PRINT #1, TAB(61); USING "####.##"; Check(1).EICAmt;
     PRINT #1, TAB(72); USING "####.##"; Check(1).AEarn(1).DAmt


     PRINT #1, TAB(59); QPTrim$(Check(1).AEarn(2).DCode);
     PRINT #1, TAB(70); QPTrim$(Check(1).AEarn(3).DCode)

     PRINT #1, TAB(1); USING "###.##"; Check(1).RegHrsWork;
     PRINT #1, TAB(8); USING "###.##"; Check(1).VacUsed;
     PRINT #1, TAB(15); USING "###.##"; Check(1).SickUsed;
     PRINT #1, TAB(32); USING "####.##"; Check(1).TotOTWage;
     PRINT #1, USING "####.##"; TAB(41); Round(Check(1).BaseRate * Check(1).HolUsed);

     PRINT #1, TAB(61); USING "####.##"; Check(1).AEarn(2).DAmt;
     PRINT #1, TAB(72); USING "####.##"; Check(1).AEarn(3).DAmt

     PRINT #1,
     PRINT #1,

     PRINT #1, TAB(25); Check(1).CDED(1).DCode;
     PRINT #1, TAB(36); Check(1).CDED(2).DCode;
     PRINT #1, TAB(47); Check(1).CDED(3).DCode;
     PRINT #1, TAB(58); Check(1).CDED(4).DCode;
     PRINT #1, TAB(72); USING "####.##"; Check(1).GrossPay

     PRINT #1, TAB(6); USING "####.##"; Round#(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
     PRINT #1, TAB(17); USING "####.##"; Check(1).StaTaxAmt;

     PRINT #1, TAB(28); USING "####.##"; Check(1).CDED(1).DAmt;
     PRINT #1, TAB(39); USING "####.##"; Check(1).CDED(2).DAmt;
     PRINT #1, TAB(50); USING "####.##"; Check(1).CDED(3).DAmt;
     PRINT #1, TAB(61); USING "####.##"; Check(1).CDED(4).DAmt

     PRINT #1, TAB(25); Check(1).CDED(5).DCode;
     PRINT #1, TAB(36); Check(1).CDED(6).DCode;
     PRINT #1, TAB(47); Check(1).CDED(7).DCode;
     PRINT #1, TAB(58); Check(1).CDED(8).DCode;
     PRINT #1, TAB(72); USING "####.##"; Round#(Check(1).MedTaxAmt + Check(1).SocTaxAmt + Check(1).FedTaxAmt + Check(1).StaTaxAmt)

     PRINT #1, TAB(6); USING "####.##"; Check(1).FedTaxAmt;
     PRINT #1, TAB(28); USING "####.##"; Check(1).CDED(5).DAmt;
     PRINT #1, TAB(39); USING "####.##"; Check(1).CDED(6).DAmt;
     PRINT #1, TAB(50); USING "####.##"; Check(1).CDED(7).DAmt;
     PRINT #1, TAB(61); USING "####.##"; Check(1).CDED(8).DAmt

     PRINT #1, TAB(3); "Retirement";
     PRINT #1, TAB(25); Check(1).CDED(9).DCode;
     PRINT #1, TAB(36); Check(1).CDED(10).DCode;
     PRINT #1, TAB(47); Check(1).CDED(11).DCode;
     PRINT #1, TAB(58); Check(1).CDED(12).DCode

     PRINT #1, TAB(6); USING "####.##"; Check(1).RetireAmt;
     PRINT #1, TAB(28); USING "####.##"; Check(1).CDED(9).DAmt;
     PRINT #1, TAB(39); USING "####.##"; Check(1).CDED(10).DAmt;
     PRINT #1, TAB(50); USING "####.##"; Check(1).CDED(11).DAmt;
     PRINT #1, TAB(61); USING "####.##"; Check(1).CDED(12).DAmt

     PRINT #1, TAB(72); USING "####.##"; Check(1).TotDedAmt

     PRINT #1, TAB(19); USING "######.##"; Check(1).YTDGrossPay; TAB(38); Check(1).YTDFederal

     TempEmp$ = Num2Date$(Check(1).CheckDate)
     TempEmp$ = (LEFT$(TempEmp$, 6) + RIGHT$(TempEmp$, 2))
     PRINT #1, TempEmp$;  'actually the check date

     PRINT #1, TAB(19); USING "######.##"; Round(Check(1).YTDSocial + Check(1).YTDMedicare); TAB(38); Check(1).YTDState;
     PRINT #1, TAB(72); USING "####.##"; Check(1).NetPay
     PRINT #1, TAB(11); "Vac. Due: "; USING "####.##"; Check(1).VactBal;
     PRINT #1, TAB(30); "Sick Due: "; USING "####.##"; Check(1).SickBal;
     PRINT #1, TAB(62); "Comp Due: "; USING "####.##"; Check(1).CompBal

     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(61); Num2Date$(Check(1).CheckDate); " "; Check(1).CheckNum
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(68); USING "$$#####.##"; Check(1).NetPay
     PRINT #1, TAB(12); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##")); 'Print the whole number part
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(12); Check(1).EmpName
     PRINT #1, TAB(12); Check(1).EmpAddr1
     PRINT #1, TAB(12); Check(1).EmpCity; " "; Check(1).EmpState; " "; Check(1).EmpZip
     PRINT #1,
     PRINT #1,
     PRINT #1,
'-=-=-=-=-=-=-=-=-=-=-*

