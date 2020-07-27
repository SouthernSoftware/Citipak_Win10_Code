'This is the "STANDARD" check format 6/27/94
     TempEmp$ = LEFT$(Check(1).EmpName, 33)
     'PRINT #1, CHR$(27) + CHR$(58);'sets compressed mode   12 dpi
     'PRINT #1, "Top line                                                       Top line"
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(2); TempEmp$; TAB(66); USING "Rate: ####.##"; Check(1).BaseRate
     PRINT #1,
     PRINT #1, TAB(65); "Tax Frng"; TAB(76); "EIC"

     PRINT #1, TAB(1); LEFT$(LTRIM$(Check(1).EMPNO), 5);

     'IF QPTrim$(Check(1).EmpNo) = "6" THEN STOP
     PRINT #1, TAB(7); USING "###.##"; Check(1).RegHrsWork;
'-new
     PRINT #1, TAB(15); USING "###.##"; Check(1).OTHrsPaid;
     PRINT #1, TAB(22); USING "###.##"; Check(1).HolUsed;
     PRINT #1, TAB(29); USING "###.##"; Check(1).CompUsed;
     PRINT #1, TAB(37); USING "####.##"; Check(1).TotRegWage;
     PRINT #1, TAB(64); USING "#####.##"; Check(1).TaxFring;
     PRINT #1, TAB(73); USING "####.##"; Check(1).EICAmt

     PRINT #1, TAB(45); QPTrim$(Check(1).AEarn(3).DCode);
     PRINT #1, TAB(55); QPTrim$(Check(1).AEarn(2).DCode);
     PRINT #1, TAB(64); QPTrim$(Check(1).AEarn(1).DCode)


     'PRINT #1,

     PRINT #1, TAB(2); Check(1).EmpSSN;
     PRINT #1, TAB(15); USING "###.##"; Check(1).VacUsed;
     PRINT #1, TAB(22); USING "###.##"; Check(1).SickUsed;
     PRINT #1, TAB(29); USING "###.##"; Check(1).RegHrsPaid;
     PRINT #1, TAB(37); USING "####.##"; Check(1).TotOTWage;

     PRINT #1, TAB(47); USING "####.##"; Check(1).AEarn(3).DAmt;
     PRINT #1, TAB(56); USING "####.##"; Check(1).AEarn(2).DAmt;
     PRINT #1, TAB(65); USING "####.##"; Check(1).AEarn(1).DAmt;

'     PRINT #1, TAB(56); USING "####.##"; Check(1).TotAdditEarn;


     PRINT #1, TAB(73); USING "####.##"; Check(1).GrossPay
     PRINT #1,
     PRINT #1,

     PRINT #1, TAB(28); LEFT$(Check(1).CDED(1).DCode, 6);
     PRINT #1, TAB(35); LEFT$(Check(1).CDED(2).DCode, 6);
     PRINT #1, TAB(43); LEFT$(Check(1).CDED(3).DCode, 6);
     PRINT #1, TAB(51); LEFT$(Check(1).CDED(4).DCode, 6);
     PRINT #1, TAB(59); LEFT$(Check(1).CDED(5).DCode, 6);
     PRINT #1, TAB(67); LEFT$(Check(1).CDED(6).DCode, 6)

     PRINT #1, TAB(2); USING "####.##"; Round#(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
     PRINT #1, TAB(12); USING "####.##"; Check(1).StaTaxAmt;

     PRINT #1, TAB(26); USING "####.##"; Check(1).CDED(1).DAmt;
     PRINT #1, TAB(33); USING "####.##"; Check(1).CDED(2).DAmt;
     PRINT #1, TAB(42); USING "####.##"; Check(1).CDED(3).DAmt;
     PRINT #1, TAB(50); USING "####.##"; Check(1).CDED(4).DAmt;
     PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(5).DAmt;
     PRINT #1, TAB(66); USING "####.##"; Check(1).CDED(6).DAmt;
     PRINT #1, TAB(73); USING "####.##"; Round#(Check(1).MedTaxAmt + Check(1).SocTaxAmt + Check(1).FedTaxAmt + Check(1).StaTaxAmt);

     PRINT #1, TAB(28); LEFT$(Check(1).CDED(7).DCode, 6);
     PRINT #1, TAB(35); LEFT$(Check(1).CDED(8).DCode, 6);
     PRINT #1, TAB(43); LEFT$(Check(1).CDED(9).DCode, 6);
     PRINT #1, TAB(51); LEFT$(Check(1).CDED(10).DCode, 6);
     PRINT #1, TAB(59); LEFT$(Check(1).CDED(11).DCode, 6);
     PRINT #1, TAB(67); LEFT$(Check(1).CDED(12).DCode, 6)

     PRINT #1, TAB(2); USING "####.##"; Check(1).FedTaxAmt;
     PRINT #1, TAB(12); USING "####.##"; Check(1).RetireAmt;

     PRINT #1, TAB(26); USING "####.##"; Check(1).CDED(7).DAmt;
     PRINT #1, TAB(33); USING "####.##"; Check(1).CDED(8).DAmt;
     PRINT #1, TAB(42); USING "####.##"; Check(1).CDED(9).DAmt;
     PRINT #1, TAB(50); USING "####.##"; Check(1).CDED(10).DAmt;
     PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(11).DAmt;
     PRINT #1, TAB(66); USING "####.##"; Check(1).CDED(12).DAmt;

     PRINT #1, TAB(73); USING "####.##"; Check(1).TotDedAmt

     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(2); USING "######.##"; Check(1).YTDGrossPay;
     PRINT #1, TAB(14); USING "#####.##"; Round(Check(1).YTDSocial + Check(1).YTDMedicare);
     PRINT #1, TAB(23); USING "#####.##"; Check(1).YTDFederal;
     PRINT #1, TAB(33); USING "#####.##"; Check(1).YTDState;
     PRINT #1, TAB(61); Num2Date$(Check(1).CheckDate);
     PRINT #1, TAB(72); USING "#####.##"; Check(1).NetPay

     PRINT #1,
     PRINT #1, TAB(4); "Vac. Due: "; USING "####.###"; Check(1).VactBal;
     PRINT #1, TAB(30); "Sick Due: "; USING "####.###"; Check(1).SickBal;
     PRINT #1, TAB(56); "Comp Due: "; USING "####.###"; Check(1).CompBal

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
     PRINT #1,
     PRINT #1,
     PRINT #1,
     'PRINT #1, "Bottom Line                                                Bottom Line"
'-=-=-=-=-=-=-=-=-=-=-*

