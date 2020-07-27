'This is the "TROY" check format 1/05/96       print in 10 cpi
     PRINT #1, '"Start"
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(63); "Tax Frng"

     PRINT #1, TAB(1); LTRIM$(Check(1).EMPNO);

     PRINT #1, TAB(8); USING "###.##"; Check(1).RegHrsWork;
'-new
     PRINT #1, TAB(15); USING "###.##"; Check(1).OTHrsPaid;
     PRINT #1, TAB(22); USING "###.##"; Check(1).HolUsed;
     PRINT #1, TAB(29); USING "###.##"; Check(1).CompUsed;
     PRINT #1, TAB(37); USING "####.##"; Check(1).TotRegWage;
'vac pay
'sick pay
     PRINT #1, TAB(47); USING "####.##"; VacPay#;
     PRINT #1, TAB(56); USING "####.##"; SickPay#;
     PRINT #1, TAB(64); USING "####.##"; Check(1).TaxFring;
     PRINT #1, TAB(74); USING "####.##"; Check(1).GrossPay

     PRINT #1,

     PRINT #1, TAB(2); Check(1).EmpSSN;
     PRINT #1, TAB(15); USING "###.##"; Check(1).VacUsed;
     PRINT #1, TAB(22); USING "###.##"; Check(1).SickUsed;
     PRINT #1, TAB(29); USING "###.##"; Check(1).RegHrsPaid;
     PRINT #1, TAB(37); USING "####.##"; Check(1).TotOTWage;

     PRINT #1, TAB(47); USING "####.##"; HolPay#;
     PRINT #1, TAB(56); USING "####.##"; CompPay#;

     'PRINT #1, TAB(47); USING "####.##"; Check(1).AEarn(3).DAmt;
     'PRINT #1, TAB(56); USING "####.##"; Check(1).AEarn(2).DAmt;
     'PRINT #1, TAB(65); USING "####.##"; Check(1).AEarn(1).DAmt;

     PRINT #1, TAB(64); USING "####.##"; Check(1).TotAdditEarn;
     PRINT #1, TAB(74); USING "####.##"; TransRec(1).FedGrossPay
     PRINT #1,
     PRINT #1,

     PRINT #1, TAB(28); LEFT$(Check(1).CDED(1).DCode, 6);
     PRINT #1, TAB(35); LEFT$(Check(1).CDED(2).DCode, 6);
     PRINT #1, TAB(43); LEFT$(Check(1).CDED(3).DCode, 6);
     PRINT #1, TAB(51); LEFT$(Check(1).CDED(4).DCode, 6);
     PRINT #1, TAB(59); LEFT$(Check(1).CDED(5).DCode, 6);
     PRINT #1, TAB(67); LEFT$(Check(1).CDED(6).DCode, 6)

     PRINT #1, TAB(2); USING "####.##"; Round#(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
     PRINT #1, TAB(11); USING "####.##"; Check(1).StaTaxAmt;

     PRINT #1, TAB(27); USING "####.##"; Check(1).CDED(1).DAmt;
     PRINT #1, TAB(34); USING "####.##"; Check(1).CDED(2).DAmt;
     PRINT #1, TAB(42); USING "####.##"; Check(1).CDED(3).DAmt;
     PRINT #1, TAB(50); USING "####.##"; Check(1).CDED(4).DAmt;
     PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(5).DAmt;
     PRINT #1, TAB(66); USING "####.##"; Check(1).CDED(6).DAmt;
     PRINT #1, TAB(74); USING "####.##"; Round#(Check(1).MedTaxAmt + Check(1).SocTaxAmt + Check(1).FedTaxAmt + Check(1).StaTaxAmt);

     PRINT #1, TAB(12); "Retire";
     PRINT #1, TAB(28); LEFT$(Check(1).CDED(7).DCode, 6);
     PRINT #1, TAB(35); LEFT$(Check(1).CDED(8).DCode, 6);
     PRINT #1, TAB(43); LEFT$(Check(1).CDED(9).DCode, 6);
     PRINT #1, TAB(51); LEFT$(Check(1).CDED(10).DCode, 6);
     PRINT #1, TAB(59); LEFT$(Check(1).CDED(11).DCode, 6);
     PRINT #1, TAB(67); LEFT$(Check(1).CDED(12).DCode, 6)

     PRINT #1, TAB(2); USING "####.##"; Check(1).FedTaxAmt;
     PRINT #1, TAB(11); USING "####.##"; Check(1).RetireAmt;

     PRINT #1, TAB(19); USING "####.##"; Round#(Check(1).MedTaxAmt + Check(1).SocTaxAmt + Check(1).FedTaxAmt + Check(1).StaTaxAmt);
'     PRINT #1, TAB(12); USING "####.##"; Check(1).RetireAmt;

     PRINT #1, TAB(27); USING "####.##"; Check(1).CDED(7).DAmt;
     PRINT #1, TAB(34); USING "####.##"; Check(1).CDED(8).DAmt;
     PRINT #1, TAB(42); USING "####.##"; Check(1).CDED(9).DAmt;
     PRINT #1, TAB(50); USING "####.##"; Check(1).CDED(10).DAmt;
     PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(11).DAmt;
     PRINT #1, TAB(66); USING "####.##"; Check(1).CDED(12).DAmt;

     PRINT #1, TAB(74); USING "####.##"; Check(1).TotDedAmt

     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(2); USING "######.##"; Check(1).YTDGrossPay;
     PRINT #1, TAB(14); USING "#####.##"; Round(Check(1).YTDSocial + Check(1).YTDMedicare);
     PRINT #1, TAB(23); USING "#####.##"; Check(1).YTDFederal;
     PRINT #1, TAB(33); USING "#####.##"; Check(1).YTDState;

     PRINT #1, TAB(53); USING "####.##"; Check(1).YTDRetire;

     PRINT #1, TAB(62); Num2Date$(Check(1).PayEndDate);
     PRINT #1, TAB(73); USING "#####.##"; Check(1).NetPay

     PRINT #1,
     PRINT #1, TAB(3); USING "####.###"; Check(1).VactBal;
     PRINT #1, TAB(17); USING "####.###"; Check(1).SickBal;
     PRINT #1, TAB(28); USING "####.###"; Check(1).CompBal

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
     PRINT #1, TAB(5); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##")); 'Print the whole number part
     PRINT #1, TAB(68); USING "$$#####.##"; Check(1).NetPay
     PRINT #1, TAB(56); USING "$$#####.##"; Check(1).NetPay
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(12); Check(1).EmpName
     PRINT #1, TAB(12); Check(1).EmpAddr1
     PRINT #1, TAB(12); Check(1).EmpCity; " "; Check(1).EmpState; " "; Check(1).EmpZip
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, '"End"

'-=-=-=-=-=-=-=-=-=-=-*

