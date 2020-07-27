'--This is the "old standard" payroll check. Product 9007

     PRINT #1, "~"
     PRINT #1,

     '--Line 3
     PRINT #1, Check(1).EmpName;
     PRINT #1, TAB(65); USING "Rate: $$####.##"; Check(1).BaseRate

     PRINT #1,

     '--Line 5 Desc Only
     'PRINT #1,
     PRINT #1, TAB(65); "Other 1";
     PRINT #1, TAB(74); "Other 2"
     
     '--Line 6 Hours Section
     PRINT #1, TAB(2); LEFT$(QPTrim$(Check(1).EMPNO), 5);
     PRINT #1, TAB(8); USING "###.##"; Check(1).RegHrsWork;
     PRINT #1, TAB(15); USING "###.##"; Check(1).OTHrsPaid;
     PRINT #1, TAB(22); USING "###.##"; Check(1).HolUsed;
     PRINT #1, TAB(29); USING "###.##"; Check(1).CompUsed;
     '--Earnings section
     PRINT #1, TAB(38); USING "####.##"; WorkPay#;
     PRINT #1, TAB(47); USING "####.##"; VacPay#;
     PRINT #1, TAB(56); USING "####.##"; SickPay#;
     PRINT #1, TAB(65); USING "####.##"; Check(1).AEarn(1).DAmt;
     PRINT #1, TAB(74); USING "####.##"; Check(1).AEarn(2).DAmt

     '--Line 7
     PRINT #1, TAB(65); "Other 3"

     '--Line 8 Hours
     PRINT #1, TAB(2); Check(1).EmpSSN;
     PRINT #1, TAB(15); USING "###.##"; Check(1).VacUsed;
     PRINT #1, TAB(22); USING "###.##"; Check(1).SickUsed;
     PRINT #1, TAB(29); USING "###.##"; AllHrsPaid#;

     '--Line 8 Earnings
     PRINT #1, TAB(38); USING "####.##"; Check(1).TotOTWage;
     PRINT #1, TAB(47); USING "####.##"; HolPay#;
     PRINT #1, TAB(56); USING "####.##"; CompPay#;
     PRINT #1, TAB(65); USING "####.##"; Check(1).AEarn(3).DAmt;
     PRINT #1, TAB(74); USING "####.##"; TransRec(1).GrossPay

     '--Line 9
     PRINT #1, TAB(74); "Adv EIC"
     PRINT #1, TAB(74); USING "####.##"; Check(1).EICAmt

     '--Line 12
     PRINT #1, TAB(28); LEFT$(Check(1).CDED(1).DCode, 6);
     PRINT #1, TAB(35); LEFT$(Check(1).CDED(2).DCode, 6);
     PRINT #1, TAB(43); LEFT$(Check(1).CDED(3).DCode, 6);
     PRINT #1, TAB(51); LEFT$(Check(1).CDED(4).DCode, 6);
     PRINT #1, TAB(59); LEFT$(Check(1).CDED(5).DCode, 6);
     PRINT #1, TAB(67); LEFT$(Check(1).CDED(6).DCode, 6)

     '--Line 13
     PRINT #1, TAB(2); USING "####.##"; Round#(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
     PRINT #1, TAB(11); USING "####.##"; Check(1).StaTaxAmt;
     PRINT #1, TAB(27); USING "####.##"; Check(1).CDED(1).DAmt;
     PRINT #1, TAB(34); USING "####.##"; Check(1).CDED(2).DAmt;
     PRINT #1, TAB(42); USING "####.##"; Check(1).CDED(3).DAmt;
     PRINT #1, TAB(50); USING "####.##"; Check(1).CDED(4).DAmt;
     PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(5).DAmt;
     PRINT #1, TAB(66); USING "####.##"; Check(1).CDED(6).DAmt;
     PRINT #1, TAB(74); USING "####.##"; Round#(Check(1).MedTaxAmt + Check(1).SocTaxAmt + Check(1).StaTaxAmt + Check(1).FedTaxAmt)'TotTaxes#

     '--Line 14
     PRINT #1, TAB(12); "Retire";
     PRINT #1, TAB(28); LEFT$(Check(1).CDED(7).DCode, 6);
     PRINT #1, TAB(35); LEFT$(Check(1).CDED(8).DCode, 6);
     PRINT #1, TAB(43); LEFT$(Check(1).CDED(9).DCode, 6);
     PRINT #1, TAB(51); LEFT$(Check(1).CDED(10).DCode, 6);
     PRINT #1, TAB(59); LEFT$(Check(1).CDED(11).DCode, 6);
     PRINT #1, TAB(67); LEFT$(Check(1).CDED(12).DCode, 6)

     '--Line 15
     PRINT #1, TAB(2); USING "####.##"; Check(1).FedTaxAmt;
     PRINT #1, TAB(11); USING "####.##"; Check(1).RetireAmt;
     PRINT #1, TAB(27); USING "####.##"; Check(1).CDED(7).DAmt;
     PRINT #1, TAB(34); USING "####.##"; Check(1).CDED(8).DAmt;
     PRINT #1, TAB(42); USING "####.##"; Check(1).CDED(9).DAmt;
     PRINT #1, TAB(50); USING "####.##"; Check(1).CDED(10).DAmt;
     PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(11).DAmt;
     PRINT #1, TAB(66); USING "####.##"; Check(1).CDED(12).DAmt;
     PRINT #1, TAB(74); USING "####.##"; Check(1).TotDedAmt   'TotVolDed#

     PRINT #1,
     PRINT #1,
     PRINT #1,

     '--Line 18
     PRINT #1, TAB(2); USING "######.##"; Check(1).YTDGrossPay;
     PRINT #1, TAB(14); USING "#####.##"; Round#(Check(1).YTDSocial + Check(1).YTDMedicare);
     PRINT #1, TAB(23); USING "#####.##"; Check(1).YTDFederal;
     PRINT #1, TAB(33); USING "#####.##"; Check(1).YTDState;
     PRINT #1, TAB(43); USING "####.##"; Check(1).YTDRetire;
     PRINT #1, TAB(62); Num2Date$(Check(1).PayEndDate);
     PRINT #1, TAB(73); USING "#####.##"; Check(1).NetPay

     PRINT #1,

     PRINT #1, TAB(3); USING "Unused Vacation ####.###"; Check(1).VactBal;
     PRINT #1, TAB(34); USING "Unused Sick ####.###"; Check(1).SickBal;
     PRINT #1, TAB(61); USING "Unused Comp ####.###"; Check(1).CompBal

     '--Line 21 - Last line of stub"
     PRINT #1, TAB(3); USING "Other Taxable ######,.##"; Check(1).TaxFring
     
     IF Check(1).DDFlag = True THEN
       PRINT #1,
       PRINT #1, "   DIRECT DEPOSIT VOUCHER"

       PRINT #1, TAB(10); "VOID            VOID             VOID            VOID            VOID"
       PRINT #1, TAB(10); "VOID            VOID             VOID            VOID            VOID"
       PRINT #1, TAB(10); "VOID            VOID             VOID            VOID            VOID"
       PRINT #1, TAB(10); "VOID            VOID             VOID            VOID            VOID"
       PRINT #1, TAB(10); "VOID            VOID             VOID            VOID            VOID"
       PRINT #1, TAB(10); "VOID            VOID             VOID            VOID            VOID"
       PRINT #1, TAB(10); "VOID            VOID             VOID            VOID            VOID"
       PRINT #1, TAB(10); "VOID            VOID             VOID            VOID            VOID"
       PRINT #1, TAB(10); "VOID            VOID             VOID            VOID            VOID"
       PRINT #1,
       PRINT #1, TAB(12); Check(1).EmpName
       PRINT #1, TAB(12); Check(1).EmpAddr1
       PRINT #1, TAB(12); QPTrim$(Check(1).EmpCity); ", "; Check(1).EmpState; " "; Check(1).EmpZip
       PRINT #1,
       PRINT #1, TAB(10); "VOID            VOID             VOID            VOID            VOID"
       PRINT #1, TAB(10); "VOID            VOID             VOID            VOID            VOID"
       PRINT #1, TAB(10); "VOID            VOID             VOID            VOID            VOID"
       PRINT #1, TAB(10); "VOID            VOID             VOID            VOID            VOID"
       PRINT #1, "_"

     ELSE
       PRINT #1, '--Line 22 - Top of check
       PRINT #1,
       PRINT #1,
       PRINT #1, TAB(62); Num2Date$(Check(1).CheckDate)
       PRINT #1, 'TAB(73); Check(1).CheckNum
       PRINT #1,
       PRINT #1,
       PRINT #1,
       PRINT #1, TAB(65); USING "$$#####.##"; Check(1).NetPay
       PRINT #1, TAB(5); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##"))
       PRINT #1,
       PRINT #1,
       PRINT #1, TAB(12); Check(1).EmpName
       PRINT #1, TAB(12); Check(1).EmpAddr1
       PRINT #1, TAB(12); QPTrim$(Check(1).EmpCity); ", "; Check(1).EmpState; " "; Check(1).EmpZip
       PRINT #1,
       PRINT #1,
       PRINT #1,
       PRINT #1,
       PRINT #1,
       PRINT #1, "_"
     END IF

'-=-=-=-=-=-=-=-=-=-=-*

