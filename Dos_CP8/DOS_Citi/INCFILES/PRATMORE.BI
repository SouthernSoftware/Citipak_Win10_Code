'This is "ATMORE" check format 12/16/97

     'This check is in 12-pitch

     PRINT #1, "~"; TAB(80); "~"
     PRINT #1, QPTrim$(Check(1).EmpNo); TAB(10); LEFT$(QPTrim$(Check(1).EmpName), 27);
     PRINT #1, TAB(39); Check(1).EmpSSN;
     PRINT #1, TAB(64); Num2Date$(Check(1).PayEndDate); " "; Num2Date$(CheckDate);
     'PRINT #1, TAB(49); Num2Date$(TransRec(1).PayPdStart); TAB(61); Num2Date$(Check(1).PayEndDate);
     PRINT #1, TAB(88); USING "######"; TCheckNum&
     PRINT #1, ""
     PRINT #1, ""
     PRINT #1, '"                HRS          REG           YTD                            PER            YTD"
     PRINT #1, TAB(1); USING "REG HRS        ####.##"; Check(1).RegHrsWork;
     PRINT #1, TAB(28); USING "#####.##"; Check(1).TotRegWage;
     PRINT #1, TAB(42); USING "#####.##"; Round(EmpRec3(1).YTDRegPay + Check(1).TotRegWage);

     PRINT #1, TAB(56); LEFT$(Check(1).CDED(1).DCode, 19);
     PRINT #1, TAB(76); USING "####.##"; Check(1).CDED(1).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(1).YTDDAmt

     PRINT #1, TAB(1); USING "OT  HRS        ####.##"; Check(1).OTHrsPaid;    'Check(1).TotOTWage;
     PRINT #1, TAB(28); USING "#####.##"; Check(1).TotOTWage;
'fixed
     PRINT #1, TAB(42); USING "#####.##"; Round(EmpRec3(1).YTDOTPay + Check(1).TotOTWage);
'here
     PRINT #1, TAB(56); LEFT$(Check(1).CDED(2).DCode, 19);
     PRINT #1, TAB(76); USING "####.##"; Check(1).CDED(2).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(2).YTDDAmt

     PRINT #1, "ADD EARN"; '    ####.##";
     PRINT #1, TAB(28); USING "#####.##"; Check(1).TotAdditEarn;
     PRINT #1, TAB(42); USING "#####.##"; Round(EmpRec3(1).YTDEarnT + Check(1).TotAdditEarn);
     
     PRINT #1, TAB(56); LEFT$(Check(1).CDED(3).DCode, 19);
     PRINT #1, TAB(76); USING "####.##"; Check(1).CDED(3).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(3).YTDDAmt;


     PRINT #1, TAB(56); LEFT$(Check(1).CDED(4).DCode, 19);
     PRINT #1, TAB(76); USING "####.##"; Check(1).CDED(4).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(4).YTDDAmt;

     PRINT #1, TAB(1); USING "VACT BAL       ####.##"; Check(1).VacUsed;
     PRINT #1, TAB(42); USING "#####.##"; Check(1).VactBal;
     PRINT #1, TAB(56); LEFT$(Check(1).CDED(5).DCode, 19);
     PRINT #1, TAB(76); USING "####.##"; Check(1).CDED(5).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(5).YTDDAmt;

     PRINT #1, TAB(1); USING "SICK BAL       ####.##"; Check(1).SickUsed;
     PRINT #1, TAB(42); USING "#####.##"; Check(1).SickBal;
     PRINT #1, TAB(56); LEFT$(Check(1).CDED(6).DCode, 19);
     PRINT #1, TAB(76); USING "####.##"; Check(1).CDED(6).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(6).YTDDAmt;

     PRINT #1, TAB(1); USING "COMP BAL       ####.##"; Check(1).CompUsed;
     PRINT #1, TAB(42); USING "#####.##"; Check(1).CompBal;
     PRINT #1, TAB(56); LEFT$(Check(1).CDED(7).DCode, 19);
     PRINT #1, TAB(76); USING "####.##"; Check(1).CDED(7).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(7).YTDDAmt

     PRINT #1, TAB(1); USING "HOL HRS        ####.##"; TransRec(1).HOLHOURS;
     PRINT #1, TAB(56); LEFT$(Check(1).CDED(8).DCode, 19);
     PRINT #1, TAB(76); USING "####.##"; Check(1).CDED(8).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(8).YTDDAmt

     PRINT #1, "FED TAX"; TAB(28); USING "#####.##"; Check(1).FedTaxAmt;
     PRINT #1, TAB(42); USING "#####.##"; Check(1).YTDFederal;
     PRINT #1, TAB(56); LEFT$(Check(1).CDED(9).DCode, 19);
     PRINT #1, TAB(76); USING "####.##"; Check(1).CDED(9).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(9).YTDDAmt

     PRINT #1, "STA TAX"; TAB(28); USING "#####.##"; Check(1).StaTaxAmt;
     PRINT #1, TAB(42); USING "#####.##"; Check(1).YTDState;
     PRINT #1, TAB(56); LEFT$(Check(1).CDED(10).DCode, 19);
     PRINT #1, TAB(76); USING "####.##"; Check(1).CDED(10).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(10).YTDDAmt

     PRINT #1, "FICA"; TAB(28); USING "#####.##"; Round(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
     PRINT #1, TAB(42); USING "#####.##"; Check(1).YTDSocial + Check(1).YTDMedicare;
     PRINT #1, TAB(56); LEFT$(Check(1).CDED(11).DCode, 19);
     PRINT #1, TAB(76); USING "####.##"; Check(1).CDED(11).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(11).YTDDAmt


     PRINT #1, "RET "; TAB(28); USING "#####.##"; Check(1).RetireAmt;
     PRINT #1, TAB(42); USING "#####.##"; Check(1).YTDRetire;

     PRINT #1, TAB(56); LEFT$(Check(1).CDED(12).DCode, 19);
     PRINT #1, TAB(76); USING "####.##"; Check(1).CDED(12).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(12).YTDDAmt

     PRINT #1, '"   RATE       GROSS        TOT DED           NET       YTDGROSS     YTD DED          YTD NET"
     PRINT #1, ""
     PRINT #1, USING "######.##"; Check(1).BaseRate;
     PRINT #1, TAB(14); USING "#####.##"; Check(1).GrossPay;
     PRINT #1, TAB(30); USING "####.##"; Round(Check(1).MedTaxAmt + Check(1).SocTaxAmt + Check(1).TotDedAmt + Check(1).StaTaxAmt + Check(1).FedTaxAmt);

     PRINT #1, TAB(43); USING "#####.##"; Check(1).NetPay;
     PRINT #1, TAB(58); USING "#####.##"; Check(1).YTDGrossPay;
     PRINT #1, TAB(72); USING "#####.##"; Round(Check(1).YTDTotDed + Check(1).YTDFederal + Check(1).YTDState + Check(1).YTDSocial + Check(1).YTDMedicare);
   

     'PRINT #1, TAB(68); USING "#####.##"; Check(1).YTDTotDed;
     PRINT #1, TAB(87); USING "#####.##"; Check(1).YTDNetPay

     FOR Cnt2 = 1 TO 3
       PRINT #1, ""
     NEXT
     PRINT #1, 'TAB(83); USING "#######"; TCheckNum&
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(60); Num2Date$(CheckDate); USING "########"; TCheckNum&;
     PRINT #1, TAB(83); USING "$$####,#.##"; Check(1).NetPay
     PRINT #1, TAB(6); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##"))
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(12); Check(1).EmpName
     PRINT #1, TAB(12); Check(1).EmpAddr1
     PRINT #1, TAB(12); Check(1).EmpCity; " "; Check(1).EmpState; " "; Check(1).EmpZip
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, "^"
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

