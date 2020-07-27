'This is the "TROUTMAN" check format 07/25/96
'uses 12 dpi

     PRINT #1, ' "TOP"; TAB(82); "TOP"
     PRINT #1, QPTrim$(Check(1).EmpNo); TAB(11); QPTrim$(Check(1).EmpName);
     PRINT #1, TAB(40); Check(1).EmpSSN;

     EndDate$ = Num2Date$(Check(1).PayEndDate)
     ChkDate$ = Num2Date$(Check(1).CheckDate)
     EndYear$ = RIGHT$(EndDate$, 2)
     ChkYear$ = RIGHT$(ChkDate$, 2)
     EndDate$ = LEFT$(EndDate$, 6) + EndYear$
     ChkDate$ = LEFT$(ChkDate$, 6) + ChkYear$

     PRINT #1, TAB(65); EndDate$; TAB(77); ChkDate$;
     PRINT #1, TAB(87); USING "######"; TCheckNum&
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(1); USING "REG HRS          ####.##"; Check(1).RegHrsWork;
     PRINT #1, TAB(30); USING "#####.##"; Check(1).TotRegWage;
     PRINT #1, TAB(43); USING "#####.##"; Round(EmpRec3(1).YTDRegPay + Check(1).TotRegWage);

     PRINT #1, TAB(55); Check(1).CDED(1).DCode;
     PRINT #1, TAB(75); USING "####.##"; Check(1).CDED(1).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(1).YTDDAmt

     PRINT #1, TAB(1); USING "OT  HRS          ####.##"; Check(1).OTHrsPaid;    'Check(1).TotOTWage;
     PRINT #1, TAB(30); USING "#####.##"; Check(1).TotOTWage;
     PRINT #1, TAB(43); USING "#####.##"; Round(EmpRec3(1).YTDOTPay + Check(1).TotOTWage);

     PRINT #1, TAB(55); Check(1).CDED(2).DCode;
     PRINT #1, TAB(75); USING "####.##"; Check(1).CDED(2).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(2).YTDDAmt

     PRINT #1, "ADD EARN"; '    ####.##";
     PRINT #1, TAB(30); USING "#####.##"; Check(1).TotAdditEarn;
     PRINT #1, TAB(43); USING "#####.##"; Round(EmpRec3(1).YTDEarnT + Check(1).TotAdditEarn);
     
     PRINT #1, TAB(55); Check(1).CDED(3).DCode;
     PRINT #1, TAB(75); USING "####.##"; Check(1).CDED(3).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(3).YTDDAmt;
     
     PRINT #1, TAB(55); Check(1).CDED(4).DCode;
     PRINT #1, TAB(75); USING "####.##"; Check(1).CDED(4).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(4).YTDDAmt;

     PRINT #1, TAB(1); USING "VACT BAL         ####.##"; Check(1).VacUsed;
     PRINT #1, TAB(43); USING "#####.##"; Check(1).VactBal;
     PRINT #1, TAB(55); Check(1).CDED(5).DCode;
     PRINT #1, TAB(75); USING "####.##"; Check(1).CDED(5).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(5).YTDDAmt;

     PRINT #1, TAB(1); USING "SICK BAL         ####.##"; Check(1).SickUsed;
     PRINT #1, TAB(43); USING "#####.##"; Check(1).SickBal;
     PRINT #1, TAB(55); Check(1).CDED(6).DCode;
     PRINT #1, TAB(75); USING "####.##"; Check(1).CDED(6).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(6).YTDDAmt;

     PRINT #1, TAB(1); USING "COMP BAL         ####.##"; Check(1).CompUsed;
     PRINT #1, TAB(43); USING "#####.##"; Check(1).CompBal;
     PRINT #1, TAB(55); Check(1).CDED(7).DCode;
     PRINT #1, TAB(75); USING "####.##"; Check(1).CDED(7).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(7).YTDDAmt

     PRINT #1, TAB(55); Check(1).CDED(8).DCode;
     PRINT #1, TAB(75); USING "####.##"; Check(1).CDED(8).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(8).YTDDAmt

     PRINT #1, "FED TAX"; TAB(30); USING "#####.##"; Check(1).FedTaxAmt;
     PRINT #1, TAB(43); USING "#####.##"; Check(1).YTDFederal;
     PRINT #1, TAB(55); Check(1).CDED(9).DCode;
     PRINT #1, TAB(75); USING "####.##"; Check(1).CDED(9).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(9).YTDDAmt

     PRINT #1, "STA TAX"; TAB(30); USING "#####.##"; Check(1).StaTaxAmt;
     PRINT #1, TAB(43); USING "#####.##"; Check(1).YTDState;
     PRINT #1, TAB(55); Check(1).CDED(10).DCode;
     PRINT #1, TAB(75); USING "####.##"; Check(1).CDED(10).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(10).YTDDAmt

     PRINT #1, "FICA"; TAB(30); USING "#####.##"; Round(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
     PRINT #1, TAB(43); USING "#####.##"; Check(1).YTDSocial + Check(1).YTDMedicare;
     PRINT #1, TAB(55); Check(1).CDED(11).DCode;
     PRINT #1, TAB(75); USING "####.##"; Check(1).CDED(11).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(11).YTDDAmt

     PRINT #1, "RET "; TAB(30); USING "#####.##"; Check(1).RetireAmt;
     PRINT #1, TAB(43); USING "#####.##"; Check(1).YTDRetire;
     
     PRINT #1, TAB(55); Check(1).CDED(12).DCode;
     PRINT #1, TAB(75); USING "####.##"; Check(1).CDED(12).DAmt;
     PRINT #1, TAB(88); USING "####.##"; Check(1).CDED(12).YTDDAmt

     PRINT #1,
     PRINT #1,
     PRINT #1,

     PRINT #1, USING "#####.##"; Check(1).BaseRate;
     PRINT #1, TAB(15); USING "#####.##"; Check(1).GrossPay;
     PRINT #1, TAB(31); USING "####.##"; Round(Check(1).MedTaxAmt + Check(1).SocTaxAmt + Check(1).TotDedAmt + Check(1).StaTaxAmt + Check(1).FedTaxAmt);

     PRINT #1, TAB(46); USING "#####.##"; Check(1).NetPay;
     PRINT #1, TAB(59); USING "#####.##"; Check(1).YTDGrossPay;
     PRINT #1, TAB(73); USING "#####.##"; Round(Check(1).YTDTotDed + Check(1).YTDFederal + Check(1).YTDState + Check(1).YTDSocial + Check(1).YTDMedicare);
     PRINT #1, TAB(88); USING "#####.##"; Check(1).YTDNetPay

     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
               
     PRINT #1, TAB(60); Num2Date$(CheckDate);
     PRINT #1, TAB(73); USING "#######"; TCheckNum&;
     PRINT #1, TAB(84); USING "$$####,#.##"; Check(1).NetPay
     PRINT #1,
     PRINT #1, TAB(8); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##"))'Print the whole number part
     PRINT #1,
     PRINT #1,

     PRINT #1, TAB(13); Check(1).EmpName
     PRINT #1, TAB(13); Check(1).EmpAddr1
     PRINT #1, TAB(13); Check(1).EmpCity; " "; Check(1).EmpState; " "; Check(1).EmpZip
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, ' "bottom"; TAB(82); "bottom"
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

