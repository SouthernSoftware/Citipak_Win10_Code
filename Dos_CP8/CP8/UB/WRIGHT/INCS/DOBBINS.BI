'This is the "DOBBINS" check format 8/08/94

'     PRINT #1, CHR$(27) + CHR$(58); 'sets compressed mode   12 dpi
     PRINT #1, ""
     PRINT #1, QPTrim$(Check(1).EmpNo); TAB(10); QPTrim$(Check(1).EmpName);
     PRINT #1, TAB(40); Check(1).EmpSSN;
     PRINT #1, TAB(63); Num2Date$(Check(1).PayEndDate); TAB(74); Num2Date$(CheckDate);
     PRINT #1, TAB(86); USING "######"; TCheckNum&
     PRINT #1, ""
     PRINT #1, ""
     PRINT #1, TAB(3); USING "REG HRS        ####.##"; Check(1).RegHrsPaid;
     PRINT #1, TAB(55); Check(1).CDED(1).DCode;
     PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(1).DAmt;
     PRINT #1, TAB(83); USING "####.##"; Check(1).CDED(1).YTDDAmt

     PRINT #1, TAB(3); USING "OT  HRS        ####.##"; Check(1).OTHrsPaid;    'Check(1).TotOTWage;
     PRINT #1, TAB(55); Check(1).CDED(2).DCode;
     PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(2).DAmt;
     PRINT #1, TAB(83); USING "####.##"; Check(1).CDED(2).YTDDAmt;

     PRINT #1, TAB(3); USING "ALT ERN        ####.##"; Check(1).TotAdditEarn;
     PRINT #1, TAB(55); Check(1).CDED(3).DCode;
     PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(3).DAmt;
     PRINT #1, TAB(83); USING "####.##"; Check(1).CDED(3).YTDDAmt;


     PRINT #1, TAB(55); Check(1).CDED(4).DCode;
     PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(4).DAmt;
     PRINT #1, TAB(83); USING "####.##"; Check(1).CDED(4).YTDDAmt;

     PRINT #1, TAB(3); USING "VACT BAL       ####.##"; Check(1).VactBal;
     PRINT #1, TAB(55); Check(1).CDED(5).DCode;
     PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(5).DAmt;
     PRINT #1, TAB(83); USING "####.##"; Check(1).CDED(5).YTDDAmt;

     PRINT #1, TAB(3); USING "SICK BAL       ####.##"; Check(1).SickBal;
     PRINT #1, TAB(55); Check(1).CDED(6).DCode;
     PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(6).DAmt;
     PRINT #1, TAB(83); USING "####.##"; Check(1).CDED(6).YTDDAmt;

     PRINT #1, TAB(3); USING "COMP BAL       ####.##"; Check(1).CompBal;
     PRINT #1, TAB(55); Check(1).CDED(7).DCode;
     PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(7).DAmt;
     PRINT #1, TAB(83); USING "####.##"; Check(1).CDED(7).YTDDAmt

     PRINT #1, TAB(55); Check(1).CDED(8).DCode;
     PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(8).DAmt;
     PRINT #1, TAB(83); USING "####.##"; Check(1).CDED(8).YTDDAmt

     PRINT #1, TAB(3); USING " FED TAX       ####.##"; Check(1).FedTaxAmt;
     PRINT #1, TAB(42); USING "#####.##"; Check(1).YTDFederal;
     PRINT #1, TAB(55); Check(1).CDED(9).DCode;
     PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(9).DAmt;
     PRINT #1, TAB(83); USING "####.##"; Check(1).CDED(9).YTDDAmt

     PRINT #1, TAB(3); USING " STA TAX       ####.##"; Check(1).StaTaxAmt;
     PRINT #1, TAB(42); USING "#####.##"; Check(1).YTDState;
     PRINT #1, TAB(55); Check(1).CDED(10).DCode;
     PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(10).DAmt;
     PRINT #1, TAB(83); USING "####.##"; Check(1).CDED(10).YTDDAmt

     PRINT #1, TAB(3); USING "    FICA       ####.##"; Check(1).MedTaxAmt + Check(1).SocTaxAmt;
     PRINT #1, TAB(42); USING "#####.##"; Check(1).YTDSocial + Check(1).YTDMedicare;
     PRINT #1, TAB(55); Check(1).CDED(11).DCode;
     PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(11).DAmt;
     PRINT #1, TAB(83); USING "####.##"; Check(1).CDED(11).YTDDAmt

     PRINT #1, TAB(55); Check(1).CDED(12).DCode;
     PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(12).DAmt;
     PRINT #1, TAB(83); USING "####.##"; Check(1).CDED(12).YTDDAmt
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(2); USING "#####.##"; Check(1).BaseRate;
     PRINT #1, TAB(15); USING "#####.##"; Check(1).GrossPay;

     PRINT #1, TAB(26); USING "####.##"; Check(1).TotDedAmt;
     PRINT #1, TAB(41); USING "#####.##"; Check(1).NetPay;
     PRINT #1, TAB(56); USING "#####.##"; Check(1).YTDGrossPay;
     PRINT #1, TAB(72); USING "#####.##"; Check(1).YTDTotDed;
     PRINT #1, TAB(83); USING "#####.##"; Check(1).YTDNetPay

     FOR Cnt2 = 1 TO 9
       PRINT #1, ""
     NEXT

     PRINT #1, TAB(57); Num2Date$(CheckDate);
     PRINT #1, TAB(69); USING "#####"; TCheckNum&;
     PRINT #1, TAB(79); USING "$$####,#.##"; Check(1).NetPay
     PRINT #1,
     PRINT #1, TAB(3); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##")) 'Print the whole number part
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(10); Check(1).EmpName
     PRINT #1, TAB(10); Check(1).EmpAddr1
     PRINT #1, TAB(10); Check(1).EmpCity; " "; Check(1).EmpState; " "; Check(1).EmpZip
     PRINT #1, ""
     PRINT #1, ""
     PRINT #1, ""
     PRINT #1, ""
     PRINT #1, ""
'     PRINT #1, ""
'     PRINT #1, ""
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

