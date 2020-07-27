'This is Brookneal check format

'      PRINT #1,
'      PRINT #1, ""
'      PRINT #1, ""
      PRINT #1, Check(1).EmpName
      PRINT #1, "   EMP NO.         SSN        PERIOD END    CHK DATE    BASE RATE"
      PRINT #1, Check(1).EmpNo;
      PRINT #1, TAB(16); Check(1).EmpSSN;
      PRINT #1, TAB(31); Num2Date$(Check(1).PayEndDate); TAB(43); Num2Date$(CheckDate);
      PRINT #1, TAB(56); USING "#####.##"; Check(1).BaseRate
      PRINT #1, "      HRS    Period        YTD    Deductions    Period     YTD  BALANCE"

      PRINT #1, "REG";
      PRINT #1, TAB(6); USING "###.##"; Check(1).RegHrsWork;
      PRINT #1, TAB(14); USING "#####.##"; Round(Check(1).RegHrsWork * Check(1).BaseRate); 'TAB(25); Check(1).YTDGrossPay;
      PRINT #1, TAB(35); Check(1).CDED(1).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(1).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(1).YTDDAmt;
      PRINT #1, "  VAC  "; USING "####.##"; Check(1).VactBal

      PRINT #1, "OT";
      PRINT #1, TAB(6); USING "###.##"; Check(1).OtHrsPaid;
      PRINT #1, TAB(14); USING "#####.##"; Check(1).TotOTWage;
      PRINT #1, TAB(35); Check(1).CDED(2).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(2).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(2).YTDDAmt;
      PRINT #1, "  SICK "; USING "####.##"; Check(1).SickBal

      PRINT #1, "VAC";
      PRINT #1, TAB(6); USING "###.##"; Check(1).VacUsed;
      PRINT #1, TAB(14); USING "#####.##"; VacPay#;
      PRINT #1, TAB(35); Check(1).CDED(1).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(3).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(3).YTDDAmt;
      PRINT #1, "  COMP "; USING "####.##"; Check(1).CompBal

      PRINT #1, "SICK";
      PRINT #1, TAB(6); USING "###.##"; Check(1).SickUsed;
      PRINT #1, TAB(14); USING "#####.##"; SickPay#;
      PRINT #1, TAB(35); Check(1).CDED(4).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(4).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(4).YTDDAmt

      PRINT #1, "HOL";
      PRINT #1, TAB(6); USING "###.##"; Check(1).HolUsed;
      PRINT #1, TAB(14); USING "#####.##"; HolPay#;
      PRINT #1, TAB(35); Check(1).CDED(5).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(5).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(5).YTDDAmt

      PRINT #1, "COMP";
      PRINT #1, TAB(6); USING "###.##"; Check(1).CompUsed;
      PRINT #1, TAB(14); USING "#####.##"; CompPay#;
      PRINT #1, TAB(35); Check(1).CDED(6).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(6).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(6).YTDDAmt

      PRINT #1, "GROSS";
      PRINT #1, TAB(14); USING "#####.##"; Check(1).GrossPay;
      PRINT #1, TAB(25); USING "#####.##"; Check(1).YTDGrossPay;
      PRINT #1, TAB(35); Check(1).CDED(7).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(7).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(7).YTDDAmt

      PRINT #1, "FWT";
      PRINT #1, TAB(14); USING "#####.##"; Check(1).FedTaxAmt;
      PRINT #1, TAB(25); USING "#####.##"; Check(1).YTDFederal;
      PRINT #1, TAB(35); Check(1).CDED(8).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(8).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(8).YTDDAmt

      PRINT #1, "FICA";
      PRINT #1, TAB(14); USING "#####.##"; Round#(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
      PRINT #1, TAB(25); USING "#####.##"; Round#(Check(1).YTDSocial + Check(1).YTDMedicare);
      PRINT #1, TAB(35); Check(1).CDED(9).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(9).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(9).YTDDAmt

      PRINT #1, "RET";
      PRINT #1, TAB(14); USING "#####.##"; Check(1).RetireAmt;
      PRINT #1, TAB(25); USING "#####.##"; Check(1).YTDRetire;
      PRINT #1, TAB(35); Check(1).CDED(10).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(10).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(10).YTDDAmt

      PRINT #1, "NET PAY";
      PRINT #1, TAB(14); USING "#####.##"; Check(1).NetPay;
      PRINT #1, TAB(25); USING "#####.##"; Check(1).YTDNetPay;
      PRINT #1, TAB(35); Check(1).CDED(11).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(11).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(11).YTDDAmt

      PRINT #1, "TOT ADD";
      PRINT #1, TAB(14); USING "#####.##"; Check(1).TotAdditEarn;
      PRINT #1, TAB(35); Check(1).CDED(12).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(12).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(12).YTDDAmt


'----
      PRINT #1, ""
      PRINT #1, ""
      PRINT #1, ""
      PRINT #1, ""

        PRINT #1, ""
        PRINT #1, ""
        PRINT #1, TAB(55); Num2Date$(CheckDate);
        PRINT #1, TAB(72); USING "######"; Check(1).CheckNum
        PRINT #1, ""
      PRINT #1,
      PRINT #1,
      PRINT #1, TAB(10); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##")) 'Print the whole number part
      PRINT #1,
      PRINT #1, TAB(67); USING "$$####,#.##"; Check(1).NetPay
      PRINT #1,
      PRINT #1, ""
      PRINT #1, TAB(12); Check(1).EmpName
      PRINT #1, TAB(12); Check(1).EmpAddr1
      PRINT #1, TAB(12); Check(1).EmpCity; " "; Check(1).EmpState; " "; Check(1).EmpZip
      PRINT #1, ""
      PRINT #1, ""
      PRINT #1, ""
      PRINT #1, ""
      PRINT #1,
      PRINT #1, ""
      PRINT #1, ""
      PRINT #1, Check(1).EmpName
      PRINT #1, "   EMP NO.         SSN        PERIOD END    CHK DATE    BASE RATE"
      PRINT #1, Check(1).EmpNo;
      PRINT #1, TAB(16); Check(1).EmpSSN;
      PRINT #1, TAB(31); Num2Date$(Check(1).PayEndDate); TAB(43); Num2Date$(CheckDate);
      PRINT #1, TAB(56); USING "#####.##"; Check(1).BaseRate
      PRINT #1, "      HRS    Period        YTD    Deductions    Period     YTD  BALANCE"

      PRINT #1, "REG";
      PRINT #1, TAB(6); USING "###.##"; Check(1).RegHrsWork;
      PRINT #1, TAB(14); USING "#####.##"; Round(Check(1).RegHrsWork * Check(1).BaseRate); ' TAB(25); Check(1).YTDGrossPay;
      PRINT #1, TAB(35); Check(1).CDED(1).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(1).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(1).YTDDAmt;
      PRINT #1, "  VAC  "; USING "####.##"; Check(1).VactBal

      PRINT #1, "OT";
      PRINT #1, TAB(6); USING "###.##"; Check(1).OtHrsPaid;
      PRINT #1, TAB(14); USING "#####.##"; Check(1).TotOTWage;
      PRINT #1, TAB(35); Check(1).CDED(2).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(2).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(2).YTDDAmt;
      PRINT #1, "  SICK "; USING "####.##"; Check(1).SickBal

      PRINT #1, "VAC";
      PRINT #1, TAB(6); USING "###.##"; Check(1).VacUsed;
      PRINT #1, TAB(14); USING "#####.##"; VacPay#;
      PRINT #1, TAB(35); Check(1).CDED(1).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(3).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(3).YTDDAmt;
      PRINT #1, "  COMP "; USING "####.##"; Check(1).CompBal

      PRINT #1, "SICK";
      PRINT #1, TAB(6); USING "###.##"; Check(1).SickUsed;
      PRINT #1, TAB(14); USING "#####.##"; SickPay#;
      PRINT #1, TAB(35); Check(1).CDED(4).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(4).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(4).YTDDAmt

      PRINT #1, "HOL";
      PRINT #1, TAB(6); USING "###.##"; Check(1).HolUsed;
      PRINT #1, TAB(14); USING "#####.##"; HolPay#;
      PRINT #1, TAB(35); Check(1).CDED(5).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(5).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(5).YTDDAmt

      PRINT #1, "COMP";
      PRINT #1, TAB(6); USING "###.##"; Check(1).CompUsed;
      PRINT #1, TAB(14); USING "#####.##"; CompPay#;
      PRINT #1, TAB(35); Check(1).CDED(6).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(6).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(6).YTDDAmt

      PRINT #1, "GROSS";
      PRINT #1, TAB(14); USING "#####.##"; Check(1).GrossPay;
      PRINT #1, TAB(25); USING "#####.##"; Check(1).YTDGrossPay;
      PRINT #1, TAB(35); Check(1).CDED(7).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(7).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(7).YTDDAmt

      PRINT #1, "FWT";
      PRINT #1, TAB(14); USING "#####.##"; Check(1).FedTaxAmt;
      PRINT #1, TAB(25); USING "#####.##"; Check(1).YTDFederal;
      PRINT #1, TAB(35); Check(1).CDED(8).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(8).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(8).YTDDAmt

      PRINT #1, "FICA";
      PRINT #1, TAB(14); USING "#####.##"; Round#(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
      PRINT #1, TAB(25); USING "#####.##"; Round#(Check(1).YTDSocial + Check(1).YTDMedicare);
      PRINT #1, TAB(35); Check(1).CDED(9).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(9).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(9).YTDDAmt

      PRINT #1, "RET";
      PRINT #1, TAB(14); USING "#####.##"; Check(1).RetireAmt;
      PRINT #1, TAB(25); USING "#####.##"; Check(1).YTDRetire;
      PRINT #1, TAB(35); Check(1).CDED(10).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(10).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(10).YTDDAmt

      PRINT #1, "NET PAY";
      PRINT #1, TAB(14); USING "#####.##"; Check(1).NetPay;
      PRINT #1, TAB(25); USING "#####.##"; Check(1).YTDNetPay;
      PRINT #1, TAB(35); Check(1).CDED(11).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(11).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(11).YTDDAmt

      PRINT #1, "TOT ADD";
      PRINT #1, TAB(14); USING "#####.##"; Check(1).TotAdditEarn;
      PRINT #1, TAB(35); Check(1).CDED(12).DCode;
      PRINT #1, TAB(48); USING "####.##"; Check(1).CDED(12).DAmt;
      PRINT #1, TAB(56); USING "####.##"; Check(1).CDED(12).YTDDAmt

      PRINT #1, CHR$(12);


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

