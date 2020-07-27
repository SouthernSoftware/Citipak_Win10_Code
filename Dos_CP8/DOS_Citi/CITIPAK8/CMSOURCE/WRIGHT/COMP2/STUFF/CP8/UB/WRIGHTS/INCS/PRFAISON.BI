'Faison 42 line check.

    '--Line 1
    PRINT #1, "~"
    PRINT #1, TAB(9); QPTrim$(Check(1).EmpNo); TAB(24); QPTrim$(Check(1).EmpName)
    PRINT #1,
    PRINT #1, '"                HRS         EARN           YTD                         PERIOD            YTD"
    PRINT #1, TAB(3); USING "HRS WORKED ####.##"; Check(1).RegHrsWork;
    PRINT #1, TAB(22); USING "#####.##"; Check(1).TotRegWage;
    PRINT #1, TAB(43); "RETIREMENT";
    PRINT #1, TAB(58); USING "####.##"; Check(1).RetireAmt;
    PRINT #1, TAB(70); USING "####.##"; Check(1).YTDRetire;

    '--Line 5
    PRINT #1, TAB(3); USING "HOL HRS    ####.##"; Check(1).HolUsed;
    PRINT #1, TAB(22); USING "#####.##"; HolPay#;
    PRINT #1, TAB(43); QPTrim$(Check(1).CDED(1).DCode);
    PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(1).DAmt;
    PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(1).YTDDAmt;
    PRINT #1, TAB(85); USING "#######"; TCheckNum&

    '--Line 6
    PRINT #1, TAB(3); USING "SICK HRS   ####.##"; Check(1).SickUsed;
    PRINT #1, TAB(22); USING "#####.##"; SickPay#;
    PRINT #1, TAB(31); USING "#####.##"; Check(1).SickBal;
    PRINT #1, TAB(43); QPTrim$(Check(1).CDED(2).DCode);
    PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(2).DAmt;
    PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(2).YTDDAmt

    '--Line 7
    PRINT #1, TAB(3); USING "COMP HRS   ####.##"; Check(1).CompUsed;
    PRINT #1, TAB(22); USING "#####.##"; CompPay#;
    PRINT #1, TAB(31); USING "#####.##"; Check(1).CompBal;
    PRINT #1, TAB(43); QPTrim$(Check(1).CDED(3).DCode);
    PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(3).DAmt;
    PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(3).YTDDAmt;

    '--Line 8
    PRINT #1, TAB(3); USING "VAC HRS    ####.##"; Check(1).VacUsed;
    PRINT #1, TAB(22); USING "#####.##"; VacPay#;
    PRINT #1, TAB(31); USING "#####.##"; Check(1).VactBal;
    PRINT #1, TAB(43); QPTrim$(Check(1).CDED(4).DCode);
    PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(4).DAmt;
    PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(4).YTDDAmt;

    '--Line 9
    PRINT #1, TAB(3); USING "TOT REG HRS####.##"; Check(1).RegHrsPaid;
    PRINT #1, TAB(22); USING "#####.##"; Check(1).TotRegWage;
    'PRINT #1, tab(31); USING "#####.##"; Round(EmpRec3(1).YTDRegWage);
    PRINT #1, TAB(43); QPTrim$(Check(1).CDED(5).DCode);
    PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(5).DAmt;
    PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(5).YTDDAmt;

    '--Line 10
    PRINT #1, TAB(3); USING "OT  HRS    ####.##"; Check(1).OTHrsPaid;
    PRINT #1, TAB(22); USING "#####.##"; Check(1).TotOTWage;
    'PRINT #1, tab(31); USING "#####.##"; Round(EmpRec3(1).YTDOTPay + Check(1).TotOTWage);
    PRINT #1, TAB(43); QPTrim$(Check(1).CDED(6).DCode);
    PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(6).DAmt;
    PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(6).YTDDAmt;
    PRINT #1, TAB(83); Num2Date$(CheckDate)

    '--Line 11
    TEarn# = Round#(Check(1).AEarn(1).DAmt + Check(1).AEarn(2).DAmt + Check(1).AEarn(3).DAmt)
    PRINT #1, TAB(3); "OTH EARN";
    PRINT #1, TAB(22); USING "#####.##"; TEarn#;
    PRINT #1, TAB(43); QPTrim$(Check(1).CDED(7).DCode);
    PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(7).DAmt;
    PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(7).YTDDAmt

    PRINT #1, TAB(3); "GROSS PAY";
    PRINT #1, TAB(22); USING "#####.##"; Check(1).GrossPay;
    PRINT #1, TAB(31); USING "#####.##"; Check(1).YTDGrossPay;
    PRINT #1, TAB(43); QPTrim$(Check(1).CDED(8).DCode);
    PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(8).DAmt;
    PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(8).YTDDAmt

    PRINT #1, TAB(3); "FED TAX"; TAB(22); USING "#####.##"; Check(1).FedTaxAmt;
    PRINT #1, TAB(31); USING "#####.##"; Check(1).YTDFederal;
    PRINT #1, TAB(43); QPTrim$(Check(1).CDED(9).DCode);
    PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(9).DAmt;
    PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(9).YTDDAmt

    '--Line 14
    PRINT #1, TAB(3); "STA TAX"; TAB(22); USING "#####.##"; Check(1).StaTaxAmt;
    PRINT #1, TAB(31); USING "#####.##"; Check(1).YTDState;
    PRINT #1, TAB(43); QPTrim$(Check(1).CDED(10).DCode);
    PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(10).DAmt;
    PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(10).YTDDAmt

   '--Line 15
    PRINT #1, TAB(3); "FICA"; TAB(22); USING "#####.##"; Round(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
    PRINT #1, TAB(31); USING "#####.##"; Check(1).YTDSocial + Check(1).YTDMedicare;
    PRINT #1, TAB(43); QPTrim$(Check(1).CDED(11).DCode);
    PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(11).DAmt;
    PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(11).YTDDAmt

    '--Line 16
    'PRINT #1, TAB(43); QPTrim$(Check(1).CDED(12).DCode);
    'PRINT #1, TAB(58); USING "####.##"; Check(1).CDED(12).DAmt;
    'PRINT #1, TAB(70); USING "####.##"; Check(1).CDED(12).YTDDAmt

    '--Line 17
    'PRINT #1, TAB(43); "NET PAY";
    PRINT #1, TAB(85); USING "#####.##"; Check(1).NetPay
    PRINT #1,
    'PRINT #1, TAB(85); USING "#####.##"; Check(1).YTDNetPay

    '--Line 18 - Last line of stub
    PRINT #1, ""

    'PRINT #1, "2"
    PRINT #1, TAB(3); USING "#####.##"; Check(1).YTDGrossPay; TAB(31); Check(1).YTDFederal; TAB(43); Check(1).YTDSocial + Check(1).YTDMedicare;
    PRINT #1, USING "#####.##"; TAB(58); Check(1).YTDState

    PRINT #1, ""
    PRINT #1, ""

    '--Line 19 - First line of check
      PRINT #1, ""
      PRINT #1, ""
      PRINT #1,
      PRINT #1,
      PRINT #1,
      PRINT #1, TAB(63); QPTrim$(Check(1).EmpNo); TAB(80); USING "#######"; TCheckNum&
      PRINT #1,
      PRINT #1,
      PRINT #1,
      PRINT #1, TAB(3); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##")) '--Line 30
      PRINT #1, TAB(63); Num2Date$(CheckDate); TAB(77); USING "$$#####,.##"; Check(1).NetPay
      PRINT #1, TAB(10); Check(1).EmpName
      PRINT #1, TAB(10); Check(1).EmpAddr1
      PRINT #1, TAB(10); Check(1).EmpCity; " "; Check(1).EmpState; " "; Check(1).EmpZip
      PRINT #1,
      PRINT #1,
      PRINT #1,
      PRINT #1,
      PRINT #1,
      PRINT #1, "~" '--Line 42
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

