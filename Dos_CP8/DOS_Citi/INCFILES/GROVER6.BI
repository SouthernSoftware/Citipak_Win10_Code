'Grover, NC 10/21/99
'#6!!!!!

    PRINT #1, "~"
    PRINT #1,
    PRINT #1, QPTrim$(Check(1).EmpNo) + " " + QPTrim$(Check(1).EmpName);
    PRINT #1, TAB(38); "SSN:"; Check(1).EmpSSN; "   PPE: "; Num2Date$(Check(1).PayEndDate);
    PRINT #1, USING "     RATE:####.##"; Check(1).BaseRate
    PRINT #1,
    PRINT #1, "                HRS         EARN           YTD                         PERIOD            YTD"
    PRINT #1, TAB(1); USING "HRS WORKED  ####.##"; Check(1).RegHrsWork;
    PRINT #1, TAB(25); USING "#####.##"; Round#(Check(1).RegHrsWork * HRate#);
    PRINT #1, TAB(39); USING "#####.##"; Round(EmpRec3(1).YTDRegPay + Round#(Check(1).RegHrsWork * HRate#));
    PRINT #1, TAB(51); "RETIREMENT";
    PRINT #1, TAB(71); USING "####.##"; Check(1).RetireAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).YTDRetire;
    PRINT #1, TAB(1); USING "OT  HRS     ####.##"; Check(1).OTHrsPaid;
    PRINT #1, TAB(25); USING "#####.##"; Check(1).TotOTWage;
    PRINT #1, TAB(39); USING "#####.##"; Round(EmpRec3(1).YTDOTPay + Check(1).TotOTWage);

    PRINT #1, TAB(51); QPTrim$(Check(1).CDED(1).DCode);
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(1).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(1).YTDDAmt
    PRINT #1, TAB(1); USING "HOL  PAY    ####.##"; Check(1).HolUsed;
    PRINT #1, TAB(25); USING "#####.##"; HolPay#;
    PRINT #1, TAB(51); QPTrim$(Check(1).CDED(2).DCode);
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(2).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(2).YTDDAmt
    PRINT #1, TAB(1); USING "VAC  PAY    ####.##"; Check(1).VacUsed;
    PRINT #1, TAB(25); USING "#####.##"; VacPay#;
    PRINT #1, TAB(51); QPTrim$(Check(1).CDED(3).DCode);
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(3).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(3).YTDDAmt
    PRINT #1, TAB(1); USING "SICK PAY    ####.##"; Check(1).SickUsed;
    PRINT #1, TAB(25); USING "#####.##"; SickPay#;
    PRINT #1, TAB(51); QPTrim$(Check(1).CDED(4).DCode);
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(4).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(4).YTDDAmt

    PRINT #1, TAB(1); USING "COMP PAY    ####.##"; Check(1).CompUsed;
    PRINT #1, TAB(25); USING "#####.##"; CompPay#;

    PRINT #1, TAB(51); QPTrim$(Check(1).CDED(5).DCode);
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(5).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(5).YTDDAmt

    'PRINT #1, TAB(51); "TOT DED"; TAB(70); USING "#####.##"; Check(1).TotDedAmt;
    'PRINT #1, TAB(85); USING "#####.##"; Check(1).YTDTotDed

    PRINT #1, TAB(1); QPTrim$(Check(1).AEarn(1).DCode);
    PRINT #1, TAB(25); USING "#####.##"; Check(1).AEarn(1).DAmt;
    PRINT #1, TAB(39); USING "#####.##"; Round(Check(1).AEarn(1).DAmt + Check(1).AEarn(1).YTDDAmt);
    PRINT #1, TAB(51); "FED TAX"; TAB(70); USING "#####.##"; Check(1).FedTaxAmt;
    PRINT #1, TAB(85); USING "#####.##"; Check(1).YTDFederal
    PRINT #1, TAB(1); QPTrim$(Check(1).AEarn(2).DCode);
    PRINT #1, TAB(25); USING "#####.##"; Check(1).AEarn(2).DAmt;
    PRINT #1, TAB(39); USING "#####.##"; Round(Check(1).AEarn(2).DAmt + Check(1).AEarn(2).YTDDAmt);
    PRINT #1, TAB(51); "STA TAX"; TAB(70); USING "#####.##"; Check(1).StaTaxAmt;
    PRINT #1, TAB(85); USING "#####.##"; Check(1).YTDState
    PRINT #1, TAB(1); QPTrim$(Check(1).AEarn(3).DCode);
    PRINT #1, TAB(25); USING "#####.##"; Check(1).AEarn(3).DAmt;
    PRINT #1, TAB(39); USING "#####.##"; Round(Check(1).AEarn(3).DAmt + Check(1).AEarn(3).YTDDAmt);
    PRINT #1, TAB(51); "FICA"; TAB(70); USING "#####.##"; Round(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
    PRINT #1, TAB(85); USING "#####.##"; Check(1).YTDSocial + Check(1).YTDMedicare
    PRINT #1, TAB(1); "GROSS PAY";
    PRINT #1, TAB(25); USING "#####.##"; Check(1).GrossPay;
    PRINT #1, TAB(39); USING "#####.##"; Check(1).YTDGrossPay;
    PRINT #1, TAB(51); "VAC  BAL";
    PRINT #1, TAB(85); USING "#####.##"; Check(1).VactBal

    PRINT #1, TAB(1); "NET PAY";
    PRINT #1, TAB(25); USING "#####.##"; Check(1).NetPay;
    PRINT #1, TAB(39); USING "#####.##"; Check(1).YTDNetPay;
    PRINT #1, TAB(51); "SICK BAL";
    PRINT #1, TAB(85); USING "#####.##"; Check(1).SickBal

    PRINT #1, TAB(1); "TOT DED"; TAB(25); USING "#####.##"; Check(1).TotDedAmt;
    PRINT #1, TAB(39); USING "#####.##"; Check(1).YTDTotDed;
    PRINT #1, TAB(51); "COMP BAL";

    PRINT #1, TAB(85); USING "#####.##"; Check(1).CompBal
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1, "~" '<-Line 22 - First line of check
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1, TAB(4); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##"));
    PRINT #1,
    PRINT #1,
    PRINT #1, TAB(65); Num2Date$(CheckDate);
    PRINT #1, TAB(82); USING "$$#####,.##"; Check(1).NetPay
    PRINT #1, TAB(10); Check(1).EmpName
    PRINT #1, TAB(10); Check(1).EmpAddr1
    PRINT #1, TAB(10); Check(1).EmpCity; Check(1).EmpState; " "; Check(1).EmpZip
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1, "~"  '--Line 39

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

