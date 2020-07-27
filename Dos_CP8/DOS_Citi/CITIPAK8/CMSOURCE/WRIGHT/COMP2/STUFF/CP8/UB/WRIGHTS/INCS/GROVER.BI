'Grover, NC 01/30/97
'
    PRINT #1, "~"
    '--Line 3
    PRINT #1, QPTrim$(Check(1).EmpNo) + " " + QPTrim$(Check(1).EmpName);
    PRINT #1, TAB(38); "SSN:"; Check(1).EmpSSN; "  PPE: "; Num2Date$(Check(1).PayEndDate)
    '--Line 4
    PRINT #1,
    '--Line 5
    PRINT #1, "                HRS         EARN           YTD                         PERIOD            YTD"
    '--Line 6
    PRINT #1, TAB(1); USING "HRS WORKED  ####.##"; Check(1).RegHrsWork;
    PRINT #1, TAB(25); USING "#####.##"; Round#(Check(1).RegHrsWork * HRate#);
    PRINT #1, TAB(39); USING "#####.##"; Round(EmpRec3(1).YTDRegPay + Round#(Check(1).RegHrsWork * HRate#));

    PRINT #1, TAB(51); "RETIREMENT";
    PRINT #1, TAB(71); USING "####.##"; Check(1).RetireAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).YTDRetire;
    '--Line 7
    PRINT #1, TAB(1); USING "OT  HRS     ####.##"; Check(1).OTHrsPaid;
    PRINT #1, TAB(25); USING "#####.##"; Check(1).TotOTWage;
    PRINT #1, TAB(39); USING "#####.##"; Round(EmpRec3(1).YTDOTPay + Check(1).TotOTWage);

    PRINT #1, TAB(51); Check(1).CDED(1).DCode;
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(1).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(1).YTDDAmt
    '--Line 8
    PRINT #1, TAB(1); USING "HOL HRS     ####.##"; Check(1).HolUsed;
    PRINT #1, TAB(25); USING "#####.##"; HolPay#;
    PRINT #1, TAB(51); Check(1).CDED(2).DCode;
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(2).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(2).YTDDAmt
    '--Line 9
    PRINT #1, TAB(1); USING "VAC HRS     ####.##"; Check(1).VacUsed;
    PRINT #1, TAB(25); USING "#####.##"; VacPay#;
    PRINT #1, TAB(39); USING "#####.##"; Check(1).VactBal;
    'PRINT #1, TAB(51); Check(1).CDED(3).DCode;
    'PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(3).DAmt;
    'PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(3).YTDDAmt;

    '--Line 10
    PRINT #1, TAB(1); USING "SICK HRS    ####.##"; Check(1).SickUsed;
    PRINT #1, TAB(25); USING "#####.##"; SickPay#;
    PRINT #1, TAB(39); USING "#####.##"; Check(1).SickBal;
   ' PRINT #1, TAB(51); Check(1).CDED(4).DCode;
   ' PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(4).DAmt;
   ' PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(4).YTDDAmt;
    '--Line 11
    PRINT #1, TAB(1); USING "COMP HRS    ####.##"; Check(1).CompUsed;
    PRINT #1, TAB(25); USING "#####.##"; CompPay#;
    PRINT #1, TAB(39); USING "#####.##"; Check(1).CompBal;
    PRINT #1, TAB(51); Check(1).CDED(5).DCode;
    'PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(5).DAmt;
    'PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(5).YTDDAmt;
    '--Line 12
    PRINT #1, TAB(51); Check(1).CDED(6).DCode;
    'PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(6).DAmt;
    'PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(6).YTDDAmt;
    '--Line 13
    PRINT #1, TAB(1); Check(1).AEarn(1).DCode;
    PRINT #1, TAB(25); USING "#####.##"; Check(1).AEarn(1).DAmt;
    'PRINT #1, TAB(51); Check(1).CDED(7).DCode;
    'PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(7).DAmt;
    'PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(7).YTDDAmt
    '--Line 14
    PRINT #1, TAB(1); Check(1).AEarn(2).DCode;
    PRINT #1, TAB(25); USING "#####.##"; Check(1).AEarn(2).DAmt;
    'PRINT #1, TAB(51); Check(1).CDED(8).DCode;
    'PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(8).DAmt;
    'PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(8).YTDDAmt
    '--Line 15
    PRINT #1, TAB(1); Check(1).AEarn(3).DCode;
    PRINT #1, TAB(25); USING "#####.##"; Check(1).AEarn(3).DAmt;
    'PRINT #1, TAB(51); Check(1).CDED(9).DCode;
    'PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(9).DAmt;
    'PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(9).YTDDAmt
    '--Line 16
    PRINT #1, TAB(1); "GROSS PAY";
    PRINT #1, TAB(25); USING "#####.##"; Check(1).GrossPay;
    PRINT #1, TAB(39); USING "#####.##"; Check(1).YTDGrossPay
    'PRINT #1, TAB(51); Check(1).CDED(10).DCode;
    'PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(10).DAmt;
    'PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(10).YTDDAmt
    '--Line 17
    PRINT #1, "FED TAX"; TAB(25); USING "#####.##"; Check(1).FedTaxAmt;
    PRINT #1, TAB(39); USING "#####.##"; Check(1).YTDFederal
    'PRINT #1, TAB(51); Check(1).CDED(11).DCode;
    'PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(11).DAmt;
    'PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(11).YTDDAmt
    '--Line 18
    PRINT #1, "STA TAX"; TAB(25); USING "#####.##"; Check(1).StaTaxAmt;
    PRINT #1, TAB(39); USING "#####.##"; Check(1).YTDState;
    PRINT #1, TAB(51); Check(1).CDED(12).DCode
    'PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(12).DAmt;
    'PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(12).YTDDAmt
    '--Line 19
    PRINT #1, "FICA"; TAB(25); USING "#####.##"; Round(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
    PRINT #1, TAB(39); USING "#####.##"; Check(1).YTDSocial + Check(1).YTDMedicare;
    PRINT #1, TAB(51); "NET PAY";
    PRINT #1, TAB(70); USING "#####.##"; Check(1).NetPay;
    PRINT #1, TAB(85); USING "#####.##"; Check(1).YTDNetPay
    PRINT #1, ""
    PRINT #1, "  RATE       GROSS       TOT DED          NET        YTD GROSS        YTD DED        YTD NET"
    PRINT #1, USING "####.##"; Check(1).BaseRate;
    PRINT #1, TAB(11); USING "#####.##"; Check(1).GrossPay;
    PRINT #1, TAB(25); USING "#####.##"; Check(1).TotDedAmt;
    PRINT #1, TAB(38); USING "#####.##"; Check(1).NetPay;
'    STOP
    PRINT #1, TAB(55); USING "#####.##"; Check(1).YTDGrossPay;
    PRINT #1, TAB(70); USING "#####.##"; Check(1).YTDTotDed;
    PRINT #1, TAB(85); USING "#####.##"; Check(1).YTDNetPay
    PRINT #1,  '<-Line 22 - First line of check
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    '30
    PRINT #1, TAB(10); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##"))
    PRINT #1,
    PRINT #1,
    '33
    PRINT #1, TAB(63); Num2Date$(CheckDate);
    PRINT #1, TAB(75); USING "#######"; TCheckNum&;
    PRINT #1, TAB(85); USING "$$#####,.##"; Check(1).NetPay
    PRINT #1, TAB(10); Check(1).EmpName
    PRINT #1, TAB(10); Check(1).EmpAddr1
    PRINT #1, TAB(10); Check(1).EmpCity; Check(1).EmpState; " "; Check(1).EmpZip
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1, "_" '--Line 42

    'PRINT #1, TAB(1); USING "TOT REG     ####.##"; Check(1).RegHrsPaid;
    'PRINT #1, TAB(25); USING "#####.##"; Check(1).TotRegWage;
    'PRINT #1, TAB(39); USING "#####.##"; Round(EmpRec3(1).YTDRegWage);

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

