'Wrightsville Beech 02/24/98
'Blank top stub format  39 lines

    '--Line 1
    PRINT #1, "~"

    IF Check(1).DDFlag = True THEN
      PRINT #1, " DIRECT DEPOSIT VOUCHER";
    END IF

    PRINT #1, TAB(30); QPTrim$(Check(1).EmpNo) + " " + QPTrim$(Check(1).EmpName);
    PRINT #1, TAB(58); "PPE: "; Num2Date$(Check(1).PayEndDate);
    PRINT #1, TAB(75); USING "Rate: ####.##"; Check(1).BaseRate
    '--Line 3

    PRINT #1, "                HRS         EARN           YTD                         PERIOD            YTD"
    '--Line 4
    PRINT #1, TAB(1); USING "HRS WORKED  ####.##"; Check(1).RegHrsWork;
    PRINT #1, TAB(25); USING "#####.##"; WorkPay#;
    PRINT #1, TAB(51); "RETIREMENT";
    PRINT #1, TAB(71); USING "####.##"; Check(1).RetireAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).YTDRetire;

    '--Line 5
    PRINT #1, TAB(1); USING "HOL HRS     ####.##"; Check(1).HolUsed;
    PRINT #1, TAB(25); USING "#####.##"; HolPay#;
    PRINT #1, TAB(39); USING "#####.##"; HolBal#;
    PRINT #1, TAB(51); Check(1).CDED(1).DCode;
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(1).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(1).YTDDAmt

    '--Line 6
    PRINT #1, TAB(1); USING "SICK HRS    ####.##"; Check(1).SickUsed;
    PRINT #1, TAB(25); USING "#####.##"; SickPay#;
    PRINT #1, TAB(39); USING "#####.##"; Check(1).SickBal;
    PRINT #1, TAB(51); Check(1).CDED(2).DCode;
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(2).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(2).YTDDAmt

    '--Line 7
    PRINT #1, TAB(1); USING "COMP HRS    ####.##"; Check(1).CompUsed;
    PRINT #1, TAB(25); USING "#####.##"; CompPay#;
    PRINT #1, TAB(39); USING "#####.##"; Check(1).CompBal;
    PRINT #1, TAB(51); Check(1).CDED(3).DCode;
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(3).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(3).YTDDAmt;

    '--Line 8
    PRINT #1, TAB(1); USING "VAC HRS     ####.##"; Check(1).VacUsed;
    PRINT #1, TAB(25); USING "#####.##"; VacPay#;
    PRINT #1, TAB(39); USING "#####.##"; Check(1).VactBal;
    PRINT #1, TAB(51); Check(1).CDED(4).DCode;
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(4).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(4).YTDDAmt;

    '--Line 9
    PRINT #1, TAB(1); USING "PERS HRS    ####.##"; Check(1).PerUsed;
    PRINT #1, TAB(25); USING "#####.##"; PerPay#;
    PRINT #1, TAB(39); USING "#####.##"; PerBal#;
    PRINT #1, TAB(51); Check(1).CDED(5).DCode;
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(5).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(5).YTDDAmt;

    '--Line 10
    PRINT #1, TAB(1); USING "TOT REG HRS ####.##"; Check(1).RegHrsPaid;
    PRINT #1, TAB(25); USING "#####.##"; Check(1).TotRegWage;
    PRINT #1, TAB(51); Check(1).CDED(6).DCode;
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(6).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(6).YTDDAmt;

    '--Line 11
    PRINT #1, TAB(1); USING "OT  HRS     ####.##"; Check(1).OTHrsPaid;
    PRINT #1, TAB(25); USING "#####.##"; Check(1).TotOTWage;
    PRINT #1, TAB(51); Check(1).CDED(7).DCode;
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(7).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(7).YTDDAmt

    '--Line 12
    PRINT #1, TAB(1); Check(1).AEarn(1).DCode;
    PRINT #1, TAB(25); USING "#####.##"; Check(1).AEarn(1).DAmt;
    PRINT #1, TAB(51); Check(1).CDED(8).DCode;
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(8).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(8).YTDDAmt

    '--Line 13
    PRINT #1, TAB(1); Check(1).AEarn(2).DCode;
    PRINT #1, TAB(25); USING "#####.##"; Check(1).AEarn(2).DAmt;
    PRINT #1, TAB(51); Check(1).CDED(9).DCode;
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(9).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(9).YTDDAmt

    '--Line 14
    PRINT #1, TAB(1); Check(1).AEarn(3).DCode;
    PRINT #1, TAB(25); USING "#####.##"; Check(1).AEarn(3).DAmt;
    PRINT #1, TAB(51); Check(1).CDED(10).DCode;
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(10).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(10).YTDDAmt

   '--Line 15
    PRINT #1, TAB(1); "GROSS PAY";
    PRINT #1, TAB(25); USING "#####.##"; Check(1).GrossPay;
    PRINT #1, TAB(39); USING "#####.##"; Check(1).YTDGrossPay;
    PRINT #1, TAB(51); Check(1).CDED(11).DCode;
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(11).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(11).YTDDAmt

    '--Line 16
    PRINT #1, "FED TAX"; TAB(25); USING "#####.##"; Check(1).FedTaxAmt;
    PRINT #1, TAB(39); USING "#####.##"; Check(1).YTDFederal;
    PRINT #1, TAB(51); Check(1).CDED(12).DCode;
    PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(12).DAmt;
    PRINT #1, TAB(86); USING "####.##"; Check(1).CDED(12).YTDDAmt

    '--Line 17
    PRINT #1, "STA TAX"; TAB(25); USING "#####.##"; Check(1).StaTaxAmt;
    PRINT #1, TAB(39); USING "#####.##"; Check(1).YTDState;
    PRINT #1, TAB(51); "NET PAY";
    PRINT #1, TAB(70); USING "#####.##"; Check(1).NetPay;
    PRINT #1, TAB(85); USING "#####.##"; Check(1).YTDNetPay

    '--Line 18 - Last line of stub
    PRINT #1, "FICA"; TAB(25); USING "#####.##"; Round(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
    PRINT #1, TAB(39); USING "#####.##"; Check(1).YTDSocial + Check(1).YTDMedicare
    
    IF Check(1).DDFlag = True THEN
    '--Line 19 - First line of check
      PRINT #1,
      PRINT #1,
      PRINT #1,
      PRINT #1, TAB(5); "NON NEGOTIABLE     NON NEGOTIABLE     NON NEGOTIABLE    NON NEGOTIABLE    NON NEGOTIABLE"
      PRINT #1, TAB(5); "NON NEGOTIABLE     NON NEGOTIABLE     NON NEGOTIABLE    NON NEGOTIABLE    NON NEGOTIABLE"
      PRINT #1, TAB(5); "NON NEGOTIABLE     NON NEGOTIABLE     NON NEGOTIABLE    NON NEGOTIABLE    NON NEGOTIABLE"
      PRINT #1, TAB(5); "NON NEGOTIABLE     NON NEGOTIABLE     NON NEGOTIABLE    NON NEGOTIABLE    NON NEGOTIABLE"
      PRINT #1, TAB(5); "NON NEGOTIABLE     NON NEGOTIABLE     NON NEGOTIABLE    NON NEGOTIABLE    NON NEGOTIABLE"
      PRINT #1, TAB(5); "NON NEGOTIABLE     NON NEGOTIABLE     NON NEGOTIABLE    NON NEGOTIABLE    NON NEGOTIABLE"
      PRINT #1,
      PRINT #1,
      PRINT #1, TAB(68); Num2Date$(CheckDate); TAB(84); USING "$$#####,.##"; Check(1).NetPay
      PRINT #1,
      PRINT #1, TAB(10); Check(1).EmpName
      PRINT #1, TAB(10); Check(1).EmpAddr1; TAB(61); "NON NEGOTIABLE    NON NEGOTIABLE"
      PRINT #1, TAB(10); EmpRec2(1).EmpAddr2; TAB(61); "NON NEGOTIABLE    NON NEGOTIABLE"
      PRINT #1, TAB(10); Check(1).EmpCity; " "; Check(1).EmpState; " "; Check(1).EmpZip; TAB(61); "NON NEGOTIABLE    NON NEGOTIABLE"
      PRINT #1, TAB(5);
      PRINT #1, TAB(5); "NON NEGOTIABLE     NON NEGOTIABLE     NON NEGOTIABLE    NON NEGOTIABLE    NON NEGOTIABLE"
      PRINT #1, TAB(5); "NON NEGOTIABLE     NON NEGOTIABLE     NON NEGOTIABLE    NON NEGOTIABLE    NON NEGOTIABLE"
      PRINT #1, TAB(5); "NON NEGOTIABLE     NON NEGOTIABLE     NON NEGOTIABLE    NON NEGOTIABLE    NON NEGOTIABLE"
      PRINT #1, "~" '--Line 39
    ELSE
    '--Line 19 - First line of check
      PRINT #1,
      PRINT #1,
      PRINT #1,
      PRINT #1,
      PRINT #1, TAB(83); USING "#######"; TCheckNum&
      PRINT #1,
      PRINT #1,
      PRINT #1,
      PRINT #1,
      PRINT #1, TAB(3); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##")) '--Line 30
      PRINT #1,
      PRINT #1, TAB(68); Num2Date$(CheckDate); TAB(84); USING "$$#####,.##"; Check(1).NetPay
      PRINT #1,
      PRINT #1, TAB(10); Check(1).EmpName
      PRINT #1, TAB(10); Check(1).EmpAddr1
      PRINT #1, TAB(10); EmpRec2(1).EmpAddr2
      PRINT #1, TAB(10); Check(1).EmpCity; " "; Check(1).EmpState; " "; Check(1).EmpZip
      PRINT #1,
      PRINT #1,
      PRINT #1,
      PRINT #1, "~" '--Line 39
    END IF
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

