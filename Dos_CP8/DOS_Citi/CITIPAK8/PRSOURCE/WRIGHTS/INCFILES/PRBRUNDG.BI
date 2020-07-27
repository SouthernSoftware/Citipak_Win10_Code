'42 line
    '--Line 1
    PRINT #1, "~"; TAB(68); Num2Date$(Check(1).PayEndDate)
    PRINT #1,
    PRINT #1, 'line 3
    PRINT #1, TAB(70); USING "#####.##"; Check(1).GrossPay
    PRINT #1, TAB(2); USING "####.##"; Check(1).BaseRate;
    PRINT #1, TAB(13); USING "####.##"; Check(1).RegHrsWork;
    PRINT #1, TAB(22); USING "####.##"; Check(1).OTHrsPaid;
    PRINT #1, TAB(32); USING "#####.##"; Check(1).TotRegWage;
    PRINT #1, TAB(43); USING "#####.##"; Check(1).TotOTWage;
    PRINT #1, TAB(54); USING "#####.##"; Check(1).TotAdditEarn
    PRINT #1,
    PRINT #1, TAB(70); USING "#####.##"; Check(1).NetPay
    PRINT #1,

    PRINT #1, TAB(55); USING "####.##"; Check(1).CDED(1).DAmt;
    PRINT #1, TAB(63); LEFT$(Check(1).CDED(1).DCode, 3);
    PRINT #1, TAB(68); USING "####.##"; Check(1).CDED(2).DAmt;
    PRINT #1, TAB(76); LEFT$(Check(1).CDED(2).DCode, 3)

    PRINT #1, TAB(2); USING "#####.##"; Check(1).FedTaxAmt;
    PRINT #1, TAB(13); USING "#####.##"; Round(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
    PRINT #1, TAB(24); USING "#####.##"; Check(1).StaTaxAmt;
    PRINT #1, TAB(55); USING "####.##"; Check(1).CDED(3).DAmt;
    PRINT #1, TAB(63); LEFT$(Check(1).CDED(3).DCode, 3);
    PRINT #1, TAB(68); USING "####.##"; Check(1).CDED(4).DAmt;
    PRINT #1, TAB(76); LEFT$(Check(1).CDED(4).DCode, 3)
    PRINT #1, TAB(55); USING "####.##"; Check(1).CDED(5).DAmt;
    PRINT #1, TAB(63); LEFT$(Check(1).CDED(5).DCode, 3);
    PRINT #1, TAB(68); USING "####.##"; Check(1).CDED(6).DAmt;
    PRINT #1, TAB(76); LEFT$(Check(1).CDED(6).DCode, 3)
    PRINT #1, TAB(55); USING "####.##"; Check(1).CDED(7).DAmt;
    PRINT #1, TAB(63); LEFT$(Check(1).CDED(7).DCode, 3);
    PRINT #1, TAB(68); USING "####.##"; Check(1).CDED(8).DAmt;
    PRINT #1, TAB(76); LEFT$(Check(1).CDED(8).DCode, 3)

    PRINT #1, TAB(2); USING "#####.##"; Check(1).YTDFederal;
    PRINT #1, TAB(13); USING "#####.##"; Round(Check(1).YTDSocial + Check(1).YTDMedicare);
    PRINT #1, TAB(24); USING "#####.##"; Check(1).YTDState;

    PRINT #1, TAB(55); USING "####.##"; Check(1).CDED(9).DAmt;
    PRINT #1, TAB(63); LEFT$(Check(1).CDED(9).DCode, 3);
    PRINT #1, TAB(68); USING "####.##"; Check(1).CDED(10).DAmt;
    PRINT #1, TAB(76); LEFT$(Check(1).CDED(10).DCode, 3)
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1, TAB(55); "Check Date"; TAB(71); "Amount"
    PRINT #1,
    PRINT #1, TAB(55); Num2Date$(CheckDate); TAB(67); USING "$$#####,.##"; Check(1).NetPay
    PRINT #1,
    PRINT #1, TAB(11); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##")) '--Line 30
    PRINT #1,
    PRINT #1, TAB(11); Check(1).EmpName
    PRINT #1, TAB(11); Check(1).EmpAddr1
    PRINT #1, TAB(11); Check(1).EmpCity; " "; Check(1).EmpState; " "; Check(1).EmpZip
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1,
    PRINT #1, "~" '--Line 42
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

