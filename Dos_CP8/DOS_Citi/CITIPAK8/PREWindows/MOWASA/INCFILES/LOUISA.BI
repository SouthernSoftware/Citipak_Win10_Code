'This is the "LOUISA" check format 02/17/95

     TempEmp$ = LEFT$(Check(1).EmpName, 33)
'     PRINT #1, CHR$(27) + CHR$(58); "***"; 'sets compressed mode   12 dpi
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, USING "###.##"; Check(1).RegHrsWork;
     PRINT #1, TAB(8); USING "###.##"; Check(1).OTHrsPaid;
     PRINT #1, TAB(16); USING "####.##"; Check(1).BaseRate;
     PRINT #1, TAB(27); USING "####.##"; Check(1).TotRegWage;
     PRINT #1, TAB(38); USING "####.##"; Check(1).TotOTWage;
     PRINT #1, TAB(48); QPTrim$(Check(1).AEarn(1).DCode);
     PRINT #1, TAB(64); USING "####.##"; Check(1).AEarn(1).DAmt;
     PRINT #1, TAB(77); USING "####.##"; Check(1).GrossPay;
     PRINT #1, TAB(86); Num2Date$(Check(1).PayEndDate)

     PRINT #1, TAB(48); QPTrim$(Check(1).AEarn(2).DCode);
     PRINT #1, TAB(64); USING "####.##"; Check(1).AEarn(2).DAmt;
     PRINT #1, TAB(81); "YTD"

     PRINT #1, TAB(48); QPTrim$(Check(1).AEarn(3).DCode);
     PRINT #1, TAB(64); USING "####.##"; Check(1).AEarn(3).DAmt;
     PRINT #1, TAB(76); USING "#####.##"; Check(1).YTDGrossPay

     PRINT #1, TAB(48); "YTD OTH"; TAB(63); USING "#####.##"; Round#(EmpRec3(1).YTDEarnT + Check(1).TotAdditEarn);
     PRINT #1, TAB(87); USING "####.##"; Round(Check(1).TotRegWage + Check(1).TotOTWage + Check(1).TotAdditEarn)
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(37); Check(1).CDED(1).DCode;
     PRINT #1, TAB(49); Check(1).CDED(2).DCode;
     PRINT #1, TAB(61); Check(1).CDED(3).DCode;
     PRINT #1, TAB(73); Check(1).CDED(4).DCode

     PRINT #1, TAB(3); USING "####.##"; Round#(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
     PRINT #1, TAB(16); USING "####.##"; Check(1).FedTaxAmt;
     PRINT #1, TAB(27); USING "####.##"; Check(1).StaTaxAmt;
     PRINT #1, TAB(41); USING "####.##"; Check(1).CDED(1).DAmt;
     PRINT #1, TAB(53); USING "####.##"; Check(1).CDED(2).DAmt;
     PRINT #1, TAB(65); USING "####.##"; Check(1).CDED(3).DAmt;
     PRINT #1, TAB(77); USING "####.##"; Check(1).CDED(4).DAmt

     PRINT #1, TAB(7); "YTD";
     PRINT #1, TAB(20); "YTD";
     PRINT #1, TAB(31); "YTD";
     PRINT #1, TAB(37); Check(1).CDED(5).DCode;
     PRINT #1, TAB(49); Check(1).CDED(6).DCode;
     PRINT #1, TAB(61); Check(1).CDED(7).DCode;
     PRINT #1, TAB(73); Check(1).CDED(8).DCode
     PRINT #1, TAB(2); USING "#####.##"; Round#(Check(1).YTDSocial + Check(1).YTDMedicare);
     PRINT #1, TAB(15); USING "#####.##"; Check(1).YTDFederal;
     PRINT #1, TAB(26); USING "#####.##"; Check(1).YTDState;
     PRINT #1, TAB(41); USING "####.##"; Check(1).CDED(5).DAmt;
     PRINT #1, TAB(53); USING "####.##"; Check(1).CDED(6).DAmt;
     PRINT #1, TAB(65); USING "####.##"; Check(1).CDED(7).DAmt;
     PRINT #1, TAB(77); USING "####.##"; Check(1).CDED(8).DAmt

     PRINT #1, TAB(37); Check(1).CDED(9).DCode;
     PRINT #1, TAB(49); Check(1).CDED(10).DCode;
     PRINT #1, TAB(61); Check(1).CDED(11).DCode;
     PRINT #1, TAB(73); Check(1).CDED(12).DCode;
     PRINT #1, TAB(87); USING "####.##"; Check(1).TotDedAmt

     PRINT #1, TAB(41); USING "####.##"; Check(1).CDED(9).DAmt;
     PRINT #1, TAB(53); USING "####.##"; Check(1).CDED(10).DAmt;
     PRINT #1, TAB(65); USING "####.##"; Check(1).CDED(11).DAmt;
     PRINT #1, TAB(77); USING "####.##"; Check(1).CDED(12).DAmt
     'PRINT #1,
'     PRINT #1, TAB(3); "Retirement";
'     PRINT #1, TAB(6); USING "####.##"; Check(1).RetireAmt;
'     PRINT #1, TAB(61); USING "####.##"; Check(1).EICAmt;

     PRINT #1, TAB(38); "YTD Net        YTD Ded"
     PRINT #1, QPTrim$(TempEmp$); TAB(37); USING "#####.##"; Check(1).YTDNetPay; TAB(52); Check(1).YTDTotDed

     PRINT #1, "SSN: "; Check(1).EmpSSN;
     PRINT #1, TAB(87); USING "####.##"; Check(1).NetPay
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(11); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##"))
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(11); Check(1).EmpName; TAB(52); Num2Date$(Check(1).CheckDate); TAB(81); USING "$$#####.##"; Check(1).NetPay
     PRINT #1, TAB(11); Check(1).EmpAddr1
     PRINT #1, TAB(11); Check(1).EmpCity; " "; Check(1).EmpState; " "; Check(1).EmpZip
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
'-=-=-=-=-=-=-=-=-=-=-*

