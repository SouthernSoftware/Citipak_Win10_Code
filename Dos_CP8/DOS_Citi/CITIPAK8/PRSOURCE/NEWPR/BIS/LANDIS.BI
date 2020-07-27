'This is "LANDIS" check format 9/26/96

     'PRINT #1, CHR$(27) + CHR$(58);'sets compressed mode   12 dpi
     PRINT #1, "~"
     PRINT #1,
     PRINT #1, "  "; LTRIM$(Check(1).EMPNO); TAB(10); LEFT$(Check(1).EmpName, 26); Check(1).EmpSSN;
     PRINT #1, " "; Num2Date(Check(1).PayEndDate); " "; Num2Date(Check(1).CheckDate);
     PRINT #1, "  "; Check(1).CheckNum
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, " "; USING "#####.##"; Check(1).RegHrsWork; TAB(11); Check(1).BaseRate;
     PRINT #1, "  "; USING "#####.##"; Check(1).TotRegWage; Check(1).OTHrsPaid; TAB(45); Check(1).TotOTWage; TAB(60); Check(1).TotAdditEarn; TAB(70); Check(1).GrossPay
'----
     PRINT #1, "  "; USING "COMP USE ####.##"; Check(1).CompUsed;
     PRINT #1, USING "  VACT USE ####.##"; Check(1).VacUsed;
     PRINT #1, USING "  SICK USE ####.##"; Check(1).SickUsed

     PRINT #1, "  "; USING "COMP BAL ####.##"; Check(1).CompBal;
     PRINT #1, USING "  VACT BAL ####.##"; Check(1).VactBal;
     PRINT #1, USING "  SICK BAL ####.##"; Check(1).SickBal
     PRINT #1,
     PRINT #1,
     PRINT #1, USING "#####.##"; Check(1).RetireAmt; TAB(11); Check(1).FedTaxAmt; Round#(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
     PRINT #1, "  "; USING "#####.##"; Check(1).StaTaxAmt;
     PRINT #1, "   "; USING "#####.##"; Check(1).TotDedAmt; TAB(70); Check(1).NetPay
     PRINT #1, TAB(3); Check(1).CDED(1).DCode;
     PRINT #1, TAB(20); USING "####.##"; Check(1).CDED(1).DAmt;
     PRINT #1, TAB(31); USING "####.##"; Check(1).CDED(1).YTDDAmt;

     PRINT #1, TAB(43); Check(1).CDED(2).DCode;
     PRINT #1, TAB(60); USING "####.##"; Check(1).CDED(2).DAmt;
     PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(2).YTDDAmt

     PRINT #1, TAB(3); Check(1).CDED(3).DCode;
     PRINT #1, TAB(20); USING "####.##"; Check(1).CDED(3).DAmt;
     PRINT #1, TAB(31); USING "####.##"; Check(1).CDED(3).YTDDAmt;

     PRINT #1, TAB(43); Check(1).CDED(4).DCode;
     PRINT #1, TAB(60); USING "####.##"; Check(1).CDED(4).DAmt;
     PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(4).YTDDAmt

     PRINT #1, TAB(3); Check(1).CDED(5).DCode;
     PRINT #1, TAB(20); USING "####.##"; Check(1).CDED(5).DAmt;
     PRINT #1, TAB(31); USING "####.##"; Check(1).CDED(5).YTDDAmt;

     PRINT #1, TAB(43); Check(1).CDED(6).DCode;
     PRINT #1, TAB(60); USING "####.##"; Check(1).CDED(6).DAmt;
     PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(6).YTDDAmt

'************
     PRINT #1,
     PRINT #1,
     PRINT #1, USING "#####.##"; TAB(5); Check(1).YTDFederal; TAB(15); Round(Check(1).YTDSocial + Check(1).YTDMedicare); Check(1).YTDState; TAB(36); Check(1).YTDTotDed; TAB(70); Check(1).YTDGrossPay

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
     PRINT #1, TAB(3); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##"))  'Print the whole number part
     PRINT #1, TAB(49); Num2Date(Check(1).CheckDate); TAB(61); Check(1).CheckNum; TAB(70); USING "####.##"; Check(1).NetPay
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(10); Check(1).EmpName
     PRINT #1, TAB(10); Check(1).EmpAddr1
     PRINT #1, TAB(10); Check(1).EmpCity; " "; Check(1).EmpState; " "; Check(1).EmpZip
'    GOTO skip
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, "~"
'-=-=-=-=-=-=-=-=-=-=-*

