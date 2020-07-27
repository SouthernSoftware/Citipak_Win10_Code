'This is the "Mt.Gilead" check format 01/09/96

     'PRINT #1, CHR$(27) + CHR$(58); 'sets compressed mode   12 dpi
     PRINT #1, ""
     PRINT #1, TAB(43); Num2Date$(CheckDate)
     PRINT #1,
     PRINT #1, QPTrim$(Check(1).EmpName);
     PRINT #1, TAB(46); USING "$$####,#.##"; Check(1).NetPay
     PRINT #1,
     PRINT #1,
     PRINT #1, USING "######.##"; Check(1).GrossPay;
     PRINT #1, TAB(10); USING "#####.##"; Round(Check(1).MedTaxAmt + Check(1).SocTaxAmt + Check(1).StaTaxAmt + Check(1).FedTaxAmt);
     PRINT #1, TAB(19); USING "#####.##"; Round(Check(1).TotDedAmt + Check(1).RetireAmt);
     PRINT #1, TAB(27); USING "######.##"; Check(1).NetPay;
     PRINT #1, TAB(38); "FED TAX"; TAB(46); USING "#####.##"; Check(1).FedTaxAmt; TAB(55); Check(1).YTDFederal
     PRINT #1, TAB(38); "STA TAX"; TAB(46); USING "#####.##"; Check(1).StaTaxAmt; TAB(55); Check(1).YTDState

     PRINT #1, " Reg Hrs"; TAB(12); USING "####.##"; Check(1).RegHrsPaid;
     PRINT #1, TAB(19); USING "#####.##"; Check(1).TotRegWage;
     PRINT #1, TAB(28); USING "#####.##"; Round(EmpRec3(1).YTDRegPay + Check(1).TotRegWage);

     PRINT #1, TAB(38); "FICA"; TAB(46); USING "#####.##"; Round(Check(1).MedTaxAmt + Check(1).SocTaxAmt); TAB(55); Round(Check(1).YTDSocial + Check(1).YTDMedicare)

     PRINT #1, " OT  Hrs"; TAB(12); USING "####.##"; Check(1).OTHrsPaid;
     PRINT #1, TAB(19); USING "#####.##"; Check(1).TotOTWage;
     PRINT #1, TAB(28); USING "#####.##"; Round(EmpRec3(1).YTDOTPay + Check(1).TotOTWage);
     PRINT #1, TAB(38); "RET"; TAB(46); USING "#####.##"; Check(1).RetireAmt; TAB(55); Round(Check(1).RetireAmt + Check(1).YTDRetire)

     'PRINT #1, ;  '    ####.##";
     PRINT #1, "ADD EARN"; TAB(19); USING "#####.##"; Check(1).TotAdditEarn;
     PRINT #1, TAB(28); USING "#####.##"; Round(EmpRec3(1).YTDEarnT + Check(1).TotAdditEarn);
     PRINT #1, TAB(38); LEFT$(Check(1).CDED(1).DCode, 7); TAB(46); USING "#####.##"; Check(1).CDED(1).DAmt; TAB(55); Check(1).CDED(1).YTDDAmt

     PRINT #1, "VACT BAL"; TAB(12); USING "####.##"; Check(1).VactBal;
     PRINT #1, TAB(38); LEFT$(Check(1).CDED(2).DCode, 7); TAB(46); USING "#####.##"; Check(1).CDED(2).DAmt; TAB(55); Check(1).CDED(2).YTDDAmt

     PRINT #1, "SICK BAL"; TAB(12); USING "####.##"; Check(1).SickBal;
     PRINT #1, TAB(38); LEFT$(Check(1).CDED(3).DCode, 7); TAB(46); USING "#####.##"; Check(1).CDED(3).DAmt; TAB(55); Check(1).CDED(3).YTDDAmt

     PRINT #1, "COMP BAL"; TAB(12); USING "####.##"; Check(1).CompBal;
     PRINT #1, TAB(38); LEFT$(Check(1).CDED(4).DCode, 7); TAB(46); USING "#####.##"; Check(1).CDED(4).DAmt; TAB(55); Check(1).CDED(4).YTDDAmt

     PRINT #1, TAB(38); LEFT$(Check(1).CDED(5).DCode, 7); TAB(46); USING "#####.##"; Check(1).CDED(5).DAmt; TAB(55); Check(1).CDED(5).YTDDAmt
     PRINT #1, TAB(38); LEFT$(Check(1).CDED(6).DCode, 7); TAB(46); USING "#####.##"; Check(1).CDED(6).DAmt; TAB(55); Check(1).CDED(6).YTDDAmt
     
     OtherDed# = Round(Check(1).CDED(7).DAmt + Check(1).CDED(8).DAmt + Check(1).CDED(9).DAmt + Check(1).CDED(10).DAmt + Check(1).CDED(11).DAmt + Check(1).CDED(12).DAmt)
     YTDOtherDed# = Round(Check(1).CDED(7).YTDDAmt + Check(1).CDED(8).YTDDAmt + Check(1).CDED(9).YTDDAmt + Check(1).CDED(10).YTDDAmt + Check(1).CDED(11).YTDDAmt + Check(1).CDED(12).YTDDAmt)
     PRINT #1, TAB(38); "DED 7-12"; TAB(46); USING "#####.##"; OtherDed#; TAB(55); YTDOtherDed#


      PRINT #1, ""
      PRINT #1, ""
      PRINT #1, ""
      PRINT #1, ""
      PRINT #1, ""
     PRINT #1,
     PRINT #1, TAB(58); USING "#######"; TCheckNum&
     PRINT #1,
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(50); Num2Date$(CheckDate); TAB(60); USING "$$####,#.##"; Check(1).NetPay
     PRINT #1,
     PRINT #1,
     PRINT #1, TAB(8); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##"))'Print the whole number part
     PRINT #1, ""
     PRINT #1, ""
     PRINT #1, ""
     PRINT #1, TAB(12); Check(1).EmpName
     PRINT #1, TAB(12); Check(1).EmpAddr1
     PRINT #1, TAB(12); Check(1).EmpCity; " "; Check(1).EmpState; " "; Check(1).EmpZip
     PRINT #1, ""
     PRINT #1, ""
     PRINT #1, ""
     PRINT #1, ""
     PRINT #1, ""
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

