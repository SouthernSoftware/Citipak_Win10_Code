'This is "BURNSVILLE" check format   2/17/95
      PRINT #1, ""
      PRINT #1, ""
      PRINT #1, ""
      PRINT #1, Check(1).EmpNo;
      PRINT #1, TAB(16); Check(1).EmpSSN;
      PRINT #1, TAB(31); Num2Date$(Check(1).PayEndDate); TAB(43); Num2Date$(CheckDate);
      PRINT #1, TAB(56); USING "#####.##"; Check(1).BaseRate
      PRINT #1, ""
      PRINT #1, ""
      PRINT #1, TAB(4); USING "#####.##"; Check(1).GrossPay;
      PRINT #1, TAB(19); USING "####.##"; Check(1).FedTaxAmt;
      PRINT #1, TAB(32); USING "####.##"; Check(1).StaTaxAmt;
      PRINT #1, TAB(45); USING "####.##"; Round#(Check(1).MedTaxAmt + Check(1).SocTaxAmt);
      PRINT #1, TAB(58); USING "####.##"; Check(1).TotDedAmt;
      PRINT #1, TAB(71); USING "#####.##"; Check(1).NetPay
      PRINT #1, ""
      PRINT #1, ""
      PRINT #1, TAB(4); USING "#####.##"; Check(1).YTDGrossPay;
      PRINT #1, TAB(19); USING "####.##"; Check(1).YTDFederal;
      PRINT #1, TAB(32); USING "####.##"; Check(1).YTDState;
      PRINT #1, TAB(45); USING "####.##"; Round#(Check(1).YTDSocial + Check(1).YTDMedicare);

      PRINT #1, TAB(58); USING "#####.##"; Check(1).YTDTotDed;
      PRINT #1, TAB(71); USING "#####.##"; Check(1).YTDNetPay

      PRINT #1, ""
      PRINT #1, ""

      PRINT #1, "REG HRS"; TAB(20); USING "###.##"; Check(1).RegHrsWork;

      PRINT #1, TAB(43); "RETIRE"; TAB(60); USING "####.##"; Check(1).RetireAmt;
      PRINT #1, TAB(73); USING "####.##"; Check(1).YTDRetire

      PRINT #1, "OT HRS"; TAB(20); USING "###.##"; Check(1).OTHrsPaid;
      PRINT #1, TAB(43); "EIC"; TAB(60); USING "####.##"; Check(1).EICAmt

      PRINT #1, "VAC HRS"; TAB(20); USING "###.##"; Check(1).VacUsed;
      PRINT #1, TAB(32); USING "####.##"; Check(1).VactBal;
      PRINT #1, TAB(43); Check(1).CDED(1).DCode; TAB(60); USING "####.##"; Check(1).CDED(1).DAmt;
      PRINT #1, TAB(73); USING "####.##"; Check(1).CDED(1).YTDDAmt

      PRINT #1, "COMP HRS"; TAB(20); USING "###.##"; Check(1).CompUsed;
      PRINT #1, TAB(32); USING "####.##"; Check(1).CompBal;
      PRINT #1, TAB(43); Check(1).CDED(2).DCode; TAB(60); USING "####.##"; Check(1).CDED(2).DAmt;
      PRINT #1, TAB(73); USING "####.##"; Check(1).CDED(2).YTDDAmt

      PRINT #1, "SICK HRS"; TAB(20); USING "###.##"; Check(1).SickUsed;
      PRINT #1, TAB(32); USING "####.##"; Check(1).SickBal;
      PRINT #1, TAB(43); Check(1).CDED(3).DCode; TAB(60); USING "####.##"; Check(1).CDED(3).DAmt;
      PRINT #1, TAB(73); USING "####.##"; Check(1).CDED(3).YTDDAmt

      PRINT #1, "TOT HRS"; TAB(20); USING "###.##"; Check(1).RegHrsPaid;
      PRINT #1, TAB(43); Check(1).CDED(4).DCode; TAB(60); USING "####.##"; Check(1).CDED(4).DAmt;
      PRINT #1, TAB(73); USING "####.##"; Check(1).CDED(4).YTDDAmt

      PRINT #1, "REG PAY"; TAB(19); USING "####.##"; Check(1).TotRegWage;

      PRINT #1, TAB(31); USING "#####.##"; EmpRec3(1).YTDRegPay;

      PRINT #1, TAB(43); Check(1).CDED(5).DCode; TAB(60); USING "####.##"; Check(1).CDED(5).DAmt;
      PRINT #1, TAB(73); USING "####.##"; Check(1).CDED(5).YTDDAmt

      PRINT #1, "OT PAY"; TAB(19); USING "####.##"; Check(1).TotOTWage;
      PRINT #1, TAB(31); USING "#####.##"; EmpRec3(1).YTDOTPay;
      PRINT #1, TAB(43); Check(1).CDED(6).DCode; TAB(60); USING "####.##"; Check(1).CDED(6).DAmt;
      PRINT #1, TAB(73); USING "####.##"; Check(1).CDED(6).YTDDAmt

      PRINT #1, "ADD EARN"; TAB(19); USING "####.##"; Check(1).TotAdditEarn;
      PRINT #1, TAB(31); USING "#####.##"; EmpRec3(1).YTDEarnT;
      TDAmt# = 0
      YTDAmt# = 0

      FOR cnt2 = 7 TO 12
        TDAmt# = Round(TDAmt# + Check(1).CDED(cnt2).DAmt)
        YTDAmt# = Round(YTDAmt# + Check(1).CDED(cnt2).YTDDAmt)
      NEXT

      PRINT #1, TAB(43); "DED 7-12";
      PRINT #1, TAB(60); USING "####.##"; TDAmt#; TAB(73); YTDAmt#

      FOR cnt2 = 1 TO 6
        PRINT #1, ""
      NEXT
'STOP
      PRINT #1, TAB(59); USING "#####"; Check(1).CheckNum
      PRINT #1, ""
      PRINT #1, ""
      PRINT #1, ""
      PRINT #1, TAB(4); Num2Date$(CheckDate);
      PRINT #1, TAB(18); Check(1).EmpName;
      PRINT #1, TAB(54); USING "$$####,#.##"; Check(1).NetPay
      PRINT #1, ""
      PRINT #1, TAB(3); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##")) 'Print the whole number part
      FOR cnt2 = 35 TO 42
        PRINT #1, ""
      NEXT
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

