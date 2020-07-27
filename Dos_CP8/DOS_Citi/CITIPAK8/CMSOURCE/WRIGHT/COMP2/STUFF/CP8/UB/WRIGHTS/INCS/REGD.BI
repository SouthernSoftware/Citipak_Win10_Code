'This is REGION-D check format   5/9/94
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

      PRINT #1, TAB(3); "LGERS";
      PRINT #1, TAB(20); USING "####.##"; Check(1).RetireAmt;
      PRINT #1, TAB(31); USING "####.##"; Check(1).YTDRetire

      PRINT #1, TAB(3); Check(1).CDED(1).DCode;
      PRINT #1, TAB(20); USING "####.##"; Check(1).CDED(1).DAmt;
      PRINT #1, TAB(31); USING "####.##"; Check(1).CDED(1).YTDDAmt;

      PRINT #1, TAB(43); Check(1).CDED(7).DCode;
      PRINT #1, TAB(60); USING "####.##"; Check(1).CDED(7).DAmt;
      PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(7).YTDDAmt

      PRINT #1, TAB(3); Check(1).CDED(2).DCode;
      PRINT #1, TAB(20); USING "####.##"; Check(1).CDED(2).DAmt;
      PRINT #1, TAB(31); USING "####.##"; Check(1).CDED(2).YTDDAmt;

      PRINT #1, TAB(43); Check(1).CDED(8).DCode;
      PRINT #1, TAB(60); USING "####.##"; Check(1).CDED(8).DAmt;
      PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(8).YTDDAmt

      PRINT #1, TAB(3); Check(1).CDED(3).DCode;
      PRINT #1, TAB(20); USING "####.##"; Check(1).CDED(3).DAmt;
      PRINT #1, TAB(31); USING "####.##"; Check(1).CDED(3).YTDDAmt;

      PRINT #1, TAB(43); Check(1).CDED(9).DCode;
      PRINT #1, TAB(60); USING "####.##"; Check(1).CDED(9).DAmt;
      PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(9).YTDDAmt

      PRINT #1, TAB(3); Check(1).CDED(4).DCode;
      PRINT #1, TAB(20); USING "####.##"; Check(1).CDED(4).DAmt;
      PRINT #1, TAB(31); USING "####.##"; Check(1).CDED(4).YTDDAmt;

      PRINT #1, TAB(43); Check(1).CDED(10).DCode;
      PRINT #1, TAB(60); USING "####.##"; Check(1).CDED(10).DAmt;
      PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(10).YTDDAmt

      PRINT #1, TAB(3); Check(1).CDED(5).DCode;
      PRINT #1, TAB(20); USING "####.##"; Check(1).CDED(5).DAmt;
      PRINT #1, TAB(31); USING "####.##"; Check(1).CDED(5).YTDDAmt;

      PRINT #1, TAB(43); Check(1).CDED(11).DCode;
      PRINT #1, TAB(60); USING "####.##"; Check(1).CDED(11).DAmt;
      PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(11).YTDDAmt

      PRINT #1, TAB(3); Check(1).CDED(6).DCode;
      PRINT #1, TAB(20); USING "####.##"; Check(1).CDED(6).DAmt;
      PRINT #1, TAB(31); USING "####.##"; Check(1).CDED(6).YTDDAmt;

      PRINT #1, TAB(43); Check(1).CDED(12).DCode;
      PRINT #1, TAB(60); USING "####.##"; Check(1).CDED(12).DAmt;
      PRINT #1, TAB(71); USING "####.##"; Check(1).CDED(12).YTDDAmt

      PRINT #1, TAB(3); "Vac Bal.";
      PRINT #1, TAB(31); USING "####.##"; Check(1).VactBal;

      PRINT #1, TAB(43); "Sick Bal.";
      PRINT #1, TAB(71); USING "####.##"; Check(1).SickBal

      'PRINT #1, TAB(3); "Comp Bal.";
      PRINT #1, ' TAB(31); USING "####.##"; Check(1).CompBal

      FOR Cnt2 = 1 TO 6
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
      FOR Cnt2 = 35 TO 42
        PRINT #1, ""
      NEXT
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

