'This is Pond Creek check format   07/08/96

      PRINT #1,
      PRINT #1,
      PRINT #1, TAB(51); USING "####.##"; Check(1).RegHrsWork;
      PRINT #1, TAB(59); USING "#####.##"; Check(1).BaseRate;
      PRINT #1, TAB(69); USING "#####.##"; Check(1).TotRegWage
      PRINT #1,
      PRINT #1, TAB(51); USING "####.##"; Check(1).OTHrsPaid;
      PRINT #1, TAB(59); USING "#####.##"; EmpRec2(1).EMPORATE;
      PRINT #1, TAB(69); USING "#####.##"; Check(1).TotOTWage
      PRINT #1,
      PRINT #1, TAB(69); USING "#####.##"; Check(1).GrossPay
      PRINT #1,
      PRINT #1,
      PRINT #1,

      PRINT #1, TAB(8); USING "#####.##"; Check(1).SocTaxAmt;
      PRINT #1, TAB(20); USING "#####.##"; Check(1).FedTaxAmt;
      PRINT #1, TAB(33); USING "#####.##"; Check(1).StaTaxAmt;
      PRINT #1, TAB(42); USING "#####.##"; Check(1).MedTaxAmt;
      PRINT #1, TAB(51); USING "#####.##"; Round(Check(1).CDED(1).DAmt + Check(1).CDED(3).DAmt)
      PRINT #1,
      PRINT #1,


      PRINT #1, TAB(8); USING "#####.##"; Check(1).RetireAmt; TAB(20); Check(1).CDED(2).DAmt
      PRINT #1, TAB(69); USING "#####.##"; Check(1).TotDedAmt
      PRINT #1, TAB(69); USING "#####.##"; Check(1).NetPay
      PRINT #1,
      PRINT #1,
      PRINT #1,
      PRINT #1, Check(1).EmpNo;
      PRINT #1, TAB(20); USING "#####.##"; Check(1).YTDGrossPay;
      PRINT #1, TAB(34); USING "#####.##"; Round(Check(1).YTDSocial + Check(1).YTDMedicare);
      PRINT #1, TAB(46); USING "#####.##"; Check(1).YTDFederal;
      PRINT #1, TAB(57); USING "#####.##"; Check(1).YTDState;
      PRINT #1, TAB(72); USING "######"; Check(1).CheckNum
      PRINT #1,

      PRINT #1, TAB(72); USING "######"; Check(1).CheckNum
      PRINT #1, ""
      PRINT #1, TAB(65); Num2Date$(CheckDate);
      PRINT #1,
      PRINT #1, ""
      PRINT #1,
      PRINT #1,
      PRINT #1, TAB(70); USING "#####.##"; Check(1).NetPay
      PRINT #1,
      PRINT #1, TAB(10); Check(1).EmpName
      PRINT #1,
      PRINT #1, TAB(4); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##")) 'Print the whole number part
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
      PRINT #1,
      PRINT #1, TAB(51); USING "####.##"; Check(1).RegHrsWork;
      PRINT #1, TAB(59); USING "#####.##"; Check(1).BaseRate;
      PRINT #1, TAB(69); USING "#####.##"; Check(1).TotRegWage
      PRINT #1,
      PRINT #1, TAB(51); USING "####.##"; Check(1).OTHrsPaid;
      PRINT #1, TAB(59); USING "#####.##"; EmpRec2(1).EMPORATE;
      PRINT #1, TAB(69); USING "#####.##"; Check(1).TotOTWage
      PRINT #1,
      PRINT #1, TAB(69); USING "#####.##"; Check(1).GrossPay
      PRINT #1,
      PRINT #1,
      PRINT #1,

      PRINT #1, TAB(8); USING "#####.##"; Check(1).SocTaxAmt;
      PRINT #1, TAB(20); USING "#####.##"; Check(1).FedTaxAmt;
      PRINT #1, TAB(33); USING "#####.##"; Check(1).StaTaxAmt;
      PRINT #1, TAB(42); USING "#####.##"; Check(1).MedTaxAmt;
      PRINT #1, TAB(51); USING "#####.##"; Round(Check(1).CDED(1).DAmt + Check(1).CDED(3).DAmt)
      PRINT #1,
      PRINT #1,


      PRINT #1, TAB(8); USING "#####.##"; Check(1).RetireAmt; TAB(20); Check(1).CDED(2).DAmt
      PRINT #1, TAB(69); USING "#####.##"; Check(1).TotDedAmt
      PRINT #1, TAB(69); USING "#####.##"; Check(1).NetPay
      PRINT #1,
      PRINT #1,
      PRINT #1,
      PRINT #1, Check(1).EmpNo;
      PRINT #1, TAB(20); USING "#####.##"; Check(1).YTDGrossPay;
      PRINT #1, TAB(34); USING "#####.##"; Round(Check(1).YTDSocial + Check(1).YTDMedicare);
      PRINT #1, TAB(46); USING "#####.##"; Check(1).YTDFederal;
      PRINT #1, TAB(57); USING "#####.##"; Check(1).YTDState;
      PRINT #1, TAB(72); USING "######"; Check(1).CheckNum
      PRINT #1,

      PRINT #1, TAB(72); USING "######"; Check(1).CheckNum
      PRINT #1, ""
      PRINT #1, TAB(65); Num2Date$(CheckDate);
      PRINT #1,
      PRINT #1, ""
      PRINT #1,
      PRINT #1,
      PRINT #1, TAB(70); USING "#####.##"; Check(1).NetPay
      PRINT #1,
      PRINT #1, TAB(10); Check(1).EmpName
      PRINT #1,
      PRINT #1, TAB(4); SpellNumber$(FUsing$(STR$(Check(1).NetPay), "#####.##")) 'Print the whole number part
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


'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=

