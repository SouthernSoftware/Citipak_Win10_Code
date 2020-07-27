     LPRINT
     LPRINT
     LPRINT
     LPRINT TAB(42); USING "##.##"; taxrate
     LPRINT
     LPRINT
     LPRINT TAB(7); LEFT$(TBRECust(1).PropDesc1, 20); TAB(32); USING "######,#"; (TBRECust(1).LandValue + TBRECust(1).BldgValue) - TBRECust(1).LandUseValue;
     LPRINT TAB(41); USING "#####,#.##"; TBRECust(1).LBCurYrTaxAmount
     LPRINT TAB(7); LEFT$(TBRECust(1).PropDesc2, 20)
     LPRINT TAB(7); LEFT$(TBRECust(1).PropDesc3, 20)
     LPRINT TAB(7); LEFT$(TBRECust(1).PropDesc4, 20); TAB(52); DueDate$; TAB(63); TBRECust(1).LBCurYrBillNumber;
     LPRINT TAB(72); LEFT$(TBRECust(1).MapNumber, 4); "-"; MID$(TBRECust(1).MapNumber, 5, 2); "-"; MID$(TBRECust(1).MapNumber, 7, 2); "-"; MID$(TBRECust(1).MapNumber, 9, 4)
     LPRINT
     LPRINT TAB(67); USING "#####,#.##"; TBRECust(1).LBCurYrTaxAmount
     LPRINT TAB(41); USING "#####,#.##"; TBRECust(1).LBCurYrTaxAmount
     LPRINT
     LPRINT
     LPRINT TAB(30); AcctNumber!
     LPRINT TAB(18); TBRECust(1).CustName
     LPRINT TAB(18); TBRECust(1).Address1
     LPRINT TAB(18); TBRECust(1).Address2
     LPRINT TAB(18); RTRIM$(TBRECust(1).Address3); " "; TBRECust(1).ZipCode
     LPRINT
     LPRINT TAB(52); "ACCT # "; AcctNumber!
     LPRINT TAB(52); "CUST : "; TBRECust(1).CustName
     LPRINT

