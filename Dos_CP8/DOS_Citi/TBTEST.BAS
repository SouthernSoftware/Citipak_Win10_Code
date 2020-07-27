10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbccustrp
30 REM Customer Report Menu
40 REM Last Revision Date: 8-24-87
50 REM by Nick Picerno

   ' Must be compiled using the compile directory
   DEFINT A-Z
   DECLARE FUNCTION QPTrim$ (a$)
   DECLARE FUNCTION num2date$ (Num%)
   DECLARE FUNCTION Date2Num% (Dat$)
   DECLARE FUNCTION WaitKey% ()
   DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
   DECLARE SUB PrintTitle (Title$)
   DECLARE SUB PrintLoading ()
   DECLARE SUB ShowCursor ()
   DECLARE SUB HideCursor ()
   
'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'Qscr.BI'
'$INCLUDE: 'SetCnf.BI'
'$INCLUDE: 'lcTAX.BI'
   
60 CLS :
   TBpath$ = "F:\"                    ' change to "F:\" Before compiling
   CLOSE
   OPEN "R", 1, TBpath$ + "tbyear.dat", 4: FIELD 1, 4 AS Year$: GET 1, 1
   CurrentYear$ = Year$: CLOSE 1
   ON ERROR GOTO networkerror


     REDIM TransLink!(500)

     REDIM TBPPCUST(1) AS TBPPCUST
     REDIM TBPPCustBal(1) AS TBPPBal
     REDIM TBTrans(1) AS TBTrans
     REDIM TBPage(1) AS TBPage1         'Page 1 is Personal Property
     REDIM TBVeh(1) AS PPVehType
     
     WIDTH LPRINT 236
     
100 REM begin program
110 scren$ = "2.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500

1000 REM
1010 scren$ = "2.01.00": CLS : GOSUB 63500
     COLOR 11
     LOCATE 1, 26
      PRINT "Print Personal Property Book"
1020 LOCATE 5, 1
     PRINT TAB(21); "1- Preliminary Pers Property Proof (on greenbar)"
     PRINT
     PRINT TAB(21); "2- Final Pers Property Book (on Forms)"
     PRINT
     PRINT TAB(21); "3- Print Summary (Totals) Pages"
     PRINT
     PRINT TAB(21); "4- Print Town Pers Property Book (on greenbar)"
1025 LOCATE 14, 1
     PRINT TAB(25); "Selection # (1-4) or (ESC-End) .. "; : fl = -1
     GOSUB 62000
     IF cf = 1 THEN RUN "tbccustr"
     IF VAL(iN$) < 1 OR VAL(iN$) > 4 THEN 10
     ON VAL(iN$) GOTO 1100, 2000, 3000, 4000

 

1100 scren$ = "2.01.01": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 25: PRINT "Print Pers Property Book - For Review"
     LOCATE 3, 1
     taxtype$ = "P"
     GOSUB 60000
1150 IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10
     OPEN "O", 10, "TBPREPP.REP": WIDTH #10, 255
     GOSUB 1900
1160 IF taxtype$ = "R" THEN
        OPEN TBpath$ + "TBREname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS LastNam$, 4 AS Mrec$: EndofFile = LOF(4) / 24
        ELSE
        OPEN TBpath$ + "TBPPname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS LastNam$, 4 AS Mrec$: EndofFile = LOF(4) / 24
      END IF
     

1165 FOR District = 1 TO 10                     'MAIN LOOP
     LOCATE 3, 1: COLOR 11
     PRINT "NOW SPOOLING OUT DISTRICT # "; District
     IF District >= 2 AND District <= 8 THEN 1231       'No Districts 2 thru 8 in Personal Prop
1170 FOR X = 1 TO EndofFile
     GET 4, X
     GET 1, CVS(Mrec$), TBPPCUST(1)
     aCCOUNT! = CVS(Mrec$)
     
     
1191 IF TBPPCUST(1).Deleted = "Y" THEN 1230

     IF VAL(TBPPCUST(1).District) <> District THEN 1230
     
     IF VAL(TBPPCUST(1).LBCurYrBillNumber) <= 0 THEN 1230
     PRINT #10, "Acct # "; aCCOUNT!;
     PRINT #10, TAB(20); QPTrim$(TBPPCUST(1).CustName); TAB(60); TBPPCUST(1).LBCurYrBillNumber


     
      P# = P# + TBPPCUST(1).TBCurYrPersPropTaxAmount
      PRINT QPTrim$(TBPPCUST(1).CustName); P#
      F# = F# + TBPPCUST(1).TBCurYrFarmEquipTaxAmount
      T# = T# + TBPPCUST(1).TBCurYrMachToolsTaxAmount
      M# = M# + TBPPCUST(1).TBCurYrMerchCapTaxAmount
      H# = H# + TBPPCUST(1).TBCurYrMobileHomeTaxAmount
      L# = L# + TBPPCUST(1).TBLateFeeAmount


      BillCust# = BillCust# + 1
     TotalMerchCap# = TotalMerchCap# + TBPPCUST(1).MerchCap
     TotalFarmEquip# = TotalFarmEquip# + TBPPCUST(1).FarmEquip
     TotalMachTools# = TotalMachTools# + TBPPCUST(1).MachTools
     TotalMobileHome# = TotalMobileHome# + TBPPCUST(1).MobileHome

1220 IF lc >= 52 THEN PRINT #10, CHR$(12); : GOSUB 1900
1230 NEXT X

     PRINT #10, CHR$(12); : GOSUB 1900
      Cust = 0
1231 NEXT District
     PRINT #10, ""
     PRINT #10, "Totals Page for Personal Property Proof"
     PRINT #10, ""
     PRINT #10, "Total Customers .. "; USING "######,#"; Customers!
     PRINT #10, "Total Billable ... "; USING "######,#"; BillCust#
     PRINT #10, "Total Personal Property ... "; USING "###########,#"; TotalPersProp#
     PRINT #10, "Total Merchant Capital .... "; USING "###########,#"; TotalMerchCap#
     PRINT #10, "Total Farm Equipment ...... "; USING "###########,#"; TotalFarmEquip#
     PRINT #10, "Total Machinery/Tools ..... "; USING "###########,#"; TotalMachTools#
     PRINT #10, "Total Mobile Homes ........ "; USING "###########,#"; TotalMobileHome#
     Customers! = 0: BillCust# = 0
     TotalPersProp# = 0
     TotalMerchCap# = 0
     TotalFarmEquip# = 0
     TotalMachTools# = 0
     TotalMobileHome# = 0

     PRINT CHR$(12);
1240 CLOSE
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBPREPP.REP"
      Header$ = "Pers Property Book Preliminary Report"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%

      KILL ReportFile$
1250  GOTO 10

1900 P = P + 1
     PRINT #10, TAB(27); "Personal Property Tax Account Listing"
     PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "#####"; P
1910 PRINT #10, ""
     lc = 3
     RETURN


2000 scren$ = "2.02.01"
     CLS
     GOSUB 63500
     COLOR 11: LOCATE 1, 30: PRINT "Print Personal Property Book"
     LOCATE 3, 1: COLOR 11
     PRINT "Enter the Tax Year (YYYY) (ESC-Cancel) .... "; : fl = -4: GOSUB 62000
     IF cf = 1 THEN 10
     IF LEN(iN$) <> 4 THEN 2000
     TaxYear$ = iN$
     COLOR 11
     PRINT "Enter the Following Tax Rates Per $100 Valuation ..."
     PRINT "Personal Property (ESC-Cancel) .. "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     PPTaxRate! = VAL(iN$)
     COLOR 11
     PRINT "Farm Equipment (ESC-Cancel) ..... "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     FETaxRate! = VAL(iN$)
     COLOR 11
     PRINT "Machinery/Tools (ESC-Cancel) .... "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     MTTaxRate! = VAL(iN$)
     COLOR 11
     PRINT "Merchants Capital (ESC-Cancel) .. "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     MCTaxRate! = VAL(iN$)
     COLOR 11
     PRINT "Mobile Homes (ESC-Cancel) ....... "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     MHTaxRate! = VAL(iN$)
     COLOR 11
     PRINT "Enter the Commissioner's Name to Print .... "; : fl = 35: GOSUB 62000
     IF cf = 1 THEN 10
     COMMISSIONER$ = UCASE$(iN$)
     COLOR 11
2099 COLOR 11: LOCATE 13, 1
     PRINT "Enter the Last Name Alphabet to Print (A-Z or 1-2) . ";
     PRINT "Where 1=Town of Kenbridge & 2=Town of Victoria  .... "; : fl = 1
     GOSUB 62000
     IF cf = 1 THEN 10
     IF LEN(iN$) = 0 THEN 2099
     Alpha$ = UCASE$(iN$)
     IF ASC(Alpha$) >= 64 AND ASC(Alpha$) <= 90 THEN GOTO 2099.1
     IF iN$ = "1" OR iN$ = "2" THEN GOTO 2099.1
     GOTO 2099
2099.1 LOCATE 17, 1
     COLOR 11
     PRINT "Enter the Page Number to Start (1-??) ..... "; : fl = -4: GOSUB 62000
     IF cf = 1 THEN 10
     Page! = VAL(iN$)
     COLOR 11
     PRINT "Enter the Beginning Bill Number ........... "; : fl = -7: GOSUB 62000
     BillNumber! = VAL(iN$)
2100 REM for review
     FOR LL = 3 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
     LOCATE 3, 1
     COLOR 15
     PRINT "You Have Entered the Following ..."
     COLOR 11
     PRINT "Commissioner         "; : COLOR 15: PRINT COMMISSIONER$: COLOR 11
     PRINT "Page to Start        "; : COLOR 15: PRINT Page!: COLOR 11
     PRINT "Bill #               "; : COLOR 15: PRINT BillNumber!: COLOR 11
     PRINT "Print Letter         "; : COLOR 15: PRINT Alpha$: COLOR 11
     PRINT
     COLOR 15
     PRINT "Tax Rates ..."
     COLOR 11
     PRINT "Personal Property    "; : COLOR 15: PRINT USING "###.##"; PPTaxRate!: COLOR 11
     PRINT "Farm Equipment       "; : COLOR 15: PRINT USING "###.##"; FETaxRate!: COLOR 11
     PRINT "Machinery/Tools      "; : COLOR 15: PRINT USING "###.##"; MTTaxRate!: COLOR 11
     PRINT "Merchants Capital    "; : COLOR 15: PRINT USING "###.##"; MCTaxRate!: COLOR 11
     PRINT "Mobile Homes         "; : COLOR 15: PRINT USING "###.##"; MHTaxRate!: COLOR 11
     LOCATE 23, 1: COLOR 11
     PRINT "Is the Above Information Correct (Y/N) .. "; : fl = 1
     GOSUB 62210
     IF NOT (iN$ = "Y" OR iN$ = "N") THEN 2100
     IF iN$ = "N" THEN 2000
     Page! = Page! - 1: REM RESET PAGE TO ZERO FOR PROPER PROCESSING
     BillNumber! = BillNumber! - 1
2110 REM Now Print Alignment Test oki 395 must be in 15 character per inch mode
     FOR LL = 3 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
     LOCATE 3, 1: COLOR 15
     PRINT CHR$(7);
     PRINT CHR$(7);
     PRINT "The printer must be set to print in 15 cpi : Please Set Printer Now"
     PRINT "Press <ENTER> when the printer is set to 15 cpi "; : fl = 1
     GOSUB 62000
2111 scren$ = "2.02.01": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 30: PRINT "Print Pers Prop Book"
     LOCATE 3, 1: COLOR 11
     LOCATE 15, 1: COLOR 11
     PRINT "Print an Alignment Mask Now (Y/N) (ESC-Cancel) .. "; : fl = 1
     GOSUB 62210
        IF cf = 1 THEN CLOSE : GOTO 10
        IF NOT (iN$ = "Y" OR iN$ = "N") THEN 2111
        IF iN$ = "Y" THEN GOSUB align: GOTO 2111
     
     scren$ = "2.02.01": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 32: PRINT "Print Pers Prop Book"
     LOCATE 3, 1: COLOR 11

2120 ' Set Constants Here
     
     taxtype$ = "P"
     GOSUB 60000
     PersValue# = 0
     FarmValue# = 0
     MachineValue# = 0
     MerchantValue# = 0
     HomeValue# = 0
     LateAmt# = 0
     TaxAmt# = 0
     PPTRAVal# = 0
     PPTRADis# = 0
    'PPTRA Percent
     IF VAL(TaxYear$) = 1998 THEN Perc! = 12.5
     IF VAL(TaxYear$) = 1999 THEN Perc! = 27.5
     IF VAL(TaxYear$) = 2000 THEN Perc! = 47.5
     IF VAL(TaxYear$) = 2001 THEN Perc! = 70
     IF VAL(TaxYear$) = 2002 THEN Perc! = 100


     


2150 ' Check for No Data
     IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10

     ' Open Report File
     OPEN "O", 10, "TBPPBK.REP": WIDTH #10, 255

     ' Open Index
     OPEN TBpath$ + "TBPPname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
     FIELD 4, 20 AS LastNam$, 4 AS Mrec$: EndofFile = LOF(4) / 24
    
     
     FOR LL = 3 TO 20: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
     LOCATE 3, 1
     COLOR 11
     

2165 GOSUB TopofPage
     PageRecord = 1
     LOCATE 3, 1: COLOR 11
     PRINT "NOW SPOOLING : PLEASE WAIT!!!"
     

2170 FOR X = 1 TO EndofFile
     GET 4, X
     GET TBFILE, CVS(Mrec$), TBPPCUST(1)
     aCCOUNT! = CVS(Mrec$)

2191 IF TBPPCUST(1).Deleted = "Y" THEN 2230               'only active customers to print
     IF Alpha$ >= "A" AND Alpha$ <= "Z" THEN
      IF VAL(TBPPCUST(1).District) > 8 THEN 2230  'only customers in Town of Kenbridge
      IF Alpha$ = "A" AND LEFT$(LastNam$, 1) > "A" THEN 2230
      IF Alpha$ > "A" AND LEFT$(LastNam$, 1) <> Alpha$ THEN 2230
     END IF
     IF Alpha$ = "1" THEN
      IF VAL(TBPPCUST(1).District) <> 9 THEN 2230  'only customers in Town of Kenbridge
     END IF
     IF Alpha$ = "2" THEN
      IF VAL(TBPPCUST(1).District) <> 10 THEN 2230  'only customers in Town of Victoria
     END IF
     
    
     ' First Figure the Tax Amount
     'Calc Pers Property from Cars, Boats, Heavy Equip and Bus Furn
      PersProp# = 0
      PersProp# = TBPPCUST(1).HeavyEquip + TBPPCUST(1).Boats + TBPPCUST(1).BusFurn
      VehRecord! = TBPPCUST(1).FirstVeh
      WHILE VehRecord! <> 0
        GET TBVFile, VehRecord!, TBVeh(1)
        IF TBVeh(1).VehTyp <> "D" THEN
         PersProp# = PersProp# + TBVeh(1).VehValue
         IF TBVeh(1).VehQ = "Y" THEN
          IF TBVeh(1).VehValue > 20000 THEN
           PPTRAVal# = 20000
           ELSE
           PPTRAVal# = TBVeh(1).VehValue
          END IF
          IF PPTRAVal# <= 1000 THEN
           PPTRADis# = PPTRADis# + ((PPTRAVal# / 100) * PPTaxRate!)
           ELSE
           PPTRADis# = PPTRADis# + (((PPTRAVal# / 100) * PPTaxRate!) * (Perc! / 100))
          END IF
           PPTRADis# = INT((PPTRADis# * 100) + .500001) / 100
           PPTRAValue# = PPTRAValue# + PPTRAVal#
          END IF
        END IF
        VehRecord! = TBVeh(1).VehNext
      WEND

      'Set PersProp Value Here
      TBPPCUST(1).PersProp = PersProp#
      PPTaxAmount# = (PersProp# / 100) * PPTaxRate!
      PPTaxAmount# = INT((PPTaxAmount# * 100) + .5) / 100
      FETaxAmount# = (TBPPCUST(1).FarmEquip / 100) * FETaxRate!
      FETaxAmount# = INT((FETaxAmount# * 100) + .5) / 100
      MTTaxAmount# = (TBPPCUST(1).MachTools / 100) * MTTaxRate!
      MTTaxAmount# = INT((MTTaxAmount# * 100) + .5) / 100
      MCTaxAmount# = (TBPPCUST(1).MerchCap / 100) * MCTaxRate!
      MCTaxAmount# = INT((MCTaxAmount# * 100) + .5) / 100
      MHTaxAmount# = (TBPPCUST(1).MobileHome / 100) * MHTaxRate!
      MHTaxAmount# = INT((MHTaxAmount# * 100) + .5) / 100

   


      SubTaxAmount# = PPTaxAmount# + FETaxAmount# + MTTaxAmount# + MCTaxAmount# + MHTaxAmount#

     ' Check for Late Fee
      LateAmount# = 0
      
      IF TBPPCUST(1).LateFile = "Y" OR TBPPCUST(1).LateFile = "T" THEN
       LateAmount# = SubTaxAmount# * .1
       LateAmount# = INT((LateAmount# * 100) + .5) / 100
       IF LateAmount# < 10 THEN LateAmount# = 10
       IF LateAmount# > SubTaxAmount# THEN LateAmount# = SubTaxAmount#
      END IF

     'Total Up The Tax Amount

      TaxAmount# = PPTaxAmount# + FETaxAmount# + MTTaxAmount# + MCTaxAmount# + MHTaxAmount# + LateAmount#
      TaxAmount# = INT((TaxAmount# * 100) + .5) / 100

      IF TaxAmount# = 0 THEN 2230

     ' Set Bill Number
     BillNumber! = BillNumber! + 1



     'Check for Top of Form
     IF AcctsPrinted >= 17 THEN
      GOSUB PageSubtotal
      GOSUB TopofPage
     END IF

     'Print the Account Information

     Address1$ = RTRIM$(TBPPCUST(1).Address1)
     Address2$ = RTRIM$(TBPPCUST(1).Address2)


     PRINT #10, TAB(10); TBPPCUST(1).CustName;
      PRINT #10, TAB(58); USING "#######"; BillNumber!;
      IF TBPPCUST(1).PersProp > 0 THEN
       PRINT #10, TAB(72); USING "########,#"; TBPPCUST(1).PersProp;
      END IF
      IF TBPPCUST(1).FarmEquip > 0 THEN
       PRINT #10, TAB(92); USING "########,#"; TBPPCUST(1).FarmEquip;
      END IF
      IF TBPPCUST(1).MachTools > 0 THEN
       PRINT #10, TAB(112); USING "########,#"; TBPPCUST(1).MachTools;
      END IF
      IF TBPPCUST(1).MerchCap > 0 THEN
       PRINT #10, TAB(132); USING "########,#"; TBPPCUST(1).MerchCap;
      END IF
      IF TBPPCUST(1).MobileHome > 0 THEN
       PRINT #10, TAB(151); USING "########,#"; TBPPCUST(1).MobileHome;
      END IF
      PRINT #10, ""     'Finish the Line Here

     IF LEN(Address1$) > 0 THEN
      PRINT #10, TAB(10); TBPPCUST(1).Address1; TAB(172); USING "######"; CVS(Mrec$)
      PRINT #10, TAB(10); Address2$
      SKIP$ = "F"
     ELSE
      PRINT #10, TAB(10); Address2$; TAB(172); USING "######"; CVS(Mrec$)
      SKIP$ = "T"
     END IF
     
     PRINT #10, TAB(10); RTRIM$(TBPPCUST(1).City); ", "; TBPPCUST(1).State; "  "; TBPPCUST(1).Zip;
     IF SKIP$ = "T" THEN
      PRINT #10, ""
     END IF
     IF PPTaxAmount# > 0 THEN
       PRINT #10, TAB(70); USING "#######,#.##"; PPTaxAmount#;
     END IF
     IF FETaxAmount# > 0 THEN
      PRINT #10, TAB(90); USING "#######,#.##"; FETaxAmount#;
     END IF
     IF MTTaxAmount# > 0 THEN
      PRINT #10, TAB(110); USING "#######,#.##"; MTTaxAmount#;
     END IF
     IF MCTaxAmount# > 0 THEN
      PRINT #10, TAB(130); USING "#######,#.##"; MCTaxAmount#;
     END IF
     IF MHTaxAmount# > 0 THEN
      PRINT #10, TAB(149); USING "#######,#.##"; MHTaxAmount#;
     END IF
     IF LateAmount# > 0 THEN
      PRINT #10, TAB(165); USING "#######,#.##"; LateAmount#;
     END IF
     PRINT #10, TAB(187); USING "#######,#.##"; TaxAmount#

     
     AcctsPrinted = AcctsPrinted + 1

     CustSub = CustSub + 1
     TotalPP# = TotalPP# + TBPPCUST(1).PersProp
     TotalFE# = TotalFE# + TBPPCUST(1).FarmEquip
     TotalMT# = TotalMT# + TBPPCUST(1).MachTools
     TotalMC# = TotalMC# + TBPPCUST(1).MerchCap
     TotalMH# = TotalMH# + TBPPCUST(1).MobileHome
     TotalPPTax# = TotalPPTax# + PPTaxAmount#
     TotalFETax# = TotalFETax# + FETaxAmount#
     TotalMTTax# = TotalMTTax# + MTTaxAmount#
     TotalMCTax# = TotalMCTax# + MCTaxAmount#
     TotalMHTax# = TotalMHTax# + MHTaxAmount#
     TotalLateTax# = TotalLateTax# + LateAmount#
     TotalTaxAmount# = TotalTaxAmount# + TaxAmount#
     TotalPPTRAValue# = TotalPPTRAValue# + PPTRAValue#
     TotalPPTRADis# = TotalPPTRADis# + PPTRADis#



     'Store the TaxAmount and the BillNumber in the Customers File
     
     TBPPCUST(1).LBCurYrBillNumber = LTRIM$(STR$(BillNumber!))
     TBPPCUST(1).LBCurYrPersPropTaxAmount = PPTaxAmount#
     TBPPCUST(1).LBCurYrFarmEquipTaxAmount = FETaxAmount#
     TBPPCUST(1).LBCurYrMachToolsTaxAmount = MTTaxAmount#
     TBPPCUST(1).LBCurYrMerchCapTaxAmount = MCTaxAmount#
     TBPPCUST(1).LBCurYrMobileHomeTaxAmount = MHTaxAmount#
     TBPPCUST(1).LBLateFeeAmount = LateAmount#
     TBPPCUST(1).LBPPTRAValue = PPTRAValue#
     TBPPCUST(1).LBPPTRADiscnt = PPTRADis#
     PUT TBFILE, CVS(Mrec$), TBPPCUST(1)

2230 'Reset PPTRA Calculations
     PPTRAValue# = 0
     PPTRADis# = 0
     NEXT X

     FOR lp = AcctsPrinted + 1 TO 17
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
     NEXT lp
     GOSUB PageSubtotal



2500 CLOSE
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBPPBK.REP"
      Header$ = "Personal Property Book"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
      CLS
      PRINT "Please Note the Following ..."
      PRINT "Your Last Alpha Printed Was ... "; Alpha$
      PRINT "Where 1=Town of Kenbridge & 2=Town of Victoria"
      PRINT
      PRINT
      PRINT "Your Last Bill Number Was : "; BillNumber!
      PRINT "Your Last Printed Page # Was:  "; Page!
      PRINT
      LOCATE 10, 1: PRINT "Press <ENTER> when you have recorded this information "; : fl = 1
      GOSUB 62000
      CLOSE
      GOTO 10





PageSubtotal:
      PRINT #10, ""
      IF TotalPP# > 0 THEN
       PRINT #10, TAB(72); USING "########,#"; TotalPP#;
      END IF
      IF TotalFE# > 0 THEN
       PRINT #10, TAB(92); USING "########,#"; TotalFE#;
      END IF
      IF TotalMT# > 0 THEN
       PRINT #10, TAB(112); USING "########,#"; TotalMT#;
      END IF
      IF TotalMC# > 0 THEN
       PRINT #10, TAB(132); USING "########,#"; TotalMC#;
      END IF
      IF TotalMH# > 0 THEN
       PRINT #10, TAB(151); USING "########,#"; TotalMH#;
      END IF
      PRINT #10, ""     'Finish the Line Here
      PRINT #10, ""
   
     IF TotalPP# > 0 THEN
       PRINT #10, TAB(70); USING "#######,#.##"; TotalPPTax#;
     END IF
     IF TotalFE# > 0 THEN
      PRINT #10, TAB(90); USING "#######,#.##"; TotalFETax#;
     END IF
     IF TotalMT# > 0 THEN
      PRINT #10, TAB(110); USING "#######,#.##"; TotalMTTax#;
     END IF
     IF TotalMC# > 0 THEN
      PRINT #10, TAB(130); USING "#######,#.##"; TotalMCTax#;
     END IF
     IF TotalMH# > 0 THEN
      PRINT #10, TAB(149); USING "#######,#.##"; TotalMHTax#;
     END IF
     IF TotalLateTax# > 0 THEN
      PRINT #10, TAB(165); USING "#######,#.##"; TotalLateTax#;
     END IF
     PRINT #10, TAB(187); USING "#######,#.##"; TotalTaxAmount#

     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, ""


      GOSUB OpenSummaryFile
      
      TBPage(1).Page = Page!
      TBPage(1).AlphaWho = Alpha$
      TBPage(1).PPValue = TotalPP#
      TBPage(1).PPAmount = TotalPPTax#
      TBPage(1).FEValue = TotalFE#
      TBPage(1).FEAmount = TotalFETax#
      TBPage(1).MTValue = TotalMT#
      TBPage(1).MTAmount = TotalMTTax#
      TBPage(1).MCValue = TotalMC#
      TBPage(1).MCAmount = TotalMCTax#
      TBPage(1).MHValue = TotalMH#
      TBPage(1).MHAmount = TotalMHTax#
      TBPage(1).LateAmount = TotalLateTax#
      TBPage(1).TotalTaxAmount = TotalTaxAmount#
     PUT 14, PageRecord, TBPage(1)
     CLOSE 14
     PageRecord = PageRecord + 1
     'Now Clear Subtotals
     TotalPP# = 0
     TotalFE# = 0
     TotalMT# = 0
     TotalMC# = 0
     TotalMH# = 0
     TotalPPTax# = 0
     TotalFETax# = 0
     TotalMTTax# = 0
     TotalMCTax# = 0
     TotalMHTax# = 0
     TotalLateTax# = 0
     TotalTaxAmount# = 0
     RETURN


OpenSummaryFile:
 OPEN TBpath$ + "TBPPSUM" + Alpha$ + ".dat" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = LEN(TBPage(1))
 RETURN

3000 scren$ = "2.03.01"
'Initialize Totals Fields
       PPValue# = 0
       FEValue# = 0
       MTValue# = 0
       MCValue# = 0
       MHValue# = 0
       PPAmount# = 0
       FEAmount# = 0
       MTAmount# = 0
       MCAmount# = 0
       MHAmount# = 0
       LateAmount# = 0
       Page! = 0
       TotalTaxAmount# = 0

     CLS
     GOSUB 63500
     COLOR 11: LOCATE 1, 27: PRINT "Print Personal Property Book Totals"
     LOCATE 3, 1
     PRINT "Enter the Tax Year (YYYY) (ESC-Cancel) .... "; : fl = -4: GOSUB 62000
     IF cf = 1 THEN 10
     IF LEN(iN$) <> 4 THEN 3000
     TaxYear$ = iN$
     COLOR 11
     PRINT "Enter the Following Tax Rates Per $100 Valuation ..."
     PRINT "Personal Property (ESC-Cancel) .. "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     PPTaxRate! = VAL(iN$)
     COLOR 11
     PRINT "Farm Equipment (ESC-Cancel) ..... "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     FETaxRate! = VAL(iN$)
     COLOR 11
     PRINT "Machinery/Tools (ESC-Cancel) .... "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     MTTaxRate! = VAL(iN$)
     COLOR 11
     PRINT "Merchants Capital (ESC-Cancel) .. "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     MCTaxRate! = VAL(iN$)
     COLOR 11
     PRINT "Mobile Homes (ESC-Cancel) ....... "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     MHTaxRate! = VAL(iN$)
     COLOR 11
     PRINT "Enter the Commissioner's Name to Print .... "; : fl = 35: GOSUB 62000
     IF cf = 1 THEN 10
     COMMISSIONER$ = UCASE$(iN$)
3001 LOCATE 12, 1: COLOR 11
     PRINT "Is the Above Information Correct (Y/N) (ESC-Cancel) .. "; : fl = 1
     GOSUB 62210
     IF NOT (iN$ = "Y" OR iN$ = "N") THEN 3001
     IF iN$ = "N" THEN 3000
     FOR LL = 3 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
3002 LOCATE 3, 1: COLOR 11
     PRINT "Select One of the Following ..."
     PRINT
     PRINT "1- Print Totals Page for Lunenburg County"
     PRINT "2- Print Totals Page for Town of Kenbridge"
     PRINT "3- Print Totals Page for Town of Victoria"
     PRINT "4- Print Grand Totals Page"
     PRINT
     PRINT "Selection # (<ESC>-Cancel) .. "; : fl = -1
     GOSUB 62000
     IF cf = 1 THEN 10
     IF VAL(iN$) < 1 OR VAL(iN$) > 4 THEN 3000
     Choice = VAL(iN$)
     'First Set Alignment
     FOR LL = 3 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
     LOCATE 3, 1: COLOR 15
     PRINT CHR$(7);
     PRINT CHR$(7);
     PRINT "The printer must be set to print in 15 cpi : Please Set Printer Now"
     PRINT "Press <ENTER> when the printer is set to 15 cpi "; : fl = 1
     GOSUB 62000
AlignMent1:
     scren$ = "2.02.01": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 30: PRINT "Print Pers Prop Book"
     LOCATE 3, 1: COLOR 11
     LOCATE 15, 1: COLOR 11
     PRINT "Print an Alignment Mask Now (Y/N) (ESC-Cancel) .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN CLOSE : GOTO 10
     IF NOT (iN$ = "Y" OR iN$ = "N") THEN GOTO AlignMent1
     IF iN$ = "Y" THEN GOSUB align: GOTO AlignMent1

    'Proceed to Individual Modules
    
     ON Choice GOTO 3100, 3200, 3300, 3400

3100 REM Print County Totals Page
     OPEN "O", 10, "TBPPTOT.REP": WIDTH #10, 255

     GOSUB 3190
     FOR Cnt = 65 TO 90
     Alpha$ = CHR$(Cnt)
     GOSUB OpenSummaryFile
     IF LOF(14) = 0 THEN 3110
     FOR Record = 1 TO LOF(14) / LEN(TBPage(1))
      GET 14, Record, TBPage(1)
     
     'Check for Top of Form
     IF AcctsPrinted >= 17 THEN
      FOR lll = 1 TO 7: PRINT #10, "": NEXT lll
      GOSUB TopofPage
     END IF
     
     PRINT #10, TAB(12); "Page #"; TBPage(1).Page;
       PRINT #10, TAB(72); USING "########,#"; TBPage(1).PPValue;
       PRINT #10, TAB(92); USING "########,#"; TBPage(1).FEValue;
       PRINT #10, TAB(112); USING "########,#"; TBPage(1).MTValue;
       PRINT #10, TAB(132); USING "########,#"; TBPage(1).MCValue;
       PRINT #10, TAB(151); USING "########,#"; TBPage(1).MHValue
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, TAB(70); USING "#######,#.##"; TBPage(1).PPAmount;
      PRINT #10, TAB(90); USING "#######,#.##"; TBPage(1).FEAmount;
      PRINT #10, TAB(110); USING "#######,#.##"; TBPage(1).MTAmount;
      PRINT #10, TAB(130); USING "#######,#.##"; TBPage(1).MCAmount;
      PRINT #10, TAB(149); USING "#######,#.##"; TBPage(1).MHAmount;
      PRINT #10, TAB(165); USING "#######,#.##"; TBPage(1).LateAmount;
      PRINT #10, TAB(187); USING "#######,#.##"; TBPage(1).TotalTaxAmount

      'Add Them Up
       AcctsPrinted = AcctsPrinted + 1
       PPValue# = PPValue# + TBPage(1).PPValue
       FEValue# = FEValue# + TBPage(1).FEValue
       MTValue# = MTValue# + TBPage(1).MTValue
       MCValue# = MCValue# + TBPage(1).MCValue
       MHValue# = MHValue# + TBPage(1).MHValue

       PPAmount# = PPAmount# + TBPage(1).PPAmount
       FEAmount# = FEAmount# + TBPage(1).FEAmount
       MTAmount# = MTAmount# + TBPage(1).MTAmount
       MCAmount# = MCAmount# + TBPage(1).MCAmount
       MHAmount# = MHAmount# + TBPage(1).MHAmount
       LateAmount# = LateAmount# + TBPage(1).LateAmount
       TotalTaxAmount# = TotalTaxAmount# + TBPage(1).TotalTaxAmount
      NEXT Record
3110 CLOSE 14
     NEXT Cnt

     FOR lp = AcctsPrinted + 1 TO 17
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
     NEXT lp
     FOR lll = 1 TO 7: PRINT #10, "": NEXT lll

     GOSUB TopofPage
     PRINT #10, TAB(12); "Final Totals";
       PRINT #10, TAB(72); USING "########,#"; PPValue#;
       PRINT #10, TAB(92); USING "########,#"; FEValue#;
       PRINT #10, TAB(112); USING "########,#"; MTValue#;
       PRINT #10, TAB(132); USING "########,#"; MCValue#;
       PRINT #10, TAB(151); USING "########,#"; MHValue#
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, TAB(70); USING "#######,#.##"; PPAmount#;
      PRINT #10, TAB(90); USING "#######,#.##"; FEAmount#;
      PRINT #10, TAB(110); USING "#######,#.##"; MTAmount#;
      PRINT #10, TAB(130); USING "#######,#.##"; MCAmount#;
      PRINT #10, TAB(149); USING "#######,#.##"; MHAmount#;
      PRINT #10, TAB(165); USING "#######,#.##"; LateAmount#;
      PRINT #10, TAB(187); USING "#######,#.##"; TotalTaxAmount#
     FOR lp = 2 TO 17
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
     NEXT lp
     FOR lll = 1 TO 7: PRINT #10, "": NEXT lll
     
      CLOSE
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBPPTOT.REP"
      Header$ = "Personal Property Totals Pages"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
     GOTO 3000
     
3190 'Top of Page
        Page! = Page! + 1
        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, TAB(15); "PERSONS AND PROPERTY ASSESSED FOR THE YEAR "; TaxYear$; TAB(190); "PAGE "; Page!
        PRINT #10, TAB(25); "COUNTY"; TAB(65); "LUNENBURG"; TAB(125); COMMISSIONER$
      
PRINT #10,
PRINT #10, TAB(70); "PERSONAL"; TAB(92); "FARM"; TAB(110); "MACHINERY"; TAB(130); "MERCHANT'S"; TAB(149); "MOBILE"; TAB(167); "LATE FILING"; TAB(190); "TOTAL TAX"
PRINT #10, TAB(70); "PROPERTY"; TAB(90); "EQUIPMENT"; TAB(110); " & TOOLS"; TAB(130); " CAPITAL"; TAB(149); "HOMES"; TAB(167); "  PENALTY"
PRINT #10, TAB(75); USING "##.##"; PPTaxRate!;
 PRINT #10, TAB(95); USING "##.##"; FETaxRate!;
 PRINT #10, TAB(113); USING "##.##"; MTTaxRate!;
 PRINT #10, TAB(134); USING "##.##"; MCTaxRate!;
 PRINT #10, TAB(153); USING "##.##"; MHTaxRate!
PRINT #10, ""
AcctsPrinted = 0
RETURN



  
3200 REM Print Town of Kenbridge Totals Page
     OPEN "O", 10, "TBPPTOT.REP": WIDTH #10, 255
     Alpha$ = "1"
     GOSUB OpenSummaryFile
     GOSUB 3290
     FOR Record = 1 TO LOF(14) / LEN(TBPage(1))
      GET 14, Record, TBPage(1)

     'Check for Top of Form
     IF AcctsPrinted >= 17 THEN
      FOR lll = 1 TO 7: PRINT #10, "": NEXT lll
      GOSUB TopofPage
     END IF
     PRINT #10, TAB(12); "Page #"; TBPage(1).Page;
       PRINT #10, TAB(72); USING "########,#"; TBPage(1).PPValue;
       PRINT #10, TAB(92); USING "########,#"; TBPage(1).FEValue;
       PRINT #10, TAB(112); USING "########,#"; TBPage(1).MTValue;
       PRINT #10, TAB(132); USING "########,#"; TBPage(1).MCValue;
       PRINT #10, TAB(151); USING "########,#"; TBPage(1).MHValue
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, TAB(70); USING "#######,#.##"; TBPage(1).PPAmount;
      PRINT #10, TAB(90); USING "#######,#.##"; TBPage(1).FEAmount;
      PRINT #10, TAB(110); USING "#######,#.##"; TBPage(1).MTAmount;
      PRINT #10, TAB(130); USING "#######,#.##"; TBPage(1).MCAmount;
      PRINT #10, TAB(149); USING "#######,#.##"; TBPage(1).MHAmount;
      PRINT #10, TAB(165); USING "#######,#.##"; TBPage(1).LateAmount;
      PRINT #10, TAB(187); USING "#######,#.##"; TBPage(1).TotalTaxAmount

      'Add Them Up
       AcctsPrinted = AcctsPrinted + 1
       PPValue# = PPValue# + TBPage(1).PPValue
       FEValue# = FEValue# + TBPage(1).FEValue
       MTValue# = MTValue# + TBPage(1).MTValue
       MCValue# = MCValue# + TBPage(1).MCValue
       MHValue# = MHValue# + TBPage(1).MHValue

       PPAmount# = PPAmount# + TBPage(1).PPAmount
       FEAmount# = FEAmount# + TBPage(1).FEAmount
       MTAmount# = MTAmount# + TBPage(1).MTAmount
       MCAmount# = MCAmount# + TBPage(1).MCAmount
       MHAmount# = MHAmount# + TBPage(1).MHAmount
       LateAmount# = LateAmount# + TBPage(1).LateAmount
       TotalTaxAmount# = TotalTaxAmount# + TBPage(1).TotalTaxAmount
      NEXT Record
     FOR lp = AcctsPrinted + 1 TO 17
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
     NEXT lp
     FOR lll = 1 TO 7: PRINT #10, "": NEXT lll
3210 CLOSE 14
     GOSUB TopofPage
     PRINT #10, TAB(12); "Final Totals";
       PRINT #10, TAB(72); USING "########,#"; PPValue#;
       PRINT #10, TAB(92); USING "########,#"; FEValue#;
       PRINT #10, TAB(112); USING "########,#"; MTValue#;
       PRINT #10, TAB(132); USING "########,#"; MCValue#;
       PRINT #10, TAB(151); USING "########,#"; MHValue#
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, TAB(70); USING "#######,#.##"; PPAmount#;
      PRINT #10, TAB(90); USING "#######,#.##"; FEAmount#;
      PRINT #10, TAB(110); USING "#######,#.##"; MTAmount#;
      PRINT #10, TAB(130); USING "#######,#.##"; MCAmount#;
      PRINT #10, TAB(149); USING "#######,#.##"; MHAmount#;
      PRINT #10, TAB(165); USING "#######,#.##"; LateAmount#;
      PRINT #10, TAB(187); USING "#######,#.##"; TotalTaxAmount#
     FOR lp = 2 TO 17
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
     NEXT lp
     FOR lll = 1 TO 7: PRINT #10, "": NEXT lll

      CLOSE
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBPPTOT.REP"
      Header$ = "Personal Property Totals Pages"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
     GOTO 3000

3290 'Top of Page
        Page! = Page! + 1
        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, TAB(15); "PERSONS AND PROPERTY ASSESSED FOR THE YEAR "; TaxYear$; TAB(190); "PAGE "; Page!
        PRINT #10, TAB(25); "TOWN"; TAB(65); "KENBRIDGE"; TAB(125); COMMISSIONER$

PRINT #10,
PRINT #10, TAB(70); "PERSONAL"; TAB(92); "FARM"; TAB(110); "MACHINERY"; TAB(130); "MERCHANT'S"; TAB(149); "MOBILE"; TAB(167); "LATE FILING"; TAB(190); "TOTAL TAX"
PRINT #10, TAB(70); "PROPERTY"; TAB(90); "EQUIPMENT"; TAB(110); " & TOOLS"; TAB(130); " CAPITAL"; TAB(149); "HOMES"; TAB(167); "  PENALTY"
PRINT #10, TAB(75); USING "##.##"; PPTaxRate!;
 PRINT #10, TAB(95); USING "##.##"; FETaxRate!;
 PRINT #10, TAB(113); USING "##.##"; MTTaxRate!;
 PRINT #10, TAB(134); USING "##.##"; MCTaxRate!;
 PRINT #10, TAB(153); USING "##.##"; MHTaxRate!
PRINT #10, ""
AcctsPrinted = 0
RETURN

  
3300 REM Print Town of Victoria Totals Page
     OPEN "O", 10, "TBPPTOT.REP": WIDTH #10, 255
     Alpha$ = "2"
     GOSUB OpenSummaryFile
     GOSUB 3390
     FOR Record = 1 TO LOF(14) / LEN(TBPage(1))
      GET 14, Record, TBPage(1)

     'Check for Top of Form
     IF AcctsPrinted >= 17 THEN
      FOR lll = 1 TO 7: PRINT #10, "": NEXT lll
      GOSUB TopofPage
     END IF
     PRINT #10, TAB(12); "Page #"; TBPage(1).Page;
       PRINT #10, TAB(72); USING "########,#"; TBPage(1).PPValue;
       PRINT #10, TAB(92); USING "########,#"; TBPage(1).FEValue;
       PRINT #10, TAB(112); USING "########,#"; TBPage(1).MTValue;
       PRINT #10, TAB(132); USING "########,#"; TBPage(1).MCValue;
       PRINT #10, TAB(151); USING "########,#"; TBPage(1).MHValue
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, TAB(70); USING "#######,#.##"; TBPage(1).PPAmount;
      PRINT #10, TAB(90); USING "#######,#.##"; TBPage(1).FEAmount;
      PRINT #10, TAB(110); USING "#######,#.##"; TBPage(1).MTAmount;
      PRINT #10, TAB(130); USING "#######,#.##"; TBPage(1).MCAmount;
      PRINT #10, TAB(149); USING "#######,#.##"; TBPage(1).MHAmount;
      PRINT #10, TAB(165); USING "#######,#.##"; TBPage(1).LateAmount;
      PRINT #10, TAB(187); USING "#######,#.##"; TBPage(1).TotalTaxAmount

      'Add Them Up
       AcctsPrinted = AcctsPrinted + 1
       PPValue# = PPValue# + TBPage(1).PPValue
       FEValue# = FEValue# + TBPage(1).FEValue
       MTValue# = MTValue# + TBPage(1).MTValue
       MCValue# = MCValue# + TBPage(1).MCValue
       MHValue# = MHValue# + TBPage(1).MHValue

       PPAmount# = PPAmount# + TBPage(1).PPAmount
       FEAmount# = FEAmount# + TBPage(1).FEAmount
       MTAmount# = MTAmount# + TBPage(1).MTAmount
       MCAmount# = MCAmount# + TBPage(1).MCAmount
       MHAmount# = MHAmount# + TBPage(1).MHAmount
       LateAmount# = LateAmount# + TBPage(1).LateAmount
       TotalTaxAmount# = TotalTaxAmount# + TBPage(1).TotalTaxAmount
      NEXT Record
     FOR lp = AcctsPrinted + 1 TO 17
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
     NEXT lp
     FOR lll = 1 TO 7: PRINT #10, "": NEXT lll
3310 CLOSE 14
     GOSUB TopofPage
     PRINT #10, TAB(12); "Final Totals";
       PRINT #10, TAB(72); USING "########,#"; PPValue#;
       PRINT #10, TAB(92); USING "########,#"; FEValue#;
       PRINT #10, TAB(112); USING "########,#"; MTValue#;
       PRINT #10, TAB(132); USING "########,#"; MCValue#;
       PRINT #10, TAB(151); USING "########,#"; MHValue#
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, TAB(70); USING "#######,#.##"; PPAmount#;
      PRINT #10, TAB(90); USING "#######,#.##"; FEAmount#;
      PRINT #10, TAB(110); USING "#######,#.##"; MTAmount#;
      PRINT #10, TAB(130); USING "#######,#.##"; MCAmount#;
      PRINT #10, TAB(149); USING "#######,#.##"; MHAmount#;
      PRINT #10, TAB(165); USING "#######,#.##"; LateAmount#;
      PRINT #10, TAB(187); USING "#######,#.##"; TotalTaxAmount#
     FOR lp = 2 TO 17
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
     NEXT lp
     FOR lll = 1 TO 7: PRINT #10, "": NEXT lll

      CLOSE
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBPPTOT.REP"
      Header$ = "Personal Property Totals Pages"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
     GOTO 3000

3390 'Top of Page
        Page! = Page! + 1
        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, TAB(15); "PERSONS AND PROPERTY ASSESSED FOR THE YEAR "; TaxYear$; TAB(190); "PAGE "; Page!
        PRINT #10, TAB(25); "TOWN"; TAB(65); "VICTORIA"; TAB(125); COMMISSIONER$

PRINT #10,
PRINT #10, TAB(70); "PERSONAL"; TAB(92); "FARM"; TAB(110); "MACHINERY"; TAB(130); "MERCHANT'S"; TAB(149); "MOBILE"; TAB(167); "LATE FILING"; TAB(190); "TOTAL TAX"
PRINT #10, TAB(70); "PROPERTY"; TAB(90); "EQUIPMENT"; TAB(110); " & TOOLS"; TAB(130); " CAPITAL"; TAB(149); "HOMES"; TAB(167); "  PENALTY"
PRINT #10, TAB(75); USING "##.##"; PPTaxRate!;
 PRINT #10, TAB(95); USING "##.##"; FETaxRate!;
 PRINT #10, TAB(113); USING "##.##"; MTTaxRate!;
 PRINT #10, TAB(134); USING "##.##"; MCTaxRate!;
 PRINT #10, TAB(153); USING "##.##"; MHTaxRate!
PRINT #10, ""
AcctsPrinted = 0
RETURN

  
3400 REM Print Grand Totals Page

     OPEN "O", 10, "TBTOTAL.REP": WIDTH #10, 255
    GOSUB 3490
     FOR Cnt = 65 TO 90
     Alpha$ = CHR$(Cnt)
     GOSUB OpenSummaryFile
     IF LOF(14) = 0 THEN 3410
     FOR Record = 1 TO LOF(14) / LEN(TBPage(1))
      GET 14, Record, TBPage(1)


      'Add Them Up
     
       PPValue# = PPValue# + TBPage(1).PPValue
       FEValue# = FEValue# + TBPage(1).FEValue
       MTValue# = MTValue# + TBPage(1).MTValue
       MCValue# = MCValue# + TBPage(1).MCValue
       MHValue# = MHValue# + TBPage(1).MHValue

       PPAmount# = PPAmount# + TBPage(1).PPAmount
       FEAmount# = FEAmount# + TBPage(1).FEAmount
       MTAmount# = MTAmount# + TBPage(1).MTAmount
       MCAmount# = MCAmount# + TBPage(1).MCAmount
       MHAmount# = MHAmount# + TBPage(1).MHAmount
       LateAmount# = LateAmount# + TBPage(1).LateAmount
       TotalTaxAmount# = TotalTaxAmount# + TBPage(1).TotalTaxAmount
      NEXT Record
3410 CLOSE 14
     NEXT Cnt

     PRINT #10, TAB(31); "PERSONAL PROPERTY"; TAB(63); "FARM EQUIPMENT"; TAB(91); "MACHINERY & TOOLS"; TAB(121); "MERCHANTS CAPITAL"; TAB(153); "MOBILE HOMES"; TAB(172); "  PENALTY FOR"; TAB(193); "Total"
     PRINT #10, TAB(10); "DISTRICT"; TAB(29); "VALUE"; TAB(45); "LEVIES"; TAB(59); "VALUE"; TAB(75); "LEVIES"; TAB(89); "VALUE"; TAB(105); "LEVIES"; TAB(119); "VALUE"; TAB(135); "LEVIES"; TAB(149); "VALUE"; TAB(165); "LEVIES"; TAB(173);  _
"FAILURE TO FILE"; TAB(190); "LOCAL LEVIES"



       PRINT #10, TAB(10); "ALL"; TAB(25); USING "########,#"; PPValue#;
        PRINT #10, TAB(40); USING "#######,#.##"; PPAmount#;
        PRINT #10, TAB(55); USING "########,#"; FEValue#;
        PRINT #10, TAB(70); USING "#######,#.##"; FEAmount#;
        PRINT #10, TAB(85); USING "########,#"; MTValue#;
        PRINT #10, TAB(100); USING "#######,#.##"; MTAmount#;
        PRINT #10, TAB(115); USING "########,#"; MCValue#;
        PRINT #10, TAB(130); USING "#######,#.##"; MCAmount#;
        PRINT #10, TAB(145); USING "########,#"; MHValue#;
        PRINT #10, TAB(160); USING "#######,#.##"; MHAmount#;
        PRINT #10, TAB(175); USING "#######,#.##"; LateAmount#;
        PRINT #10, TAB(190); USING "#######,#.##"; TotalTaxAmount#
        
       PRINT #10, ""
       PRINT #10, ""
       PRINT #10, TAB(10); "INCORPORATED TOWNS"
       PRINT #10, ""
     Alpha$ = "1"
     GOSUB OpenSummaryFile
     IF LOF(14) = 0 THEN 3420
     FOR Record = 1 TO LOF(14) / LEN(TBPage(1))
      GET 14, Record, TBPage(1)


      'Add Them Up


       T1PPValue# = T1PPValue# + TBPage(1).PPValue
       T1FEValue# = T1FEValue# + TBPage(1).FEValue
       T1MTValue# = T1MTValue# + TBPage(1).MTValue
       T1MCValue# = T1MCValue# + TBPage(1).MCValue
       T1MHValue# = T1MHValue# + TBPage(1).MHValue

       T1PPAmount# = T1PPAmount# + TBPage(1).PPAmount
       T1FEAmount# = T1FEAmount# + TBPage(1).FEAmount
       T1MTAmount# = T1MTAmount# + TBPage(1).MTAmount
       T1MCAmount# = T1MCAmount# + TBPage(1).MCAmount
       T1MHAmount# = T1MHAmount# + TBPage(1).MHAmount
       T1LateAmount# = T1LateAmount# + TBPage(1).LateAmount
       T1TotalTaxAmount# = T1TotalTaxAmount# + TBPage(1).TotalTaxAmount


      ' ADD UP FOR GRAND TOTALS

       PPValue# = PPValue# + TBPage(1).PPValue
       FEValue# = FEValue# + TBPage(1).FEValue
       MTValue# = MTValue# + TBPage(1).MTValue
       MCValue# = MCValue# + TBPage(1).MCValue
       MHValue# = MHValue# + TBPage(1).MHValue

       PPAmount# = PPAmount# + TBPage(1).PPAmount
       FEAmount# = FEAmount# + TBPage(1).FEAmount
       MTAmount# = MTAmount# + TBPage(1).MTAmount
       MCAmount# = MCAmount# + TBPage(1).MCAmount
       MHAmount# = MHAmount# + TBPage(1).MHAmount
       LateAmount# = LateAmount# + TBPage(1).LateAmount
       TotalTaxAmount# = TotalTaxAmount# + TBPage(1).TotalTaxAmount
      NEXT Record
3420 CLOSE 14

       PRINT #10, TAB(10); "KENBRIDGE"; TAB(25); USING "########,#"; T1PPValue#;
        PRINT #10, TAB(40); USING "#######,#.##"; T1PPAmount#;
        PRINT #10, TAB(55); USING "########,#"; T1FEValue#;
        PRINT #10, TAB(70); USING "#######,#.##"; T1FEAmount#;
        PRINT #10, TAB(85); USING "########,#"; T1MTValue#;
        PRINT #10, TAB(100); USING "#######,#.##"; T1MTAmount#;
        PRINT #10, TAB(115); USING "########,#"; T1MCValue#;
        PRINT #10, TAB(130); USING "#######,#.##"; T1MCAmount#;
        PRINT #10, TAB(145); USING "########,#"; T1MHValue#;
        PRINT #10, TAB(160); USING "#######,#.##"; T1MHAmount#;
        PRINT #10, TAB(175); USING "#######,#.##"; T1LateAmount#;
        PRINT #10, TAB(190); USING "#######,#.##"; T1TotalTaxAmount#
      PRINT #10, ""
     Alpha$ = "2"
     GOSUB OpenSummaryFile
     IF LOF(14) = 0 THEN 3430
     FOR Record = 1 TO LOF(14) / LEN(TBPage(1))
      GET 14, Record, TBPage(1)


      'Add Them Up


       T2PPValue# = T2PPValue# + TBPage(1).PPValue
       T2FEValue# = T2FEValue# + TBPage(1).FEValue
       T2MTValue# = T2MTValue# + TBPage(1).MTValue
       T2MCValue# = T2MCValue# + TBPage(1).MCValue
       T2MHValue# = T2MHValue# + TBPage(1).MHValue

       T2PPAmount# = T2PPAmount# + TBPage(1).PPAmount
       T2FEAmount# = T2FEAmount# + TBPage(1).FEAmount
       T2MTAmount# = T2MTAmount# + TBPage(1).MTAmount
       T2MCAmount# = T2MCAmount# + TBPage(1).MCAmount
       T2MHAmount# = T2MHAmount# + TBPage(1).MHAmount
       T2LateAmount# = T2LateAmount# + TBPage(1).LateAmount
       T2TotalTaxAmount# = T2TotalTaxAmount# + TBPage(1).TotalTaxAmount


      ' ADD UP FOR GRAND TOTALS

       PPValue# = PPValue# + TBPage(1).PPValue
       FEValue# = FEValue# + TBPage(1).FEValue
       MTValue# = MTValue# + TBPage(1).MTValue
       MCValue# = MCValue# + TBPage(1).MCValue
       MHValue# = MHValue# + TBPage(1).MHValue

       PPAmount# = PPAmount# + TBPage(1).PPAmount
       FEAmount# = FEAmount# + TBPage(1).FEAmount
       MTAmount# = MTAmount# + TBPage(1).MTAmount
       MCAmount# = MCAmount# + TBPage(1).MCAmount
       MHAmount# = MHAmount# + TBPage(1).MHAmount
       LateAmount# = LateAmount# + TBPage(1).LateAmount
       TotalTaxAmount# = TotalTaxAmount# + TBPage(1).TotalTaxAmount
      NEXT Record
3430 CLOSE 14
      
       PRINT #10, TAB(10); "VICTORIA"; TAB(25); USING "########,#"; T2PPValue#;
        PRINT #10, TAB(40); USING "#######,#.##"; T2PPAmount#;
        PRINT #10, TAB(55); USING "########,#"; T2FEValue#;
        PRINT #10, TAB(70); USING "#######,#.##"; T2FEAmount#;
        PRINT #10, TAB(85); USING "########,#"; T2MTValue#;
        PRINT #10, TAB(100); USING "#######,#.##"; T2MTAmount#;
        PRINT #10, TAB(115); USING "########,#"; T2MCValue#;
        PRINT #10, TAB(130); USING "#######,#.##"; T2MCAmount#;
        PRINT #10, TAB(145); USING "########,#"; T2MHValue#;
        PRINT #10, TAB(160); USING "#######,#.##"; T2MHAmount#;
        PRINT #10, TAB(175); USING "#######,#.##"; T2LateAmount#;
        PRINT #10, TAB(190); USING "#######,#.##"; T2TotalTaxAmount#
      PRINT #10, ""
      PRINT #10, ""
       PRINT #10, TAB(10); "TOTAL/TOWNS"; TAB(25); USING "########,#"; T1PPValue# + T2PPValue#;
        PRINT #10, TAB(40); USING "#######,#.##"; T1PPAmount# + T2PPAmount#;
        PRINT #10, TAB(55); USING "########,#"; T1FEValue# + T2FEValue#;
        PRINT #10, TAB(70); USING "#######,#.##"; T1FEAmount# + T2FEAmount#;
        PRINT #10, TAB(85); USING "########,#"; T1MTValue# + T2MTValue#;
        PRINT #10, TAB(100); USING "#######,#.##"; T1MTAmount# + T2MTAmount#;
        PRINT #10, TAB(115); USING "########,#"; T1MCValue# + T2MCValue#;
        PRINT #10, TAB(130); USING "#######,#.##"; T1MCAmount# + T2MCAmount#;
        PRINT #10, TAB(145); USING "########,#"; T1MHValue# + T2MHValue#;
        PRINT #10, TAB(160); USING "#######,#.##"; T1MHAmount# + T2MHAmount#;
        PRINT #10, TAB(175); USING "#######,#.##"; T1LateAmount# + T2LateAmount#;
        PRINT #10, TAB(190); USING "#######,#.##"; T1TotalTaxAmount# + T2TotalTaxAmount#
       PRINT #10, ""
       PRINT #10, ""
       PRINT #10, TAB(10); "AGGREGATE"; TAB(25); USING "########,#"; PPValue#;
        PRINT #10, TAB(40); USING "#######,#.##"; PPAmount#;
        PRINT #10, TAB(55); USING "########,#"; FEValue#;
        PRINT #10, TAB(70); USING "#######,#.##"; FEAmount#;
        PRINT #10, TAB(85); USING "########,#"; MTValue#;
        PRINT #10, TAB(100); USING "#######,#.##"; MTAmount#;
        PRINT #10, TAB(115); USING "########,#"; MCValue#;
        PRINT #10, TAB(130); USING "#######,#.##"; MCAmount#;
        PRINT #10, TAB(145); USING "########,#"; MHValue#;
        PRINT #10, TAB(160); USING "#######,#.##"; MHAmount#;
        PRINT #10, TAB(175); USING "#######,#.##"; LateAmount#;
        PRINT #10, TAB(190); USING "#######,#.##"; TotalTaxAmount#
        PRINT #10, ""
       PRINT #10, TAB(10); "TOTAL LOCAL VALUES AND LEVIES"; TAB(60); "   VALUE"; TAB(80); "TOTAL LEVIES"
       PRINT #10,
       PRINT #10, TAB(15); "Personal Property"; TAB(60); USING "########,#"; PPValue#;
        PRINT #10, TAB(80); USING "#######,#.##"; PPAmount#
        PRINT #10, TAB(15); "Farm Equipment"; TAB(60); USING "########,#"; FEValue#;
        PRINT #10, TAB(80); USING "#######,#.##"; FEAmount#
        PRINT #10, TAB(15); "Machinery & Tools"; TAB(60); USING "########,#"; MTValue#;
        PRINT #10, TAB(80); USING "#######,#.##"; MTAmount#
        PRINT #10, TAB(15); "Merchant Capital"; TAB(60); USING "########,#"; MCValue#;
        PRINT #10, TAB(80); USING "#######,#.##"; MCAmount#
        PRINT #10, TAB(15); "Mobile Homes"; TAB(60); USING "########,#"; MHValue#;
        PRINT #10, TAB(80); USING "#######,#.##"; MHAmount#
        PRINT #10, TAB(15); "Failure to File Penalty"; TAB(80); USING "#######,#.##"; LateAmount#
        PRINT #10, TAB(15); "Aggregate"; TAB(60); USING "########,#"; PPValue# + FEValue# + MTValue# + MCValue# + MHValue#;
        PRINT #10, TAB(80); USING "#######,#.##"; TotalTaxAmount#



        'Last Page Thank You Lord

        FOR LL = 1 TO 4
        PRINT #10, ""
        NEXT LL
        PRINT #10, TAB(20); "I, _____________________________ THE COMMISSIONER OF REVENUE FOR THE COUNTY OF __________________________, VIRGINIA"
        PRINT #10, TAB(20); "DO SOLOMNLY SWEAR AND AFFIRM THAT IN MAKING THIS BOOK I HAVE TO THE BEST OF MY KNOWLEDGE"
        PRINT #10, TAB(20); "AND ABILITY, COMPLIED WITH THE LAWS PRESCRIBING THE DUTIES OF COMMISSIONER OF THE REVENUE."
        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, TAB(20); "GIVEN UNDER MY HAND THIS ____________ DAY OF ______________, 19____"
        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, TAB(20); "_____________________________________"
        PRINT #10, TAB(20); "COMMISSIONER OF THE REVENUE"
        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, TAB(20); "SWORN TO BEFORE ME __________________________________, A NOTORY PUBLIC"
        PRINT #10, ""
        PRINT #10, TAB(20); "(OR OTHER PERSON AUTHORIZED TO ADMINISTER OATHS) FOR THE COUNTY OF __________________________."
        PRINT #10, ""
        PRINT #10, TAB(20); "ON THE ______________________ DAY OF ___________________________, 1997"
        PRINT #10,
        PRINT #10,
        PRINT #10,
        PRINT #10, TAB(60); "__________________________________________________"
        PRINT #10, TAB(60); "(PERSON AUTHORIZE TO ADMINISTER OATH)"
        PRINT #10, ""

        PRINT #10, CHR$(12);


      CLOSE
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBTOTAL.REP"
      Header$ = "Personal Property Totals Pages"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
     GOTO 3000

    
3490 'Top of Page
        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, TAB(15); "SUMMARY OF ALL PERSONS AND PROPERTY ASSESSED FOR TAXATION FOR YEAR "; TaxYear$
        PRINT #10, TAB(25); "COUNTY"; TAB(65); "LUNENBURG"; TAB(125); COMMISSIONER$
        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, ""
        RETURN

     

     ' SUBS *************************************************

TopofPage:
        Page! = Page! + 1
        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, TAB(15); "PERSONS AND PROPERTY ASSESSED FOR THE YEAR "; TaxYear$; TAB(190); "PAGE "; Page!
     IF Alpha$ >= "A" AND Alpha$ <= "Z" THEN
       PRINT #10, TAB(25); "COUNTY"; TAB(65); "LUNENBURG"; TAB(125); COMMISSIONER$
      END IF
     IF Alpha$ = "1" THEN
      PRINT #10, TAB(25); "TOWN"; TAB(65); "KENBRIDGE"; TAB(125); COMMISSIONER$
     END IF
     IF Alpha$ = "2" THEN
       PRINT #10, TAB(25); "TOWN"; TAB(65); "VICTORIA"; TAB(125); COMMISSIONER$  'only customers in Town of Victoria
     END IF
PRINT #10,
PRINT #10, TAB(70); "PERSONAL"; TAB(92); "FARM"; TAB(110); "MACHINERY"; TAB(130); "MERCHANT'S"; TAB(149); "MOBILE"; TAB(167); "LATE FILING"; TAB(190); "TOTAL TAX"
PRINT #10, TAB(70); "PROPERTY"; TAB(90); "EQUIPMENT"; TAB(110); " & TOOLS"; TAB(130); " CAPITAL"; TAB(149); "HOMES"; TAB(167); "  PENALTY"
PRINT #10, TAB(75); USING "##.##"; PPTaxRate!;
 PRINT #10, TAB(95); USING "##.##"; FETaxRate!;
 PRINT #10, TAB(113); USING "##.##"; MTTaxRate!;
 PRINT #10, TAB(134); USING "##.##"; MCTaxRate!;
 PRINT #10, TAB(153); USING "##.##"; MHTaxRate!
PRINT #10, ""
AcctsPrinted = 0
RETURN

align:                  ' alignment mask for land book forms
ReportFile$ = "TBALGN.REP"
OPEN "O", 10, "TBALGN.REP": WIDTH #10, 255
PRINT #10, ""
PRINT #10, ""
PRINT #10, TAB(10); "PERSONS AND PROPERTY ASSESSED FOR THE YEAR "; TaxYear$; TAB(190); "PAGE #####"
PRINT #10, TAB(25); "COUNTY"; TAB(65); "LUNENBURG"; TAB(125); COMMISSIONER$
PRINT #10,
PRINT #10, TAB(70); "PERSONAL"; TAB(92); "FARM"; TAB(110); "MACHINERY"; TAB(130); "MERCHANTS"; TAB(149); "MOBILE"; TAB(167); "LATE FILING"; TAB(190); "TOTAL TAX"
PRINT #10, TAB(70); "PROPERTY"; TAB(90); "EQUIPMENT"; TAB(110); " & TOOLS"; TAB(130); " CAPITAL"; TAB(149); "HOMES"; TAB(167); "  PENALTY"
PRINT #10, TAB(75); USING "##.##"; PPTaxRate!;
 PRINT #10, TAB(95); USING "##.##"; FETaxRate!;
 PRINT #10, TAB(113); USING "##.##"; MTTaxRate!;
 PRINT #10, TAB(134); USING "##.##"; MCTaxRate!;
 PRINT #10, TAB(153); USING "##.##"; MHTaxRate!
PRINT #10, ""

PRINT #10, TAB(10); "CUSTOMER NAME"; TAB(58); "XXXXXXX";
  PRINT #10, TAB(70); "XX,XXX,XXX";
  PRINT #10, TAB(90); "XX,XXX,XXX";
  PRINT #10, TAB(110); "XX,XXX,XXX";
  PRINT #10, TAB(130); "XX,XXX,XXX";
  PRINT #10, TAB(150); "XX,XXX,XXX"
PRINT #10, TAB(10); "ADDRESS LINE 1"; TAB(170); "XXXXX"
PRINT #10, TAB(10); "ADDRESS LINE 2";
PRINT #10, TAB(10); "CITY, STATE & ZIP";
  PRINT #10, TAB(70); "XXX,XXX.XX";
  PRINT #10, TAB(90); "XXX,XXX.XX";
  PRINT #10, TAB(110); "XXX,XXX.XX";
  PRINT #10, TAB(130); "XXX,XXX.XX";
  PRINT #10, TAB(150); "XXX,XXX.XX";
  PRINT #10, TAB(187); "XXX,XXX.XX"
PRINT #10, TAB(10); "CUSTOMER NAME"; TAB(58); "XXXXXXX";
  PRINT #10, TAB(70); "XX,XXX,XXX";
  PRINT #10, TAB(90); "XX,XXX,XXX";
  PRINT #10, TAB(110); "XX,XXX,XXX";
  PRINT #10, TAB(130); "XX,XXX,XXX";
  PRINT #10, TAB(150); "XX,XXX,XXX"
PRINT #10, TAB(10); "ADDRESS LINE 1"; TAB(170); "XXXXX"
PRINT #10, TAB(10); "ADDRESS LINE 2";
PRINT #10, TAB(10); "CITY, STATE & ZIP";
  PRINT #10, TAB(70); "XXX,XXX.XX";
  PRINT #10, TAB(90); "XXX,XXX.XX";
  PRINT #10, TAB(110); "XXX,XXX.XX";
  PRINT #10, TAB(130); "XXX,XXX.XX";
  PRINT #10, TAB(150); "XXX,XXX.XX";
  PRINT #10, TAB(187); "XXX,XXX.XX"
PRINT #10, TAB(10); "CUSTOMER NAME";
  PRINT #10, TAB(58); "XXXXXXX";
  PRINT #10, TAB(70); "XX,XXX,XXX";
  PRINT #10, TAB(90); "XX,XXX,XXX";
  PRINT #10, TAB(110); "XX,XXX,XXX";
  PRINT #10, TAB(130); "XX,XXX,XXX";
  PRINT #10, TAB(150); "XX,XXX,XXX"
PRINT #10, TAB(10); "ADDRESS LINE 1"; TAB(170); "XXXXX"
PRINT #10, TAB(10); "ADDRESS LINE 2";
PRINT #10, TAB(10); "CITY, STATE & ZIP";
  PRINT #10, TAB(70); "XXX,XXX.XX";
  PRINT #10, TAB(90); "XXX,XXX.XX";
  PRINT #10, TAB(110); "XXX,XXX.XX";
  PRINT #10, TAB(130); "XXX,XXX.XX";
  PRINT #10, TAB(150); "XXX,XXX.XX";
  PRINT #10, TAB(187); "XXX,XXX.XX"
FOR LCNT! = 22 TO 78: PRINT #10, "": NEXT LCNT!

PRINT #10, TAB(70); "XX,XXX,XXX";
  PRINT #10, TAB(90); "XX,XXX,XXX";
  PRINT #10, TAB(110); "XX,XXX,XXX";
  PRINT #10, TAB(130); "XX,XXX,XXX";
  PRINT #10, TAB(150); "XX,XXX,XXX"
PRINT #10, ""
PRINT #10, TAB(70); "XXX,XXX.XX";
  PRINT #10, TAB(90); "XXX,XXX.XX";
  PRINT #10, TAB(110); "XXX,XXX.XX";
  PRINT #10, TAB(130); "XXX,XXX.XX";
  PRINT #10, TAB(150); "XXX,XXX.XX";
  PRINT #10, TAB(170); "XXX,XXX.XX";
  PRINT #10, TAB(187); "XXX,XXX.XX"
PRINT #10, ""
PRINT #10, ""
PRINT #10, ""

      CLOSE 10

      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBALGN.REP"
      Header$ = "Pers Prop Book Preliminary Report"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
      RETURN


4000 scren$ = "2.04.00": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 25: PRINT "Print Pers Property Book - Towns"
     LOCATE 3, 1
     PRINT "Select From (1)-Kenbridge (2)-Victoria (ESC-End) .. "; : fl = -1
     GOSUB 62000
     IF cf = 1 THEN 10
     Town = VAL(iN$)
     IF Town < 1 OR Town > 2 THEN 4000
     IF Town = 1 THEN DistrictName$ = "Town of Kenbridge": District = 9
     IF Town = 2 THEN DistrictName$ = "Town of Victoria": District = 10
     LOCATE 4, 1: COLOR 11
     PRINT "Enter the Tax Year (YYYY) (ESC-Cancel) .... "; : fl = -4: GOSUB 62000
     IF cf = 1 THEN 10
     IF LEN(iN$) <> 4 THEN 4000
     TaxYear$ = iN$
     COLOR 11
     PRINT "Enter the Following Tax Rates Per $100 Valuation For "; DistrictName$
     PRINT "Personal Property (ESC-Cancel) .. "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     PPTaxRate! = VAL(iN$)
     COLOR 11
     PRINT "Farm Equipment (ESC-Cancel) ..... "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     FETaxRate! = VAL(iN$)
     COLOR 11
     PRINT "Machinery/Tools (ESC-Cancel) .... "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     MTTaxRate! = VAL(iN$)
     COLOR 11
     PRINT "Merchants Capital (ESC-Cancel) .. "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     MCTaxRate! = VAL(iN$)
     COLOR 11
     PRINT "Mobile Homes (ESC-Cancel) ....... "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     MHTaxRate! = VAL(iN$)
     BillNumber! = VAL(RIGHT$(TaxYear$, 2)) * 10000            'Default Bill Number to Begin Here
     taxtype$ = "P"

    'PPTRA Percent
     IF VAL(TaxYear$) = 1998 THEN Perc! = 12.5
     IF VAL(TaxYear$) = 1999 THEN Perc! = 27.5
     IF VAL(TaxYear$) = 2000 THEN Perc! = 47.5
     IF VAL(TaxYear$) = 2001 THEN Perc! = 70
     IF VAL(TaxYear$) = 2002 THEN Perc! = 100



     GOSUB 60000
4150 IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10
     OPEN "O", 10, "TBPREPPT.REP": WIDTH #10, 255
     GOSUB 4900
4160 IF taxtype$ = "R" THEN
        OPEN TBpath$ + "TBREname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS LastNam$, 4 AS Mrec$: EndofFile = LOF(4) / 24
        ELSE
        OPEN TBpath$ + "TBPPname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS LastNam$, 4 AS Mrec$: EndofFile = LOF(4) / 24
      END IF
      
4170 FOR X = 1 TO EndofFile
     GET 4, X

     GET 1, CVS(Mrec$), TBPPCUST(1)
     aCCOUNT! = CVS(Mrec$)


4191 IF TBPPCUST(1).Deleted = "Y" THEN 4230
     IF VAL(TBPPCUST(1).District) <> District THEN 4230

     ' First Figure the Tax Amount
     'Calc Pers Property from Cars, Boats, Heavy Equip and Bus Furn
      PersProp# = 0
      PersProp# = TBPPCUST(1).HeavyEquip + TBPPCUST(1).Boats + TBPPCUST(1).BusFurn
      VehRecord! = TBPPCUST(1).FirstVeh
      WHILE VehRecord! <> 0
        GET TBVFile, VehRecord!, TBVeh(1)
        IF TBVeh(1).VehTyp <> "D" THEN
         PersProp# = PersProp# + TBVeh(1).VehValue
         IF TBVeh(1).VehQ = "Y" THEN
          IF TBVeh(1).VehValue > 20000 THEN
           PPTRAVal# = 20000
           ELSE
           PPTRAVal# = TBVeh(1).VehValue
          END IF
          IF PPTRAVal# <= 1000 THEN
           PPTRADis# = PPTRADis# + ((PPTRAVal# / 100) * PPTaxRate!)
           ELSE
           PPTRADis# = PPTRADis# + (((PPTRAVal# / 100) * PPTaxRate!) * (Perc! / 100))
          END IF
           PPTRADis# = INT((PPTRADis# * 100) + .50001) / 100
           PPTRAValue# = PPTRAValue# + PPTRAVal#
          END IF

        END IF
        VehRecord! = TBVeh(1).VehNext
      WEND

      'Set PersProp Value Here
      TBPPCUST(1).PersProp = PersProp#
      PPTaxAmount# = (PersProp# / 100) * PPTaxRate!
      PPTaxAmount# = INT((PPTaxAmount# * 100) + .5) / 100
      FETaxAmount# = (TBPPCUST(1).FarmEquip / 100) * FETaxRate!
      FETaxAmount# = INT((FETaxAmount# * 100) + .5) / 100
      MTTaxAmount# = (TBPPCUST(1).MachTools / 100) * MTTaxRate!
      MTTaxAmount# = INT((MTTaxAmount# * 100) + .5) / 100
      MCTaxAmount# = (TBPPCUST(1).MerchCap / 100) * MCTaxRate!
      MCTaxAmount# = INT((MCTaxAmount# * 100) + .5) / 100
      MHTaxAmount# = (TBPPCUST(1).MobileHome / 100) * MHTaxRate!
      MHTaxAmount# = INT((MHTaxAmount# * 100) + .5) / 100

      IF District = 9 THEN
       FETaxAmount# = 0
       MTTaxAmount# = 0
       MCTaxAmount# = 0
      END IF
      IF District = 10 THEN
       FETaxAmount# = 0
       MCTaxAmount# = 0
      END IF


      SubTaxAmount# = PPTaxAmount# + FETaxAmount# + MTTaxAmount# + MCTaxAmount# + MHTaxAmount#

     ' Check for Late Fee
       LateAmount# = 0                    ' No Late Amount Due For Individual Towns


     'Total Up The Tax Amount

      TaxAmount# = PPTaxAmount# + FETaxAmount# + MTTaxAmount# + MCTaxAmount# + MHTaxAmount# + LateAmount#
      TaxAmount# = INT((TaxAmount# * 100) + .5) / 100

      IF TaxAmount# = 0 THEN
      'Set to Nothing
       TBPPCUST(1).TBCurYrBillNumber = ""
       TBPPCUST(1).TBCurYrPersPropTaxAmount = 0
       TBPPCUST(1).TBCurYrFarmEquipTaxAmount = 0
       TBPPCUST(1).TBCurYrMachToolsTaxAmount = 0
       TBPPCUST(1).TBCurYrMerchCapTaxAmount = 0
       TBPPCUST(1).TBCurYrMobileHomeTaxAmount = 0
       TBPPCUST(1).TBLateFeeAmount = 0
       TBPPCUST(1).TBPPTRAValue = 0
       TBPPCUST(1).TBPPTRADiscnt = 0

       PUT TBFILE, CVS(Mrec$), TBPPCUST(1)
       GOTO 4230
      END IF

     ' Set Bill Number
     BillNumber! = BillNumber! + 1


     'Print the Account Information

     Address1$ = RTRIM$(TBPPCUST(1).Address1)
     Address2$ = RTRIM$(TBPPCUST(1).Address2)

     PRINT #10, USING "######"; aCCOUNT!;
      PRINT #10, TAB(10); TBPPCUST(1).CustName; TAB(50); USING "#######"; BillNumber!;
      PRINT #10, TAB(65); USING "#######,#.##"; TaxAmount#
     PRINT #10, USING "#######,#.##"; PPTaxAmount#;
      PRINT #10, TAB(15); USING "#######,#.##"; FETaxAmount#;
      PRINT #10, TAB(30); USING "#######,#.##"; MTTaxAmount#;
      PRINT #10, TAB(45); USING "#######,#.##"; MCTaxAmount#;
      PRINT #10, TAB(60); USING "#######,#.##"; MHTaxAmount#;
      PRINT #10, TAB(75); USING "#######,#.##"; LateAmount#

     IF District = 9 THEN
     PRINT #10, USING "#######,#.##"; TBPPCUST(1).PersProp;
      PRINT #10, TAB(60); USING "#######,#.##"; TBPPCUST(1).MobileHome
     END IF
     IF District = 10 THEN
     PRINT #10, USING "#######,#.##"; TBPPCUST(1).PersProp;
      PRINT #10, TAB(30); USING "#######,#.##"; TBPPCUST(1).MachTools;
      PRINT #10, TAB(60); USING "#######,#.##"; TBPPCUST(1).MobileHome
     END IF
     PRINT #10, STRING$(90, "-")
     lc = lc + 4

     AcctsPrinted = AcctsPrinted + 1

     CustSub = CustSub + 1
     IF District = 9 THEN
     TotalPP# = TotalPP# + TBPPCUST(1).PersProp
     TotalMH# = TotalMH# + TBPPCUST(1).MobileHome
     END IF
     IF District = 10 THEN
      TotalPP# = TotalPP# + TBPPCUST(1).PersProp
      TotalMT# = TotalMT# + TBPPCUST(1).MachTools
      TotalMH# = TotalMH# + TBPPCUST(1).MobileHome
     END IF

     TotalPPTax# = TotalPPTax# + PPTaxAmount#
     TotalFETax# = TotalFETax# + FETaxAmount#
     TotalMTTax# = TotalMTTax# + MTTaxAmount#
     TotalMCTax# = TotalMCTax# + MCTaxAmount#
     TotalMHTax# = TotalMHTax# + MHTaxAmount#
     TotalLateTax# = TotalLateTax# + LateAmount#
     TotalTaxAmount# = TotalTaxAmount# + TaxAmount#


   ' Store the TaxAmount and the BillNumber in the Customers File
     TBPPCUST(1).TBCurYrBillNumber = LTRIM$(STR$(BillNumber!))
     TBPPCUST(1).TBCurYrPersPropTaxAmount = PPTaxAmount#
     TBPPCUST(1).TBCurYrFarmEquipTaxAmount = FETaxAmount#
     TBPPCUST(1).TBCurYrMachToolsTaxAmount = MTTaxAmount#
     TBPPCUST(1).TBCurYrMerchCapTaxAmount = MCTaxAmount#
     TBPPCUST(1).TBCurYrMobileHomeTaxAmount = MHTaxAmount#
     TBPPCUST(1).TBLateFeeAmount = LateAmount#
     TBPPCUST(1).TBPPTRAValue = PPTRAValue#
     TBPPCUST(1).TBPPTRADiscnt = PPTRADis#

     PUT TBFILE, CVS(Mrec$), TBPPCUST(1)

     IF lc >= 55 THEN
      PRINT #10, CHR$(12);
      GOSUB 4900
     END IF
4230 PPTRAValue# = 0
     PPTRADis# = 0
     NEXT X
     PRINT #10, CHR$(12); : GOSUB 4900
     Cust = 0
     PRINT #10, ""
     PRINT #10, "Totals Page for Personal Property Tax Listing"
     PRINT #10, ""
     PRINT #10, "Total Billable Customers... "; USING "######,#"; AcctsPrinted
     PRINT #10, "Total Personal Property ... "; USING "###########,#"; TotalPP#
     PRINT #10, "Total Merchant Capital .... "; USING "###########,#"; TotalMC#
     PRINT #10, "Total Farm Equipment ...... "; USING "###########,#"; TotalFE#
     PRINT #10, "Total Machinery/Tools ..... "; USING "###########,#"; TotalMT#
     PRINT #10, "Total Mobile Homes ........ "; USING "###########,#"; TotalMH#
     PRINT #10, ""
     PRINT #10, "Taxed Amounts ..."
     PRINT #10, "Total Personal Property Tax .. "; USING "$$#######,#.##"; TotalPPTax#
     PRINT #10, "Total Farm Equip Tax ......... "; USING "$$#######,#.##"; TotalFETax#
     PRINT #10, "Total Machinery/Tools Tax .... "; USING "$$#######,#.##"; TotalMTTax#
     PRINT #10, "Total Merchants Cap Tax ...... "; USING "$$#######,#.##"; TotalMCTax#
     PRINT #10, "Total Mobile Home Tax ........ "; USING "$$#######,#.##"; TotalMHTax#
     PRINT #10, "Total Late Amount ............ "; USING "$$#######,#.##"; TotalLateTax#
     PRINT #10, "Total Tax Billed ............. "; USING "$$#######,#.##"; TotalTaxAmount#
     TotalPPTax# = 0
     TotalFETax# = 0
     TotalMTTax# = 0
     TotalMCTax# = 0
     TotalMHTax# = 0
     TotalLateTax# = 0
     TotalTaxAmount# = 0
     Customers! = 0
     TotalPersProp# = 0
     TotalMerchCap# = 0
     TotalFarmEquip# = 0
     TotalMachTools# = 0
     TotalMobileHome# = 0
     AcctsPrinted = 0
     PRINT CHR$(12);
4240 CLOSE
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBPREPPT.REP"
      Header$ = "Pers Property Town Book Report"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
4250  GOTO 10

4900 P = P + 1
     PRINT #10, TAB(22); "Personal Property Tax Listing : Town of "; DistrictName$
     PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "#####"; P
     PRINT #10, ""
4910 PRINT #10, "Acct #"; TAB(10); "Customer Name"; TAB(50); "Bill Number"; TAB(65); "* Total Tax *"
     PRINT #10, "   Pers Prop"; TAB(15); " Farm Equip"; TAB(30); "  Mach/Tools"; TAB(45); "   Merch Cap"; TAB(60); " Mobile Home"; TAB(75); "   Late Amt"
     PRINT #10, STRING$(90, "-")
     lc = 5
     RETURN




60000  IF taxtype$ = "R" THEN
         TBFILE = FREEFILE
         OPEN TBpath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFILE LEN = LEN(TBRECust(1))
        ELSE
         TBFILE = FREEFILE
         OPEN TBpath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFILE LEN = LEN(TBPPCUST(1))
         CLOSE TBVFile
         TBVFile = FREEFILE
         OPEN TBpath$ + "TBVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))

         END IF
        RETURN

60025 TBBalFile = FREEFILE
      IF taxtype$ = "R" THEN
      OPEN TBpath$ + "TBREBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBpath$ + "TBPPBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
      END IF
60030 RETURN

61000 TBTransFile = FREEFILE
      OPEN TBpath$ + "TBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTransFile LEN = LEN(TBTrans(1))
61020 RETURN

62000 COLOR 14: cf = 0: iN$ = "": RO = CSRLIN: PO = POS(XY): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
      IF LEN(z$) < 2 THEN 62065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN cf = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN cf = 12: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 61 THEN cf = 13: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 62 THEN cf = 14: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 63 THEN cf = 15: RETURN
62065 IF ASC(z$) = 0 THEN 62030

62070 IF LEN(z$) > 1 THEN cf = ASC(RIGHT$(z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
62080 IF z$ = " " THEN 62130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 62000
62090 IF z$ <> CHR$(8) THEN 62130 ELSE IF ZL = 0 THEN 62030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 62110 ELSE IF z$ = "," THEN 62120
62100 IF z$ = "." THEN ZD = 0: GOTO 62110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
62110 iN$ = LEFT$(iN$, LEN(iN$) - 1)
62120 ZL = ZL - 1: GOTO 62030
62130 IF ABS(fl) = ZL THEN 62030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 62180
62140 IF z$ = "." AND INSTR(iN$, ".") = 0 THEN GOTO 62180
62150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF z$ < "0" OR z$ > "9" THEN 62030
62180 PRINT z$; : iN$ = iN$ + z$: ZL = ZL + 1
62190 IF ABS(fl) = 1 THEN 62050 ELSE 62030
62200 LOCATE RO, PO: PRINT iN$: RETURN
62210 GOSUB 62000: IF LEN(iN$) THEN iN$ = CHR$(ASC(iN$) AND 223)
62220 RETURN
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SYSDATE$; TAB(70); scren$: PRINT STRING$(79, "_"): RETURN


FigureBalance:
     CurBal# = 0: PastDue# = 0
     CurBal# = TBPPCUST(1).CurYrPersPropTax + TBPPCUST(1).CurYrFarmEquipTax + TBPPCUST(1).CurYrMachToolsTax + TBPPCUST(1).CurYrMerchCapTax + TBPPCUST(1).CurYrMobileHomeTax + TBPPCUST(1).CurYrLateFeeTax + TBPPCUST(1).CurYrIntAmount + TBPPCUST(1). _
CurYrPenAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR LL = 1 TO 10
     PastDue# = PastDue# + TBPPCustBal(1).PerPropTaxAmt(LL) + TBPPCustBal(1).FarmEquipAmt(LL) + TBPPCustBal(1).MachToolsAmt(LL) + TBPPCustBal(1).MerchCapAmt(LL) + TBPPCustBal(1).MobileHomeAmt(LL) + TBPPCustBal(1).LateFeeAmt(LL) + TBPPCustBal(1). _
IntAmt(LL) + TBPPCustBal(1).PenAmt(LL)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT LL
     RETURN

Deleted:
    LOCATE 23, 1: COLOR 15
    PRINT "This Account Has Been Deleted! Press <ENTER> to Continue"; :
    fl = 1
    GOSUB 62000
    LOCATE 23, 1: PRINT STRING$(79, 32)
    RETURN

networkerror:
     REM ERROR CODES
     Number = ERR
     SELECT CASE ERR
          CASE IS = 24, 25
               LOCATE 23, 1: PRINT STRING$(79, 32): LOCATE 23, 1: COLOR 28
               PRINT "Error Message:"; : COLOR 15: PRINT "Device Time-out! Make Sure Printer Is ON-LINE:Press ENTER to Resume Printing";
               fl = 1: GOSUB 62000: LOCATE 23, 1: PRINT STRING$(79, 32)
               RESUME
          CASE IS = 27
               LOCATE 23, 1: PRINT STRING$(79, 32): LOCATE 23, 1: COLOR 28
               PRINT "Error Message:"; : COLOR 15: PRINT "Printer Is OUT OF PAPER! Load Paper : Press ENTER to Resume Printing";
               fl = 1: GOSUB 62000: LOCATE 23, 1: PRINT STRING$(79, 32)
               RESUME

          CASE ELSE
               CLS
               LOCATE 15, 1: COLOR 28: PRINT "Error Message: "; : COLOR 15: PRINT "Error "; Number; " Has Occured in Line "; ERL
               PRINT "Write Down Message and Hit <ESC> Key to exit."
               fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10 ELSE GOSUB 62000
            
     END SELECT

