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
   'TBpath$ = "c:\qb45\lc\"                    ' change to "F:\" Before compiling
   CLOSE
   OPEN "R", 1, TBpath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
   CurrentYear$ = year$: CLOSE 1
   ON ERROR GOTO networkerror


     REDIM TransLink!(500)
     REDIM LandValue#(99)
     REDIM BldgValue#(99)
     REDIM TaxAmt#(99)
     REDIM Customer!(99)
    
     REDIM TBRECust(1) AS TBRECust
     REDIM TBRECustBal(1) AS TBREBAL
     REDIM TBPPCust(1) AS TBPPCust
     REDIM TBPPCustBal(1) AS TBPPBal
     REDIM TBTrans(1) AS TBTrans
     REDIM TBPage(1) AS TBPage
     
     WIDTH LPRINT 236
     
100 REM begin program
110 scren$ = "2.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500

1000 REM
     
     
1010 scren$ = "2.01.00": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 32: PRINT "Print Land Book"
1020 LOCATE 5, 1
     PRINT TAB(20); " 1- Preliminary Land Book (on greenbar)"
     PRINT
     PRINT TAB(20); " 2- Final Land Book (on Forms)"
     PRINT
     PRINT TAB(20); " 3- Re-Print Land Book District Total's Pages"
     PRINT
     PRINT TAB(20); " 4- Town Land Book (Victoria/Kenbridge)"
     PRINT
     PRINT TAB(20); "    Selection # (1-4) or (ESC-End) .. "; : fl = -1
     GOSUB 62000
     IF cf = 1 THEN RUN "tbccustr"
     IF VAL(in$) < 1 OR VAL(in$) > 4 THEN 10
     ON VAL(in$) GOTO 1100, 2000, 3000, 4000
 

1100 scren$ = "2.01.01": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 32: PRINT "Print Land Book - For Review"
     LOCATE 3, 1
     taxtype$ = "R"
     CustCount = 0
     GOSUB 60000
1150 IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10
     LOCATE 3, 1: COLOR 11
     PRINT "Enter the District to Begin Printing (ENTER=ALL) (ESC-End) .. "; : fl = 2
     GOSUB 62000
     IF cf = 1 THEN RUN "tbccustr"
     DistrictToStart = VAL(in$)
     IF DistrictToStart < 1 THEN DistrictToStart = 1
     IF DistrictToStart > 10 THEN CLOSE : RUN


     OPEN "O", 10, "TBPRELND.REP": WIDTH #10, 255
     GOSUB 1900
1160 IF taxtype$ = "R" THEN
        OPEN TBpath$ + "TBREname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS MREC$: EndofFile = LOF(4) / 24
        ELSE
        OPEN TBpath$ + "TBPPname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS MREC$: EndofFile = LOF(4) / 24
      END IF
     

1165 FOR District = DistrictToStart TO 10                     'MAIN LOOP
     LOCATE 5, 1: COLOR 11
     PRINT "NOW SPOOLING OUT DISTRICT # "; District
     IF District = 8 THEN 1231
1170 FOR x = 1 TO EndofFile
     GET 4, x
     GET 1, CVS(MREC$), TBRECust(1)
     Account! = CVS(MREC$)
     
     
1191 IF TBRECust(1).Deleted = "Y" THEN 1230
     IF VAL(TBRECust(1).District) <> District THEN 1230
     PRINT #10, "Acct # "; Account!; TAB(20); "Map Numb:"; TBRECust(1).MapNumber; TAB(50); "     Class: "; TBRECust(1).ClassCode; TAB(90); "Prop Desc"; TAB(110); "REC # "; CVS(MREC$)
     PRINT #10, TBRECust(1).CustName; TAB(50); "Land Value: "; USING "$$#######,#"; TBRECust(1).LandValue;
      PRINT #10, TAB(90); TBRECust(1).PropDesc1
     PRINT #10, TBRECust(1).Address1; TAB(50); "Bldg Value: "; USING "$$#######,#"; TBRECust(1).BldgValue;
      PRINT #10, TAB(90); TBRECust(1).PropDesc2
     PRINT #10, TBRECust(1).Address2; TAB(50); "      Size: "; USING "#####.####"; VAL(TBRECust(1).Acres);
      PRINT #10, TAB(90); TBRECust(1).PropDesc3
     PRINT #10, RTRIM$(TBRECust(1).Address3); TAB(50); "  District: "; TBRECust(1).District;
      PRINT #10, TAB(90); TBRECust(1).PropDesc4
     lc = lc + 6
     PRINT #10, STRING$(120, "-")
     CustSub = CustSub + 1
     TotalLandSub# = TotalLandSub# + TBRECust(1).LandValue
     TotalLandSub# = INT((TotalLandSub# * 100) + .5) / 100
     TotalBldgSub# = TotalBldgSub# + TBRECust(1).BldgValue
     TotalBldgSub# = INT((TotalBldgSub# * 100) + .5) / 100


     IF LEFT$(TBRECust(1).ClassCode, 1) = "7" THEN
      ExCust = ExCust + 1
      ExLand# = ExLand# + TBRECust(1).LandValue
      ExLand# = INT((ExLand# * 100) + .5) / 100
      ExBldg# = ExBldg# + TBRECust(1).BldgValue
      ExBldg# = INT((ExBldg# * 100) + .5) / 100
     END IF

     CustCount = CustCount + 1
     TotalLand# = TotalLand# + TBRECust(1).LandValue
     TotalLand# = INT((TotalLand# * 100) + .5) / 100
     TotalBldg# = TotalBldg# + TBRECust(1).BldgValue
     TotalBldg# = INT((TotalBldg# * 100) + .5) / 100

     IF LEFT$(TBRECust(1).ClassCode, 1) = "7" THEN
      TotalExCust = TotalExCust + 1
      TotalExLand# = TotalExLand# + TBRECust(1).LandValue
      TotalExLand# = INT((TotalExLand# * 100) + .5) / 100
      TotalExBldg# = TotalExBldg# + TBRECust(1).BldgValue
     END IF




1220 IF lc >= 58 THEN PRINT #10, CHR$(12); : GOSUB 1900
1230 NEXT x
     PRINT #10, ""
     PRINT #10, "District # "; District; "  Subtotals"
     PRINT #10, "Customers = "; CustSub; TAB(40); "Exempted Customers = "; ExCust
     PRINT #10, "Total Land Value: "; USING "$$########,#"; TotalLandSub#;
      PRINT #10, TAB(40); " Ex Land Value: "; USING "$$#######,#"; ExLand#;
       PRINT #10, TAB(80); "Net Land Value: "; USING "$$#######,#"; TotalLandSub# - ExLand#
     PRINT #10, "Total Bldg Value: "; USING "$$########,#"; TotalBldgSub#;
      PRINT #10, TAB(40); " Ex Bldg Value: "; USING "$$#######,#"; ExBldg#;
       PRINT #10, TAB(80); "Net Bldg Value: "; USING "$$#######,#"; TotalBldgSub# - ExBldg#
     PRINT #10, "Total Valuation : "; USING "$$########,#"; TotalLandSub# + TotalBldgSub#;
      PRINT #10, TAB(40); "Ex Total Value: "; USING "$$#######,#"; ExLand# + ExBldg#;
       PRINT #10, TAB(80); "   Total Value: "; USING "$$#######,#"; (TotalLandSub# + TotalBldgSub#) - (ExLand# + ExBldg#)
     PRINT #10, CHR$(12); : GOSUB 1900
      CustSub = 0: TotalLandSub# = 0: TotalBldgSub# = 0
      ExCust = 0
      ExLand# = 0
      ExBldg# = 0
1231 NEXT District
     PRINT #10, "Total Customers .. "; USING "######,#"; CustCount;
      PRINT #10, TAB(40); "Exempted Customers = "; TotalExCust
     PRINT #10, "Total Land Value: "; USING "$$#########,#"; TotalLand#;
      PRINT #10, TAB(40); " Ex Land Value: "; USING "$$#########,#"; TotalExLand#;
      PRINT #10, TAB(80); "Net Land Value: "; USING "$$#########,#"; TotalLand# - TotalExLand#
     PRINT #10, "Total Bldg Value: "; USING "$$#########,#"; TotalBldg#;
      PRINT #10, TAB(40); " Ex Bldg Value: "; USING "$$#########,#"; TotalExBldg#;
      PRINT #10, TAB(80); "Net Bldg Value: "; USING "$$#########,#"; TotalBldg# - TotalExBldg#
     PRINT #10, "Total Valuation : "; USING "$$#########,#"; TotalLand# + TotalBldg#;
      PRINT #10, TAB(40); "Ex Total Value: "; USING "$$#########,#"; TotalExLand# + TotalExBldg#;
      PRINT #10, TAB(80); " Net Valuation: "; USING "$$#########,#"; (TotalLand# + TotalBldg#) - (TotalExLand# + TotalExBldg#)
     PRINT CHR$(12);
1240 CLOSE
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBPRELND.REP"
      Header$ = "Land Book Preliminary Report"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
1250 GOTO 10

1900 P = P + 1
     PRINT #10, TAB(27); "Real Estate Tax Account Listing": PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "#####"; P
1910 PRINT #10, "Acct. #"
     PRINT #10, STRING$(120, "="): lc = 3: RETURN



2000 scren$ = "2.02.01": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 32: PRINT "Print Land Book"
     LOCATE 3, 1: COLOR 11
     PRINT "Enter the Tax Year (YYYY) (ESC-Cancel) .... "; : fl = -4: GOSUB 62000
     IF cf = 1 THEN 10
     IF LEN(in$) <> 4 THEN 2000
     TaxYear$ = in$
     COLOR 11
     PRINT "Enter the Tax Rate Per $100 (ESC-Cancel) .. "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     TaxRate! = VAL(in$)
     COLOR 11
     PRINT "Enter the Commissioner's Name to Print .... "; : fl = 35: GOSUB 62000
     IF cf = 1 THEN 10
     Commissioner$ = UCASE$(in$)
2099 COLOR 11: LOCATE 7, 1
     PRINT "Enter the District Number to Print (1-12) . ";
     PRINT "Where 11=Tax Exempt & 12=Grand Totals ..... "; : fl = -2: GOSUB 62000
     IF cf = 1 THEN 10
     DistricttoPrint = VAL(in$)
     IF VAL(in$) < 1 OR VAL(in$) > 12 THEN 2099
     PRINT ""
     COLOR 11
     PRINT "Enter the Page Number to Start (1-??) ..... "; : fl = -4: GOSUB 62000
     IF cf = 1 THEN 10
     Page! = VAL(in$)
     COLOR 11
     PRINT "Enter the Beginning Bill Number ........... "; : fl = -7: GOSUB 62000
     BillNumber! = VAL(in$)
2100 REM for review
     FOR LL = 3 TO 12: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
     LOCATE 3, 1
     COLOR 11
     PRINT "You Have Entered the Following ..."
     PRINT "Tax Year is    "; : COLOR 15: PRINT TaxYear$: COLOR 11
     PRINT "Tax Rate is    "; : COLOR 15: PRINT USING "$##.##"; TaxRate!: COLOR 11
     PRINT "Commissioner   "; : COLOR 15: PRINT Commissioner$: COLOR 11
     PRINT "Print District "; : COLOR 15: PRINT DistricttoPrint: COLOR 11
     PRINT "Page to Start  "; : COLOR 15: PRINT Page!: COLOR 11
     PRINT "Bill #         "; : COLOR 15: PRINT BillNumber!: COLOR 11
     PRINT
     PRINT "Is the Above Information Correct (Y/N) .. "; : fl = 1
     GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN 2100
     IF in$ = "N" THEN 2000
     Page! = Page! - 1: REM RESET PAGE TO ZERO FOR PROPER PROCESSING
     BillNumber! = BillNumber! - 1
2110 REM Now Print Alignment Test oki 395 must be in 15 character per inch mode
'Initialize if DistricttoPrint=1
IF DistricttoPrint = 1 THEN
FOR District = 1 TO 12
District$ = STR$(District)
District$ = LTRIM$(District$)
OPEN TBpath$ + "TBDIST" + District$ + ".dat" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = LEN(TBPage(1))
CLOSE 14
nn$ = TBpath$ + "TBDIST" + District$ + ".dat"
KILL nn$
OPEN TBpath$ + "TBEDIS" + District$ + ".dat" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = LEN(TBPage(1))
CLOSE 14
nn$ = TBpath$ + "TBEDIS" + District$ + ".dat"
KILL nn$
NEXT District
END IF


     LOCATE 12, 1: COLOR 11
     PRINT "The printer must be set to print in 15 cpi : Please Set Printer Now"
     PRINT "Press <ENTER> when the printer is set to 15 cpi "; : fl = 1
     GOSUB 62000
2111 scren$ = "2.02.01": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 32: PRINT "Print Land Book"
     LOCATE 3, 1: COLOR 11
     LOCATE 15, 1: COLOR 11
     PRINT "Print an Alignment Mask Now (Y/N) (ESC-Cancel) .. "; : fl = 1
     GOSUB 62210
        IF cf = 1 THEN CLOSE : GOTO 10
        IF NOT (in$ = "Y" OR in$ = "N") THEN 2111
        IF in$ = "Y" THEN GOSUB align: GOTO 2111
     
     scren$ = "2.02.01": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 32: PRINT "Print Land Book"
     LOCATE 3, 1: COLOR 11

2120 ' Set Constants Here
     
     taxtype$ = "R"
     GOSUB 60000
     LandValue# = 0
     BldgValue# = 0
     TaxAmt# = 0
     ELandTotal# = 0
     EBldgTotal# = 0
     ETaxTotal# = 0
     ECustomers# = 0

     GLandTotal# = 0
     GBldgTotal# = 0
     GTaxTotal# = 0
     GCustomers# = 0


2150 ' Check for No Data
     IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10

     ' Open Report File
     OPEN "O", 10, "TBLANDBK.REP": WIDTH #10, 255

     ' Open Index
     OPEN TBpath$ + "TBREname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
     FIELD 4, 20 AS lastnam$, 4 AS MREC$: EndofFile = LOF(4) / 24
                 
     ' Clear Screen for Displaying District Print Information
     FOR LL = 3 TO 20: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
     LOCATE 3, 1
     COLOR 11
     

2165 ' MAIN LOOP NON EXEMPT PROPERTY
     IF DistricttoPrint > 10 THEN 2265
     District = DistricttoPrint
     GOSUB getdistrict
     GOSUB TopOfPage
     PageRecord = 1
     LOCATE 3, 1: COLOR 11
     PRINT "NOW SPOOLING OUT DISTRICT # "; District; "  "; DistrictName$
     

2170 FOR x = 1 TO EndofFile             'Sub Loop within district
     GET 4, x
     GET 1, CVS(MREC$), TBRECust(1)
     Account! = CVS(MREC$)

2191 IF TBRECust(1).Deleted = "Y" THEN 2230               'only active customers to print
     IF VAL(TBRECust(1).District) <> District THEN 2230   'only customers in this district
     IF LEFT$(TBRECust(1).ClassCode, 1) = "7" THEN 2230   'only taxable property first
     IF VAL(TBRECust(1).District) <> DistricttoPrint THEN 2230
     


     ' First Figure the Tax Amount
     TotalValue# = TBRECust(1).LandValue + TBRECust(1).BldgValue

     IF TotalValue# = 0 THEN 2230

     TaxAmount# = (TotalValue# / 100) * TaxRate!
     TaxAmount# = INT((TaxAmount# * 100) + .5) / 100

     
     ' Set Bill Number
     BillNumber! = BillNumber! + 1

     
     'Check for Top of Form
     IF AcctsPrinted >= 17 THEN
      GOSUB PageSubtotal
      GOSUB TopOfPage
     END IF

     'Print the Account Information

     PRINT #10, TAB(15); TBRECust(1).CustName;
      PRINT #10, TAB(60); TBRECust(1).PropDesc1;
      PRINT #10, TAB(104); USING "########,#"; TBRECust(1).LandValue;
      PRINT #10, TAB(119); USING "########,#"; TBRECust(1).BldgValue;
      PRINT #10, TAB(133); USING "########,#"; TotalValue#;
      PRINT #10, TAB(159); USING "#######,#.##"; TaxAmount#;
      PRINT #10, TAB(180); TBRECust(1).MapNumber;
      PRINT #10, TAB(198); USING "#######"; BillNumber!
     PRINT #10, TAB(15); TBRECust(1).Address1;
      PRINT #10, TAB(60); TBRECust(1).PropDesc2;
      PRINT #10, TAB(180); TBRECust(1).Comment1
     PRINT #10, TAB(15); TBRECust(1).Address2;
      PRINT #10, TAB(60); TBRECust(1).PropDesc3;
      PRINT #10, TAB(154); TBRECust(1).ClassCode;
      PRINT #10, TAB(180); TBRECust(1).Comment2
     PRINT #10, TAB(15); RTRIM$(TBRECust(1).Address3);
      PRINT #10, TAB(60); TBRECust(1).PropDesc4;
      PRINT #10, TAB(104); USING "########,#"; TBRECust(1).LandValue;
      PRINT #10, TAB(119); USING "########,#"; TBRECust(1).BldgValue;
      PRINT #10, TAB(133); USING "########,#"; TotalValue#;
      PRINT #10, TAB(159); USING "#######,#.##"; TaxAmount#;
      PRINT #10, TAB(180); USING "######"; CVS(MREC$)


     AcctsPrinted = AcctsPrinted + 1

     CustSub = CustSub + 1
     TotalLandSub# = TotalLandSub# + TBRECust(1).LandValue
     TotalLandSub# = INT((TotalLandSub# * 100) + .5) / 100
     TotalBldgSub# = TotalBldgSub# + TBRECust(1).BldgValue
     TotalBldgSub# = INT((TotalBldgSub# * 100) + .5) / 100
     TotalTaxAmount# = TotalTaxAmount# + TaxAmount#

     CustCount = CustCount + 1
     TotalLand# = TotalLand# + TBRECust(1).LandValue
     TotalLand# = INT((TotalLand# * 100) + .5) / 100
     TotalBldg# = TotalBldg# + TBRECust(1).BldgValue
     TotalBldg# = INT((TotalBldg# * 100) + .5) / 100

     Class = VAL(LEFT$(TBRECust(1).ClassCode, 1))

     TotalValue#(Class) = TotalValue#(Class) + TotalValue#
     TotalValue#(Class) = INT((TotalValue#(Class) * 100) + .5) / 100

     'Store the TaxAmount and the BillNumber in the Customers File

     TBRECust(1).LBCurYrBillNumber = LTRIM$(STR$(BillNumber!))
     TBRECust(1).LBCurYrTaxAmount = TaxAmount#
     PUT 1, CVS(MREC$), TBRECust(1)
     
2230 NEXT x

     FOR lp = AcctsPrinted + 1 TO 17
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
     NEXT lp
     GOSUB PageSubtotal

     ' Summary Pages To Print Here

    GOSUB OpenSummaryFile

    NumberofPages = LOF(14) / LEN(TBPage(1))

    GOSUB TopOfPage
    PagesPrinted = 0

    FOR Cnt = 1 TO NumberofPages
    GET 14, Cnt, TBPage(1)
    
      IF PagesPrinted >= 17 THEN
       FOR lll = 1 TO 8
        PRINT #10, ""
       NEXT lll
        PRINT #10, ""
       GOSUB TopOfPage
       PagesPrinted = 0
      END IF

      PRINT #10, TAB(16); "1- "; USING "########,#"; TBPage(1).Class1;
       PRINT #10, TAB(37); "5- "; USING "########,#"; TBPage(1).Class5;
       PRINT #10, TAB(65); "9- "; USING "########,#"; TBPage(1).Class9;
       PRINT #10, TAB(105); USING "########,#"; TBPage(1).LandValue;
       PRINT #10, TAB(120); USING "########,#"; TBPage(1).BldgValue;
       PRINT #10, TAB(135); USING "#########,#"; TBPage(1).LandValue + TBPage(1).BldgValue;
       PRINT #10, TAB(160); USING "########,#.##"; TBPage(1).TaxValue;
       PRINT #10, TAB(185); USING "####"; TBPage(1).Page
      PRINT #10, TAB(16); "2- "; USING "########,#"; TBPage(1).Class2;
       PRINT #10, TAB(37); "6- "; USING "########,#"; TBPage(1).Class6;
       PRINT #10, TAB(65); " - "; USING "########,#"; TBPage(1).Class0
      PRINT #10, TAB(16); "3- "; USING "########,#"; TBPage(1).Class3;
       PRINT #10, TAB(37); "7- "; USING "########,#"; TBPage(1).Class7
      PRINT #10, TAB(16); "4- "; USING "########,#"; TBPage(1).Class4;
       PRINT #10, TAB(37); "8- "; USING "########,#"; TBPage(1).Class8;
       PRINT #10, TAB(105); USING "########,#"; TBPage(1).LandValue;
       PRINT #10, TAB(120); USING "########,#"; TBPage(1).BldgValue;
       PRINT #10, TAB(135); USING "#########,#"; TBPage(1).LandValue + TBPage(1).BldgValue;
       PRINT #10, TAB(160); USING "########,#.##"; TBPage(1).TaxValue

       ' Subtotal Here
      
      LandValue# = LandValue# + TBPage(1).LandValue
      BldgValue# = BldgValue# + TBPage(1).BldgValue
      TaxAmt# = TaxAmt# + TBPage(1).TaxValue
      PagesPrinted = PagesPrinted + 1

     NEXT Cnt
     CLOSE 14

     FOR lll = PagesPrinted + 1 TO 17
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
     NEXT lll
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, TAB(104); USING "########,#"; LandValue#;
      PRINT #10, TAB(119); USING "########,#"; BldgValue#;
      PRINT #10, TAB(133); USING "########,#"; LandValue# + BldgValue#;
      PRINT #10, TAB(159); USING "#######,#.##"; TaxAmt#
    

     GOTO 2500



2265 ' MAIN LOOP EXEMPT PROPERTY
     IF DistricttoPrint = 11 THEN
     EXEMPTFLAG$ = "Y"
     FOR District = 1 TO 10
     IF District = 8 THEN 2331       'NO DISTRICT 8

     GOSUB getdistrict
     GOSUB TopOfPage

     PageRecord = 1

     LOCATE 3, 1: COLOR 11: PRINT STRING$(79, 32): LOCATE 3, 1
     PRINT "NOW SPOOLING OUT (TAX EXEMPT) DISTRICT # "; District; "  "; DistrictName$
     

2270 FOR x = 1 TO EndofFile             'Sub Loop within district
     GET 4, x
     GET 1, CVS(MREC$), TBRECust(1)
     Account! = CVS(MREC$)

2291 IF TBRECust(1).Deleted = "Y" THEN 2330               'only active customers to print
     IF VAL(TBRECust(1).District) <> District THEN 2330   'only customers in this district
     IF LEFT$(TBRECust(1).ClassCode, 1) <> "7" THEN 2330  'only taxable property first


     ' First Figure the Tax Amount
     TotalValue# = TBRECust(1).LandValue + TBRECust(1).BldgValue
     TaxAmount# = (TotalValue# / 100) * TaxRate!
     TaxAmount# = INT((TaxAmount# * 100) + .5) / 100

     IF TotalValue# = 0 THEN 2330               'No Zero Amounts Allowed

     ' Set Bill Number
     BillNumber! = BillNumber! + 1


     'Check for Top of Form
     IF AcctsPrinted >= 17 THEN
      GOSUB PageSubtotal
      GOSUB TopOfPage
     END IF

     'Print the Account Information

     PRINT #10, TAB(15); TBRECust(1).CustName;
      PRINT #10, TAB(60); TBRECust(1).PropDesc1;
      PRINT #10, TAB(104); USING "########,#"; TBRECust(1).LandValue;
      PRINT #10, TAB(119); USING "########,#"; TBRECust(1).BldgValue;
      PRINT #10, TAB(133); USING "########,#"; TotalValue#;
      PRINT #10, TAB(159); USING "#######,#.##"; TaxAmount#;
      PRINT #10, TAB(180); TBRECust(1).MapNumber;
      PRINT #10, TAB(198); USING "#######"; BillNumber!
     PRINT #10, TAB(15); TBRECust(1).Address1;
      PRINT #10, TAB(60); TBRECust(1).PropDesc2;
      PRINT #10, TAB(180); TBRECust(1).Comment1
     PRINT #10, TAB(15); TBRECust(1).Address2;
      PRINT #10, TAB(60); TBRECust(1).PropDesc3;
      PRINT #10, TAB(154); TBRECust(1).ClassCode;
      PRINT #10, TAB(180); TBRECust(1).Comment2
     PRINT #10, TAB(15); RTRIM$(TBRECust(1).Address3);
      PRINT #10, TAB(60); TBRECust(1).PropDesc4;
      PRINT #10, TAB(104); USING "########,#"; TBRECust(1).LandValue;
      PRINT #10, TAB(119); USING "########,#"; TBRECust(1).BldgValue;
      PRINT #10, TAB(133); USING "########,#"; TotalValue#;
      PRINT #10, TAB(159); USING "#######,#.##"; TaxAmount#;
      PRINT #10, TAB(180); USING "######"; CVS(MREC$)


     AcctsPrinted = AcctsPrinted + 1

     CustSub = CustSub + 1
     TotalLandSub# = TotalLandSub# + TBRECust(1).LandValue
     TotalLandSub# = INT((TotalLandSub# * 100) + .5) / 100
     TotalBldgSub# = TotalBldgSub# + TBRECust(1).BldgValue
     TotalBldgSub# = INT((TotalBldgSub# * 100) + .5) / 100
     TotalTaxAmount# = TotalTaxAmount# + TaxAmount#

     CustCount = CustCount + 1
     TotalLand# = TotalLand# + TBRECust(1).LandValue
     TotalLand# = INT((TotalLand# * 100) + .5) / 100
     TotalBldg# = TotalBldg# + TBRECust(1).BldgValue
     TotalBldg# = INT((TotalBldg# * 100) + .5) / 100

     Class = VAL(LEFT$(TBRECust(1).ClassCode, 1))

     TotalValue#(Class) = TotalValue#(Class) + TotalValue#
     TotalValue#(Class) = INT((TotalValue#(Class) * 100) + .5) / 100

     'Store the TaxAmount and the BillNumber in the Customers File

     TBRECust(1).LBCurYrBillNumber = LTRIM$(STR$(BillNumber!))
     TBRECust(1).LBCurYrTaxAmount = TaxAmount#
     PUT 1, CVS(MREC$), TBRECust(1)

2330 NEXT x

     FOR lp = AcctsPrinted + 1 TO 17
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
     NEXT lp
     GOSUB PageSubtotal

     ' Summary Pages To Print Here

    GOSUB OpenExemptSummaryFile

    NumberofPages = LOF(14) / LEN(TBPage(1))

    GOSUB TopOfPage
    PagesPrinted = 0

    FOR Cnt = 1 TO NumberofPages
    GET 14, Cnt, TBPage(1)

      IF PagesPrinted >= 17 THEN
       FOR lll = 1 TO 8
        PRINT #10, ""
       NEXT lll
        PRINT #10, ""
       GOSUB TopOfPage
       PagesPrinted = 0
      END IF

      PRINT #10, TAB(16); "1- "; USING "########,#"; TBPage(1).Class1;
       PRINT #10, TAB(37); "5- "; USING "########,#"; TBPage(1).Class5;
       PRINT #10, TAB(65); "9- "; USING "########,#"; TBPage(1).Class9;
       PRINT #10, TAB(105); USING "########,#"; TBPage(1).LandValue;
       PRINT #10, TAB(120); USING "########,#"; TBPage(1).BldgValue;
       PRINT #10, TAB(135); USING "#########,#"; TBPage(1).LandValue + TBPage(1).BldgValue;
       PRINT #10, TAB(160); USING "########,#.##"; TBPage(1).TaxValue;
       PRINT #10, TAB(185); USING "####"; TBPage(1).Page
      PRINT #10, TAB(16); "2- "; USING "########,#"; TBPage(1).Class2;
       PRINT #10, TAB(37); "6- "; USING "########,#"; TBPage(1).Class6;
       PRINT #10, TAB(65); " - "; USING "########,#"; TBPage(1).Class0
      PRINT #10, TAB(16); "3- "; USING "########,#"; TBPage(1).Class3;
       PRINT #10, TAB(37); "7- "; USING "########,#"; TBPage(1).Class7
      PRINT #10, TAB(16); "4- "; USING "########,#"; TBPage(1).Class4;
       PRINT #10, TAB(37); "8- "; USING "########,#"; TBPage(1).Class8;
       PRINT #10, TAB(105); USING "########,#"; TBPage(1).LandValue;
       PRINT #10, TAB(120); USING "########,#"; TBPage(1).BldgValue;
       PRINT #10, TAB(135); USING "#########,#"; TBPage(1).LandValue + TBPage(1).BldgValue;
       PRINT #10, TAB(160); USING "########,#.##"; TBPage(1).TaxValue

       ' Subtotal Here

      LandValue# = LandValue# + TBPage(1).LandValue
      BldgValue# = BldgValue# + TBPage(1).BldgValue
      TaxAmt# = TaxAmt# + TBPage(1).TaxValue
      PagesPrinted = PagesPrinted + 1

     NEXT Cnt
     CLOSE 14

     FOR lll = PagesPrinted + 1 TO 17
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
     NEXT lll
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, TAB(104); USING "########,#"; LandValue#;
      PRINT #10, TAB(119); USING "########,#"; BldgValue#;
      PRINT #10, TAB(133); USING "########,#"; LandValue# + BldgValue#;
      PRINT #10, TAB(159); USING "#######,#.##"; TaxAmt#
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, ""
     LandValue# = 0
     BldgValue# = 0
     TaxAmt# = 0

2331 NEXT District
     END IF

     IF DistricttoPrint = 12 THEN
2400 ' Grand Totals Page Goes Here
     ' 84 lines per page total
     ' First Page is Non-Exempt

        Page! = Page! + 1

        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, TAB(100); TaxYear$; TAB(155); USING "$##.##"; TaxRate!;
          PRINT #10, TAB(183); "$100"
        PRINT #10, TAB(20); "IN LUNENBURG COUNTY"; TAB(80); "NON-EXEMPT"; TAB(130); Commissioner$; TAB(188); Page!
        PRINT #10, ""
        PRINT #10, TAB(15); STRING$(187, "#")
        PRINT #10, TAB(15); STRING$(187, "#")
        PRINT #10, TAB(104); "LAND VALUE"; TAB(119); "BLDG VALUE"; TAB(136); "TOTAL VALUE"; TAB(162); "TOTAL TAX"; TAB(187); "COUNT"

          District = 0
     WHILE District < 10
        District = District + 1

        IF District = 8 THEN District = 9       'No District 8

        DistricttoPrint = District
        
        GOSUB getdistrict
        GOSUB OpenSummaryFile
     FOR Cnt = 1 TO LOF(14) / LEN(TBPage(1))
     GET 14, Cnt, TBPage(1)
       Customers# = Customers# + TBPage(1).Customers
       LandTotal# = LandTotal# + TBPage(1).LandValue
       BldgTotal# = BldgTotal# + TBPage(1).BldgValue
       TaxTotal# = TaxTotal# + TBPage(1).TaxValue
     NEXT Cnt
     CLOSE 14
     PRINT #10, TAB(15); DistrictName$;
      PRINT #10, TAB(104); USING "########,#"; LandTotal#;
      PRINT #10, TAB(119); USING "########,#"; BldgTotal#;
      PRINT #10, TAB(137); USING "########,#"; LandTotal# + BldgTotal#;
      PRINT #10, TAB(185); USING "#####,#"; Customers#
     PRINT #10, ""
     PRINT #10, TAB(104); USING "########,#"; LandTotal#;
      PRINT #10, TAB(119); USING "########,#"; BldgTotal#;
      PRINT #10, TAB(137); USING "########,#"; LandTotal# + BldgTotal#;
      PRINT #10, TAB(159); USING "#######,#.##"; TaxTotal#
     PRINT #10, ""
     GLandTotal# = GLandTotal# + LandTotal#
     GBldgTotal# = GBldgTotal# + BldgTotal#
     GTaxTotal# = GTaxTotal# + TaxTotal#
     GCustomers# = GCustomers# + Customers#
     LandTotal# = 0: BldgTotal# = 0: TaxTotal# = 0: Customers# = 0
     WEND
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, ""
     
     PRINT #10, TAB(15); "NON-EXEMPT TOTALS"; TAB(103); USING "#########,#"; GLandTotal#;
      PRINT #10, TAB(118); USING "#########,#"; GBldgTotal#;
      PRINT #10, TAB(136); USING "#########,#"; GLandTotal# + GBldgTotal#
     PRINT #10, ""
     PRINT #10, TAB(103); USING "#########,#"; GLandTotal#;
      PRINT #10, TAB(118); USING "#########,#"; GBldgTotal#;
      PRINT #10, TAB(136); USING "#########,#"; GLandTotal# + GBldgTotal#;
      PRINT #10, TAB(159); USING "#######,#.##"; GTaxTotal#;
      PRINT #10, TAB(185); USING "#####,#"; GCustomers#
     
     FOR Cnt = 1 TO 32: PRINT #10, "": NEXT Cnt
     PRINT #10, ""
     ' End of NON Exempt Totals Page

        Page! = Page! + 1

        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, TAB(100); TaxYear$; TAB(155); USING "$##.##"; TaxRate!;
          PRINT #10, TAB(183); "$100"
        PRINT #10, TAB(20); "IN LUNENBURG COUNTY"; TAB(80); "EXEMPT"; TAB(130); Commissioner$; TAB(188); Page!
        PRINT #10, ""
        PRINT #10, TAB(12); STRING$(189, "#")
        PRINT #10, TAB(12); STRING$(189, "#")
        PRINT #10, TAB(104); "LAND VALUE"; TAB(119); "BLDG VALUE"; TAB(136); "TOTAL VALUE"; TAB(162); "TOTAL TAX"; TAB(187); "COUNT"

          District = 0
     WHILE District < 10
        District = District + 1

        IF District = 8 THEN District = 9       'No District 8

        DistricttoPrint = District

        GOSUB getdistrict
        GOSUB OpenExemptSummaryFile
     FOR Cnt = 1 TO LOF(14) / LEN(TBPage(1))
     GET 14, Cnt, TBPage(1)
      Customers# = Customers# + TBPage(1).Customers
       LandTotal# = LandTotal# + TBPage(1).LandValue
       BldgTotal# = BldgTotal# + TBPage(1).BldgValue
       TaxTotal# = TaxTotal# + TBPage(1).TaxValue
     NEXT Cnt
     CLOSE 14
     PRINT #10, TAB(15); DistrictName$;
      PRINT #10, TAB(104); USING "########,#"; LandTotal#;
      PRINT #10, TAB(119); USING "########,#"; BldgTotal#;
      PRINT #10, TAB(137); USING "########,#"; LandTotal# + BldgTotal#;
      PRINT #10, TAB(185); USING "#####,#"; Customers#
     PRINT #10, ""
     PRINT #10, TAB(104); USING "########,#"; LandTotal#;
      PRINT #10, TAB(119); USING "########,#"; BldgTotal#;
      PRINT #10, TAB(137); USING "########,#"; LandTotal# + BldgTotal#;
      PRINT #10, TAB(159); USING "#######,#.##"; TaxTotal#
     PRINT #10, ""
     ELandTotal# = ELandTotal# + LandTotal#
     EBldgTotal# = EBldgTotal# + BldgTotal#
     ETaxTotal# = ETaxTotal# + TaxTotal#
     ECustomers# = ECustomers# + Customers#

     GLandTotal# = GLandTotal# + LandTotal#
     GBldgTotal# = GBldgTotal# + BldgTotal#
     GTaxTotal# = GTaxTotal# + TaxTotal#
     GCustomers# = GCustomers# + Customers#
     LandTotal# = 0: BldgTotal# = 0: TaxTotal# = 0: Customers# = 0
     WEND
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, ""

     PRINT #10, TAB(15); "EXEMPT TOTALS"; TAB(103); USING "#########,#"; ELandTotal#;
      PRINT #10, TAB(118); USING "#########,#"; EBldgTotal#;
      PRINT #10, TAB(136); USING "#########,#"; ELandTotal# + EBldgTotal#
     PRINT #10, ""
     PRINT #10, TAB(103); USING "#########,#"; ELandTotal#;
      PRINT #10, TAB(118); USING "#########,#"; EBldgTotal#;
      PRINT #10, TAB(136); USING "#########,#"; ELandTotal# + EBldgTotal#;
      PRINT #10, TAB(159); USING "#######,#.##"; ETaxTotal#;
      PRINT #10, TAB(185); USING "#####,#"; ECustomers#
     FOR Cnt = 1 TO 7: PRINT #10, "": NEXT Cnt

     PRINT #10, TAB(15); "GRAND TOTAL"; TAB(103); USING "#########,#"; GLandTotal#;
      PRINT #10, TAB(118); USING "#########,#"; GBldgTotal#;
      PRINT #10, TAB(136); USING "#########,#"; GLandTotal# + GBldgTotal#;
      PRINT #10, TAB(159); USING "#######,#.##"; GTaxTotal#;
      PRINT #10, TAB(185); USING "#####,#"; GCustomers#
     FOR Cnt = 1 TO 25: PRINT #10, "": NEXT Cnt
     ' End of NON Exempt Totals Page


         LandValue# = 0
         BldgValue# = 0
         TaxAmt# = 0
         Customer! = 0


     FOR x = 1 TO LOF(TBFile) / LEN(TBRECust(1))
                  ' NOW FIGURE CLASS TOTALS
     
     GET 1, x, TBRECust(1)
     Account! = x

     IF TBRECust(1).Deleted = "Y" THEN 2490               'only active customers to print
     
     IF VAL(TBRECust(1).LBCurYrBillNumber) = 0 THEN 2490

     TotalValue# = TBRECust(1).LandValue + TBRECust(1).BldgValue
     TaxAmount# = (TotalValue# / 100) * TaxRate!
     TaxAmount# = INT((TaxAmount# * 100) + .5) / 100
     IF TBRECust(1).LBCurYrTaxAmount = 0 THEN
     ELSE
     Class$ = LTRIM$(TBRECust(1).ClassCode)
     Class$ = RTRIM$(Class$)
     ClassValue = VAL(Class$)
     
     LandValue#(ClassValue) = LandValue#(ClassValue) + TBRECust(1).LandValue
     BldgValue#(ClassValue) = BldgValue#(ClassValue) + TBRECust(1).BldgValue
     TaxAmt#(ClassValue) = TaxAmt#(ClassValue) + TBRECust(1).LBCurYrTaxAmount
     Customer!(ClassValue) = Customer!(ClassValue) + 1
     END IF
2490 NEXT x

     ' Class Totals Page Goes Here
        Page! = Page! + 1

        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, TAB(100); TaxYear$; TAB(155); USING "$##.##"; TaxRate!;
          PRINT #10, TAB(183); "$100"
        PRINT #10, TAB(20); "IN LUNENBURG COUNTY"; TAB(130); Commissioner$; TAB(188); Page!
        PRINT #10, ""
        PRINT #10, TAB(12); STRING$(185, "#")
        PRINT #10, TAB(12); STRING$(185, "#")
        PRINT #10, TAB(15); "CLASS"; TAB(104); "LAND VALUE"; TAB(119); "BLDG VALUE"; TAB(136); "TOTAL VALUE"; TAB(162); "TOTAL TAX"; TAB(187); "COUNT"
        lc = 8
        REM 8 printed lines so far

        FOR Class = 1 TO 99
        IF Class <> 9 THEN
        IF Customer!(Class) > 0 THEN
        PRINT #10, TAB(15); USING "##"; Class;
         PRINT #10, TAB(103); USING "#########,#"; LandValue#(Class);
         PRINT #10, TAB(118); USING "#########,#"; BldgValue#(Class);
         PRINT #10, TAB(136); USING "#########,#"; LandValue#(Class) + BldgValue#(Class);
         PRINT #10, TAB(159); USING "#######,#.##"; TaxAmt#(Class);
         PRINT #10, TAB(185); USING "#####,#"; Customer!(Class)
         PRINT #10,
         lc = lc + 2
         LandValue# = LandValue# + LandValue#(Class)
         BldgValue# = BldgValue# + BldgValue#(Class)
         TaxAmt# = TaxAmt# + TaxAmt#(Class)
         Customer! = Customer! + Customer!(Class)
        END IF
        END IF
        NEXT Class
        PRINT #10, TAB(15); "INVALID";
         PRINT #10, TAB(103); USING "#########,#"; LandValue#(9) + LandValue#(0);
         PRINT #10, TAB(118); USING "#########,#"; BldgValue#(9) + BldgValue#(0);
         PRINT #10, TAB(136); USING "#########,#"; LandValue#(9) + BldgValue#(9) + LandValue#(0) + BldgValue#(0);
         PRINT #10, TAB(159); USING "#######,#.##"; TaxAmt#(9) + TaxAmt#(0);
         PRINT #10, TAB(185); USING "#####,#"; Customer!(9) + Customer!(0)
         lc = lc + 1
         LandValue# = LandValue# + LandValue#(9) + LandValue#(0)
         BldgValue# = BldgValue# + BldgValue#(9) + BldgValue#(0)
         TaxAmt# = TaxAmt# + TaxAmt#(9) + TaxAmt#(0)
         Customer! = Customer! + Customer!(9) + Customer!(0)

        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, TAB(15); "TOTALS";
         PRINT #10, TAB(103); USING "#########,#"; LandValue#;
         PRINT #10, TAB(118); USING "#########,#"; BldgValue#;
         PRINT #10, TAB(136); USING "#########,#"; LandValue# + BldgValue#;
         PRINT #10, TAB(159); USING "#######,#.##"; TaxAmt#;
         PRINT #10, TAB(185); USING "#####,#"; Customer!
         lc = lc + 3
         FOR LL = lc + 1 TO 84: PRINT #10, "": NEXT LL

        'Last Page Thank You Lord

        FOR LL = 1 TO 12
        PRINT #10, ""
        NEXT LL
        PRINT #10, TAB(20); "I, _____________________________, COMMISSIONER OF THE REVENUE FOR THE COUNTY OF LUNENBURG,"
        PRINT #10, TAB(20); "VIRGINIA, DO SOLEMNLY SWEAR AND AFFIRM THAT IN MAKING THIS BOOK, HAVE TO THE BEST OF MY ABILITY,"
        PRINT #10, TAB(20); "COMPLIED WITH THE LAWS PRESCRIBING THE DUTIES OF COMMISSIONER OF THE REVENUE"
        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, TAB(20); "GIVEN UNDER MY HAND THIS ____________ DAY OF ______________, ________ "
        PRINT #10, ""
        PRINT #10, ""
        PRINT #10, TAB(20); "_____________________________________"
        PRINT #10, TAB(20); "COMMISSIONER OF THE REVENUE"
        FOR LL = 1 TO 4
        PRINT #10, ""
        PRINT #10, ""
        NEXT LL
        PRINT #10, TAB(20); "SWORN TO BEFORE ME __________________________________, A NOTORY PUBLIC"
        PRINT #10,
        PRINT #10, TAB(20); "(OR OTHER PERSON AUTHORIZED TO ADMINISTER OATH) FOR THE COMMONWEALTH OF"
        PRINT #10,
        PRINT #10, TAB(20); "VIRGINIA, COUNTY OF LUNENBURG, ON THE ____ DAY OF ________________,______."
        PRINT #10,
        PRINT #10,
        PRINT #10, TAB(60); "__________________________________________________"
        PRINT #10, TAB(60); "(PERSON AUTHORIZE TO ADMINISTER OATH)"
        PRINT #10, CHR$(12)
        
        END IF


2500 CLOSE
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBLANDBK.REP"
      Header$ = "Land Book"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
      CLS
      PRINT "Please Note the Following ..."
      PRINT "Your Last District Printed Was ... "; DistricttoPrint
      PRINT "Where 11=Non-Taxable & 12=Grand Totals"
      PRINT
      PRINT
      PRINT "Your Last Bill Number Was : "; BillNumber!
      PRINT "Your Last Printed Page # Was:  "; Page!
      PRINT
      LOCATE 10, 1: PRINT "Press <ENTER> when you have recorded this information "; : fl = 1
      GOSUB 62000
      CLOSE

      GOTO 10


align:                  ' alignment mask for land book forms
ReportFile$ = "TBALGN.REP"
OPEN "O", 10, "TBALGN.REP": WIDTH #10, 255

PRINT #10, ""
PRINT #10, ""
PRINT #10, TAB(100); TaxYear$; TAB(155); USING "$##.##"; TaxRate!;
 PRINT #10, TAB(183); "$100"
PRINT #10, TAB(188); "XX"
PRINT #10, TAB(20); "IN LUNENBURG COUNTY"; TAB(50); DistrictName$;
 IF DistricttoPrint = 11 THEN
  PRINT #10, TAB(80); "EXEMPT";
 ELSE
  PRINT #10, TAB(80); "NON-EXEMPT";
 END IF
 PRINT #10, TAB(130); Commissioner$
PRINT #10, ""
PRINT #10, ""
     PRINT #10, TAB(15); "CUSTOMER NAME";
      PRINT #10, TAB(60); "Property Desc Line 1";
      PRINT #10, TAB(104); "########,#";
      PRINT #10, TAB(119); "########,#";
      PRINT #10, TAB(133); "########,#";
      PRINT #10, TAB(159); "#######,#.##";
      PRINT #10, TAB(180); "Map Number ";
      PRINT #10, TAB(198); "#######"
     PRINT #10, TAB(15); "Customer Address Line 1";
      PRINT #10, TAB(60); "Property Desc Line 2";
      PRINT #10, TAB(180); "Comment Line 1"
     PRINT #10, TAB(15); "Customer Addree Line 2";
      PRINT #10, TAB(60); "Property Desc Line 3";
      PRINT #10, TAB(154); "XX";
      PRINT #10, TAB(180); "Comment Line 2"
     PRINT #10, TAB(15); "Customer Address Line 3";
      PRINT #10, TAB(60); "Property Desc Line 4";
      PRINT #10, TAB(104); "########,#";
      PRINT #10, TAB(119); "########,#";
      PRINT #10, TAB(133); "########,#";
      PRINT #10, TAB(159); "#######,#.##";
      PRINT #10, TAB(180); "######"

     PRINT #10, TAB(15); "CUSTOMER NAME";
      PRINT #10, TAB(60); "Property Desc Line 1";
      PRINT #10, TAB(104); "########,#";
      PRINT #10, TAB(119); "########,#";
      PRINT #10, TAB(133); "########,#";
      PRINT #10, TAB(159); "#######,#.##";
      PRINT #10, TAB(180); "Map Number ";
      PRINT #10, TAB(198); "#######"
     PRINT #10, TAB(15); "Customer Address Line 1";
      PRINT #10, TAB(60); "Property Desc Line 2";
      PRINT #10, TAB(180); "Comment Line 1"
     PRINT #10, TAB(15); "Customer Addree Line 2";
      PRINT #10, TAB(60); "Property Desc Line 3";
      PRINT #10, TAB(154); "XX";
      PRINT #10, TAB(180); "Comment Line 2"
     PRINT #10, TAB(15); "Customer Address Line 3";
      PRINT #10, TAB(60); "Property Desc Line 4";
      PRINT #10, TAB(104); "########,#";
      PRINT #10, TAB(119); "########,#";
      PRINT #10, TAB(133); "########,#";
      PRINT #10, TAB(159); "#######,#.##";
      PRINT #10, TAB(180); "######"
      FOR LCNT = 16 TO 84: PRINT #10, "": NEXT LCNT
      CLOSE 10

      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBALGN.REP"
      Header$ = "Land Book Preliminary Report"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$



      RETURN

TopOfPage:
Page! = Page! + 1
PRINT #10, ""
PRINT #10, ""
PRINT #10, TAB(100); TaxYear$; TAB(155); USING "$##.##"; TaxRate!;
 PRINT #10, TAB(183); "$100"
PRINT #10, TAB(20); "IN LUNENBURG COUNTY"; TAB(50); DistrictName$;
IF DistricttoPrint = 11 THEN
 PRINT #10, TAB(80); "  EXEMPT";
 ELSE
 PRINT #10, TAB(80); "NON-EXEMPT";
END IF
 PRINT #10, TAB(130); Commissioner$; TAB(188); Page!
PRINT #10, ""
PRINT #10, ""
PRINT #10, ""
AcctsPrinted = 0
RETURN


PageSubtotal:
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, TAB(104); USING "########,#"; TotalLandSub#;
      PRINT #10, TAB(119); USING "########,#"; TotalBldgSub#;
      PRINT #10, TAB(133); USING "########,#"; TotalLandSub# + TotalBldgSub#;
      PRINT #10, TAB(159); USING "#######,#.##"; TotalTaxAmount#
      PRINT #10, ""
      PRINT #10, TAB(22); "Class-1"; TAB(40); "Class-2"; TAB(60); "Class-3"; TAB(80); "Class-4"; TAB(98); "Class-5"; TAB(120); "Class-6"; TAB(140); "Class-7"; TAB(160); "Class-8"; TAB(178); "Class-9"
      PRINT #10, ""
      PRINT #10, TAB(20); USING "########,#"; TotalValue#(1);
       PRINT #10, TAB(38); USING "########,#"; TotalValue#(2);
       PRINT #10, TAB(58); USING "########,#"; TotalValue#(3);
       PRINT #10, TAB(78); USING "########,#"; TotalValue#(4);
       PRINT #10, TAB(96); USING "########,#"; TotalValue#(5);
       PRINT #10, TAB(118); USING "########,#"; TotalValue#(6);
       PRINT #10, TAB(138); USING "########,#"; TotalValue#(7);
       PRINT #10, TAB(158); USING "########,#"; TotalValue#(8);
       PRINT #10, TAB(176); USING "########,#"; TotalValue#(9);
       PRINT #10, TAB(195); USING "######,#"; TotalValue#(0)
      PRINT #10, ""
      PRINT #10, ""
      IF EXEMPTFLAG$ = "Y" THEN
        GOSUB OpenExemptSummaryFile
      ELSE
        GOSUB OpenSummaryFile
      END IF
      TBPage(1).Page = Page!
      TBPage(1).Class1 = TotalValue#(1)
      TBPage(1).Class2 = TotalValue#(2)
      TBPage(1).Class3 = TotalValue#(3)
      TBPage(1).Class4 = TotalValue#(4)
      TBPage(1).Class5 = TotalValue#(5)
      TBPage(1).Class6 = TotalValue#(6)
      TBPage(1).Class7 = TotalValue#(7)
      TBPage(1).Class8 = TotalValue#(8)
      TBPage(1).Class9 = TotalValue#(9)
      TBPage(1).Class0 = TotalValue#(0)
      TBPage(1).Customers = CustSub
      TBPage(1).LandValue = TotalLandSub#
      TBPage(1).BldgValue = TotalBldgSub#
      TBPage(1).TaxValue = TotalTaxAmount#
      PUT 14, PageRecord, TBPage(1)
      CLOSE 14
      PageRecord = PageRecord + 1
      FOR Class = 0 TO 9: TotalValue#(Class) = 0: NEXT Class
      CustSub = 0
      TotalLandSub# = 0
      TotalBldgSub# = 0
      TotalTaxAmount# = 0
      RETURN


OpenSummaryFile:
District$ = STR$(District)
District$ = LTRIM$(District$)
OPEN TBpath$ + "TBDIST" + District$ + ".dat" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = LEN(TBPage(1))
RETURN

OpenExemptSummaryFile:
District$ = STR$(District)
District$ = LTRIM$(District$)
OPEN TBpath$ + "TBEDIS" + District$ + ".dat" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = LEN(TBPage(1))
RETURN



3000 scren$ = "2.02.03": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 24: PRINT "Print Land Book District Total Pages"
     LOCATE 3, 1: COLOR 11
     PRINT "Enter the Tax Year (YYYY) (ESC-Cancel) .... "; : fl = -4: GOSUB 62000
     IF cf = 1 THEN 10
     IF LEN(in$) <> 4 THEN 3000
     TaxYear$ = in$
     COLOR 11
     PRINT "Enter the Tax Rate Per $100 (ESC-Cancel) .. "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     TaxRate! = VAL(in$)
     COLOR 11
     PRINT "Enter the Commissioner's Name to Print .... "; : fl = 35: GOSUB 62000
     IF cf = 1 THEN 10
     Commissioner$ = UCASE$(in$)
3099 COLOR 11: LOCATE 7, 1
     PRINT "Enter the District Number to Print (1-10) . "; : fl = -2: GOSUB 62000
     IF cf = 1 THEN 10
     DistricttoPrint = VAL(in$)
     IF VAL(in$) < 1 OR VAL(in$) > 10 THEN 3099
     PRINT ""
     COLOR 11
     PRINT "Enter the Page Number to Start (1-??) ..... "; : fl = -4: GOSUB 62000
     IF cf = 1 THEN 10
     Page! = VAL(in$)
3100 REM for review
     FOR LL = 3 TO 12: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
     LOCATE 3, 1
     COLOR 11
     PRINT "You Have Entered the Following ..."
     PRINT "Tax Year is    "; : COLOR 15: PRINT TaxYear$: COLOR 11
     PRINT "Tax Rate is    "; : COLOR 15: PRINT USING "$##.##"; TaxRate!: COLOR 11
     PRINT "Commissioner   "; : COLOR 15: PRINT Commissioner$: COLOR 11
     PRINT "Print District "; : COLOR 15: PRINT DistricttoPrint: COLOR 11
     PRINT "Page to Start  "; : COLOR 15: PRINT Page!: COLOR 11
     PRINT

     PRINT "Is the Above Information Correct (Y/N) .. "; : fl = 1
     GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN 3100
     IF in$ = "N" THEN 3000
     Page! = Page! - 1: REM RESET PAGE TO ZERO FOR PROPER PROCESSING
     BillNumber! = BillNumber! - 1
3110 REM Now Print Alignment Test oki 395 must be in 15 character per inch mode
     LOCATE 12, 1: COLOR 11
     PRINT "The printer must be set to print in 15 cpi : Please Set Printer Now"
     PRINT "Press <ENTER> when the printer is set to 15 cpi "; : fl = 1
     GOSUB 62000
3111 scren$ = "2.02.01": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 32: PRINT "Re-Print Land Book Totals"
     LOCATE 3, 1: COLOR 11
     LOCATE 15, 1: COLOR 11
     PRINT "Print an Alignment Mask Now (Y/N) (ESC-Cancel) .. "; : fl = 1
     GOSUB 62210
        IF cf = 1 THEN CLOSE : GOTO 10
        IF NOT (in$ = "Y" OR in$ = "N") THEN 3111
        IF in$ = "Y" THEN GOSUB align: GOTO 3111

     scren$ = "2.02.01": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 32: PRINT "Print Land Book"
     LOCATE 3, 1: COLOR 11

3120 ' Set Constants Here

     ' Open Report File
     OPEN "O", 10, "TBLANDBK.REP": WIDTH #10, 255

     

     taxtype$ = "R"
     GOSUB 60000
     LandValue# = 0
     BldgValue# = 0
     TaxAmt# = 0
     District = DistricttoPrint
     GOSUB getdistrict

     ' Summary Pages To Print Here

    GOSUB OpenSummaryFile

    NumberofPages = LOF(14) / LEN(TBPage(1))

    GOSUB TopOfPage
    PagesPrinted = 0

    FOR Cnt = 1 TO NumberofPages
    GET 14, Cnt, TBPage(1)

      IF PagesPrinted >= 17 THEN
       FOR lll = 1 TO 8
        PRINT #10, ""
       NEXT lll
        PRINT #10, ""
       GOSUB TopOfPage
       PagesPrinted = 0
      END IF

      PRINT #10, TAB(16); "1- "; USING "########,#"; TBPage(1).Class1;
       PRINT #10, TAB(37); "5- "; USING "########,#"; TBPage(1).Class5;
       PRINT #10, TAB(65); "9- "; USING "########,#"; TBPage(1).Class9;
       PRINT #10, TAB(105); USING "########,#"; TBPage(1).LandValue;
       PRINT #10, TAB(120); USING "########,#"; TBPage(1).BldgValue;
       PRINT #10, TAB(135); USING "#########,#"; TBPage(1).LandValue + TBPage(1).BldgValue;
       PRINT #10, TAB(160); USING "########,#.##"; TBPage(1).TaxValue;
       PRINT #10, TAB(185); USING "####"; TBPage(1).Page
      PRINT #10, TAB(16); "2- "; USING "########,#"; TBPage(1).Class2;
       PRINT #10, TAB(37); "6- "; USING "########,#"; TBPage(1).Class6;
       PRINT #10, TAB(65); " - "; USING "########,#"; TBPage(1).Class0
      PRINT #10, TAB(16); "3- "; USING "########,#"; TBPage(1).Class3;
       PRINT #10, TAB(37); "7- "; USING "########,#"; TBPage(1).Class7
      PRINT #10, TAB(16); "4- "; USING "########,#"; TBPage(1).Class4;
       PRINT #10, TAB(37); "8- "; USING "########,#"; TBPage(1).Class8;
       PRINT #10, TAB(105); USING "########,#"; TBPage(1).LandValue;
       PRINT #10, TAB(120); USING "########,#"; TBPage(1).BldgValue;
       PRINT #10, TAB(135); USING "#########,#"; TBPage(1).LandValue + TBPage(1).BldgValue;
       PRINT #10, TAB(160); USING "########,#.##"; TBPage(1).TaxValue

       ' Subtotal Here

      LandValue# = LandValue# + TBPage(1).LandValue
      BldgValue# = BldgValue# + TBPage(1).BldgValue
      TaxAmt# = TaxAmt# + TBPage(1).TaxValue
      PagesPrinted = PagesPrinted + 1

     NEXT Cnt
     CLOSE 14

     FOR lll = PagesPrinted + 1 TO 17
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
      PRINT #10, ""
     NEXT lll
     PRINT #10, ""
     PRINT #10, ""
     PRINT #10, TAB(104); USING "########,#"; LandValue#;
      PRINT #10, TAB(119); USING "########,#"; BldgValue#;
      PRINT #10, TAB(133); USING "########,#"; LandValue# + BldgValue#;
      PRINT #10, TAB(159); USING "#######,#.##"; TaxAmt#

     FOR Cnt = 1 TO 6: PRINT #10, "": NEXT Cnt



3500 CLOSE
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBLANDBK.REP"
      Header$ = "Land Book Totals Reprint"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
      
      GOTO 10
     
4000 scren$ = "2.04.00": CLS : GOSUB 63500
     'SET NULL VALUES
      AcctsPrinted = 0
      TotalLand# = 0
      TotalBldg# = 0
      TOTALTAX# = 0
      P = 0

     COLOR 11: LOCATE 1, 25: PRINT "Print Land Book - Kenbridge/Victoria"
     LOCATE 3, 1: COLOR 11
     PRINT "Enter the Tax Year (YYYY) (ESC-Cancel) .... "; : fl = -4: GOSUB 62000
     IF cf = 1 THEN 10
     IF LEN(in$) <> 4 THEN 4000
     TaxYear$ = in$
     COLOR 11
     PRINT "Enter the Tax Rate Per $100 (ESC-Cancel) .. "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     TaxRate! = VAL(in$)
     COLOR 11
4099 COLOR 11: LOCATE 7, 1
     PRINT "Enter the District Number to Print (9-10) . "; : fl = -2: GOSUB 62000
     IF cf = 1 THEN 10
     DistricttoPrint = VAL(in$)
     District = VAL(in$)
     IF District < 9 OR District > 10 THEN 4099
     COLOR 11
     PRINT "Enter the Beginning Bill Number ........... "; : fl = -7: GOSUB 62000
     BillNumber! = VAL(in$)
4100 REM for review
     FOR LL = 3 TO 12: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
     LOCATE 3, 1
     COLOR 11
     PRINT "You Have Entered the Following ..."
     PRINT "Tax Year is    "; : COLOR 15: PRINT TaxYear$: COLOR 11
     PRINT "Tax Rate is    "; : COLOR 15: PRINT USING "$##.##"; TaxRate!: COLOR 11
     PRINT "Print District "; : COLOR 15: PRINT DistricttoPrint: COLOR 11
     PRINT "Bill #         "; : COLOR 15: PRINT BillNumber!: COLOR 11
     PRINT
     PRINT "Is the Above Information Correct (Y/N) .. "; : fl = 1
     GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN 4100
     IF in$ = "N" THEN 4000
     
4110 REM Start Report Now
     GOSUB getdistrict

     OPEN "O", 10, "TBTWNLND.REP": WIDTH #10, 255
     GOSUB 4900
     OPEN TBpath$ + "TBREname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
     FIELD 4, 20 AS lastnam$, 4 AS MREC$: EndofFile = LOF(4) / 24
     TBFile = FREEFILE
     OPEN TBpath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
     LOCATE 15, 1: COLOR 11
     PRINT "NOW SPOOLING OUT DISTRICT # "; DistricttoPrint
4170 FOR x = 1 TO EndofFile
     GET 4, x
     GET TBFile, CVS(MREC$), TBRECust(1)
     Account! = CVS(MREC$)


4191 IF TBRECust(1).Deleted = "Y" THEN 4230
     IF VAL(TBRECust(1).District) <> DistricttoPrint THEN 4230
     IF LEFT$(TBRECust(1).ClassCode, 1) = "7" THEN 4230

     ' First Figure the Tax Amount
     TotalValue# = TBRECust(1).LandValue + TBRECust(1).BldgValue
     IF TotalValue# = 0 THEN
     TBRECust(1).MPCurYrBillNumber = ""
     TBRECust(1).MPCurYrTaxAmount = 0

      GOTO 4230
     END IF

     TaxAmount# = (TotalValue# / 100) * TaxRate!
     TaxAmount# = INT((TaxAmount# * 100) + .5) / 100

     TOTALTAX# = TOTALTAX# + TaxAmount#

     TBRECust(1).MPCurYrBillNumber = LTRIM$(STR$(BillNumber!))
     TBRECust(1).MPCurYrTaxAmount = TaxAmount#
     PUT TBFile, CVS(MREC$), TBRECust(1)


    
    'Print the Account Information
     PRINT #10, USING "#####"; Account!;
     PRINT #10, TAB(10); TBRECust(1).CustName;
      PRINT #10, TAB(50); USING "########,#"; TBRECust(1).LandValue;
      PRINT #10, TAB(60); USING "########,#"; TBRECust(1).BldgValue;
      PRINT #10, TAB(70); USING "########,#"; TotalValue#;
      PRINT #10, TAB(80); USING "#######,#.##"; TaxAmount#;
      PRINT #10, TAB(95); USING "#######"; BillNumber!
      AcctsPrinted = AcctsPrinted + 1
      TotalTaxAmount# = TotalTaxAmount# + TaxAmount#
      TotalLand# = TotalLand# + TBRECust(1).LandValue
      TotalLand# = INT((TotalLand# * 100) + .5) / 100
      TotalBldg# = TotalBldg# + TBRECust(1).BldgValue
      TotalBldg# = INT((TotalBldg# * 100) + .5) / 100
      lc = lc + 1
      
     ' Set Bill Number
     BillNumber! = BillNumber! + 1

4220 IF lc >= 56 THEN PRINT #10, CHR$(12); : GOSUB 4900
4230 NEXT x
     PRINT #10, STRING$(110, "-")
     PRINT #10, "Total Customers .. "; USING "######,#"; AcctsPrinted
     PRINT #10, "Total Land Value: "; USING "$$#########,#"; TotalLand#
     PRINT #10, "Total Bldg Value: "; USING "$$#########,#"; TotalBldg#
     PRINT #10, "Total Valuation : "; USING "$$#########,#"; TotalLand# + TotalBldg#
     PRINT #10, "Total Tax Levied: "; USING "$$#######,#.##"; TOTALTAX#
     PRINT CHR$(12);
4240 CLOSE
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBTWNLND.REP"
      Header$ = "Land Book Town Report"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
4250 GOTO 10

4900 P = P + 1
     PRINT #10, TAB(27); "Land Book for "; DistrictName$
     PRINT #10, "Date: "; DATE$; TAB(98); "Page - "; USING "#####"; P
4910 PRINT #10, "Acct. #"; TAB(10); "Customer Name"; TAB(52); "Land Val"; TAB(62); "Bldg Val"; TAB(71); "Tot Value"; TAB(82); "Tax Amount"; TAB(96); "Bill #"
     PRINT #10, STRING$(110, "="): lc = 4
     RETURN



60000  IF taxtype$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBpath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBpath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
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

62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(Xy): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
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

62065 IF ASC(z$) = 0 THEN 62030
62070 IF LEN(z$) > 1 THEN cf = ASC(RIGHT$(z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
62080 IF z$ = " " THEN 62130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 62000
62090 IF z$ <> CHR$(8) THEN 62130 ELSE IF ZL = 0 THEN 62030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 62110 ELSE IF z$ = "," THEN 62120
62100 IF z$ = "." THEN ZD = 0: GOTO 62110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
62110 in$ = LEFT$(in$, LEN(in$) - 1)
62120 ZL = ZL - 1: GOTO 62030
62130 IF ABS(fl) = ZL THEN 62030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 62180
62140 IF z$ = "." AND ZD = 0 THEN ZD = 1: GOTO 62180
62150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF z$ < "0" OR z$ > "9" THEN 62030
62180 PRINT z$; : in$ = in$ + z$: ZL = ZL + 1
62190 IF ABS(fl) = 1 THEN 62050 ELSE 62030
62200 LOCATE RO, PO: PRINT in$: RETURN
62210 GOSUB 62000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
62220 RETURN
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SYSDATE$; TAB(70); scren$: PRINT STRING$(79, "_"): RETURN

Namesearch:
 LOCATE 8, 1:
 COLOR 11
 PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : fl = 25
 GOSUB 62000
 IF cf = 1 THEN
  ggo = 0
  FOR lll = 8 TO 13: PRINT STRING$(79, 32): NEXT lll
  RETURN
 END IF
 nn$ = in$
 nn = LEN(nn$)
 IF nn > 20 THEN nn = 20
 nn$ = LEFT$(nn$, nn)

 IF taxtype$ = "R" THEN
  TBIDXFile = FREEFILE
  OPEN TBpath$ + "TBRENAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS MREC$
  GOSUB 60000
  FOR LL = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, LL
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO n1
n3:
  NEXT LL
  FOR lll = 8 TO 13: PRINT STRING$(79, 32): NEXT lll
  ggo = 0
  RETURN

n1:
GET TBFile, CVS(MREC$), TBRECust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBRECust(1).CustName
COLOR 10
FOR lll = 10 TO 12: LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(MREC$)
n2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO n2
IF in$ = "Q" THEN
  FOR lll = 8 TO 13: LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
  CLOSE
  ggo = 0
  RETURN
END IF
IF in$ = "Y" THEN
  ggo = 1: Record! = CVS(MREC$)
  FOR lll = 8 TO 13:  LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
  CLOSE
  RETURN
END IF
IF in$ = "N" THEN GOTO n3
GOTO n2
END IF


  TBIDXFile = FREEFILE
  OPEN TBpath$ + "TBPPNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS MREC$
  GOSUB 60000
  FOR LL = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, LL
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO PPn1
PPn3:
  NEXT LL
  LOCATE 8, 1: PRINT STRING$(79, 32)
  ggo = 0
  RETURN

PPn1:
GET TBFile, CVS(MREC$), TBPPCust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBPPCust(1).CustName
COLOR 10
FOR lll = 10 TO 12: LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(MREC$)
PPn2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO PPn2
IF in$ = "Q" THEN
  FOR lll = 8 TO 13: LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
  CLOSE
  ggo = 0
  RETURN
END IF
IF in$ = "Y" THEN
  ggo = 1: Record! = CVS(MREC$)
  FOR lll = 8 TO 13: LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
  CLOSE
  RETURN
END IF
IF in$ = "N" THEN GOTO PPn3
GOTO PPn2

getdistrict:
     DistrictName$ = ""
     IF District = 1 THEN DistrictName$ = "BROWN'S STORE"
     IF District = 2 THEN DistrictName$ = "COLUMBIAN GROVE"
     IF District = 3 THEN DistrictName$ = "LEWISTON"
     IF District = 4 THEN DistrictName$ = "LOCHLEVEN"
     IF District = 5 THEN DistrictName$ = "PLEASANT GROVE"
     IF District = 6 THEN DistrictName$ = "PLYMOUTH"
     IF District = 7 THEN DistrictName$ = "REHOBOTH"
     IF District = 9 THEN DistrictName$ = "TOWN OF KENBRIDGE"
     IF District = 10 THEN DistrictName$ = "TOWN OF VICTORIA"
     RETURN


FigureBalance:
     CurBal# = 0: PastDue# = 0
     IF taxtype$ = "R" THEN
     CurBal# = TBRECust(1).CurYrTaxAmount + TBRECust(1).CurYrIntAmount + TBRECust(1).CurYrPenAmount + TBRECust(1).CurYrOthAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR LL = 1 TO 10
     PastDue# = PastDue# + TBRECustBal(1).TaxAmt(LL) + TBRECustBal(1).IntAmt(LL) + TBRECustBal(1).PenAmt(LL) + TBRECustBal(1).OthAmt(LL) + TBRECustBal(1).LateAmt(LL)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT LL
     ELSE
     CurBal# = TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax + TBPPCust(1).CurYrIntAmount + TBPPCust(1). _
CurYrPenAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR LL = 1 TO 10
     PastDue# = PastDue# + TBPPCustBal(1).PerPropTaxAmt(LL) + TBPPCustBal(1).FarmEquipAmt(LL) + TBPPCustBal(1).MachToolsAmt(LL) + TBPPCustBal(1).MerchCapAmt(LL) + TBPPCustBal(1).MobileHomeAmt(LL) + TBPPCustBal(1).LateFeeAmt(LL) + TBPPCustBal(1). _
IntAmt(LL) + TBPPCustBal(1).PenAmt(LL)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT LL


     END IF
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

