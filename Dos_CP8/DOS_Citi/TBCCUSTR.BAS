10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbcustrep
30 REM Customer Report Menu
40 REM Last Revision Date: 8-24-87
50 REM by Nick Picerno
 
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
   'TBpath$ = "F:\"                    ' change to "F:\" Before compiling

   CLOSE
   p = 0: page = 0
   OPEN "R", 1, TBpath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
   CurrentYear$ = year$
   CLOSE 1
   'ON ERROR GOTO networkerror


     REDIM TransLink!(500)
     REDIM TBRECust(1) AS TBRECust
     REDIM TBRECustBal(1) AS TBREBAL
     REDIM TBPPCust(1) AS TBPPCust
     REDIM TBPPCustBal(1) AS TBPPBal
     REDIM TBTrans(1) AS TBTrans

     
100 REM begin program
110 SCREN$ = "2.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500
130 LOCATE 3, 1: COLOR 11
200 PRINT TAB(24); "Tax Billing Reports - Commissioner "
220 PRINT
230 PRINT TAB(5); "1. Print Master Customer Listing";
     PRINT TAB(44); "10. Print Reassessment Proof"
    PRINT
    PRINT TAB(5); "2. Print Account Valuation Listing";
     PRINT TAB(44); "11. Print Reassessment Notices"
    PRINT
    PRINT TAB(5); "3. Print Mailing Labels";
     PRINT TAB(44); "12. Print Reassessment Book"

280 PRINT
290 PRINT TAB(5); "4. Print Property Record Sheets";
     PRINT TAB(44); "13. Vehicle Record Review"

300 PRINT
302 PRINT TAB(5); "5. Print Personal Property Sheets"; TAB(44); "14. Vehicle Certification Filter"
    PRINT
    PRINT TAB(5); "6. Print Land Book"; TAB(44); "15. Vehicle Certification Report"
    PRINT
    PRINT TAB(5); "7. Print Personal Prop Book"; TAB(44); "16. Customer w/Boat Value"
    PRINT
    PRINT TAB(5); "8. Print Alpha Customer Listing"; TAB(44); "17. Trailor Report"
    PRINT
    PRINT TAB(5); "9. Print Map Order Customer Listing"; TAB(44); "18. Report Menu 2"
310 LOCATE 23, 1: COLOR 11: PRINT TAB(25); "Selection No. (ESC-Master Menu) : "; : fl = -2: GOSUB 62000
320 IF cf = 1 THEN CLOSE : RUN "tbcmenu"
325 IF LEN(in$) = 0 THEN 310
330 IF VAL(in$) < 1 OR VAL(in$) > 18 THEN 310
340 ON VAL(in$) GOTO 1000, 2000, 3000, 4000, 5000, 6000, 7000, 8000, 9000, 10000, 11000, 12000, 13000, 14000, 15000, 16000, 17000, 18000

1000 REM
1010 SCREN$ = "2.01.00": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 32: PRINT "Customer Account Listing"
1020 LOCATE 4, 1
1030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount # or (L)ast Name (A,L  or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
1040 IF cf = 1 THEN 10
1050 IF NOT (in$ = "A" OR in$ = "L" OR in$ = "M") THEN 1020
1060 SortSpec$ = in$: LOCATE 4, 79: PRINT SortSpec$
1070 LOCATE 6, 1: COLOR 11
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-End) ..... "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "P" OR in$ = "A") THEN 1070
     TaxType$ = UCASE$(in$)


1080 COLOR 11: LOCATE 8, 1
     PRINT "Output Specification: Output to (S)creen or (P)rinter (S,P or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
1090 IF cf = 1 THEN 10
1100 IF NOT (in$ = "S" OR in$ = "P") THEN 1070
1105 devspec$ = in$: LOCATE 8, 77: PRINT devspec$
1110 LOCATE 10, 1: COLOR 11: PRINT STRING$(79, "="): COLOR 10: LOCATE 13, 23: PRINT "Now Processing : Please be Patient!": COLOR 14
1120 IF devspec$ = "S" THEN dev$ = "SCRN:" ELSE dev$ = "lpt1:"
1130 OPEN "O", 10, dev$: WIDTH #10, 255
1140 GOSUB 60000: IF devspec$ = "S" THEN CLS : GOSUB 1900 ELSE GOSUB 1900
1150 IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10
1160 IF SortSpec$ = "A" THEN
        IF TaxType$ = "R" THEN
         EndofFile = LOF(TBFile) / LEN(TBRECust(1))
        ELSE
         EndofFile = LOF(TBFile) / LEN(TBPPCust(1))
        END IF
     END IF

     IF SortSpec$ = "L" THEN
      IF TaxType$ = "R" THEN
        OPEN TBpath$ + "TBREname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS mrec$: EndofFile = LOF(4) / 24
        ELSE
        OPEN TBpath$ + "TBPPname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS mrec$: EndofFile = LOF(4) / 24
      END IF
     END IF
1170 FOR X = 1 TO EndofFile
     IF SortSpec$ = "L" THEN
      GET 4, X
      IF TaxType$ = "R" THEN
       GET 1, CVS(mrec$), TBRECust(1)
       ELSE
       GET 1, CVS(mrec$), TBPPCust(1)
      END IF
      Account! = CVS(mrec$)
      ELSE
      IF TaxType$ = "R" THEN
      GET 1, X, TBRECust(1)
      ELSE
      GET 1, X, TBPPCust(1)
      END IF
      Account! = X
     END IF
     
1191 IF TaxType$ = "R" THEN
     IF TBRECust(1).Deleted = "Y" THEN 1230
     PRINT #10, Account!
     PRINT #10, TBRECust(1).CustName
     PRINT #10, TBRECust(1).Address1
     PRINT #10, TBRECust(1).Address2
     PRINT #10, RTRIM$(TBRECust(1).Address3)
     ELSE
     IF TBPPCust(1).Deleted = "Y" THEN 1230
     PRINT #10, Account!
     PRINT #10, TBPPCust(1).CustName
     PRINT #10, TBPPCust(1).Address1; TAB(50); "Soc Sec #"; TBPPCust(1).SocSec
     PRINT #10, TBPPCust(1).Address2; TAB(50); "Soc Sec #"; TBPPCust(1).SocSec2
     PRINT #10, TBPPCust(1).City; ", "; TBPPCust(1).State; " "; TBPPCust(1).Zip
     END IF
     IF TaxType$ = "R" THEN
     PRINT #10, "Prop Desc:"; TAB(12); TBRECust(1).PropDesc1; TAB(50); "Land Value: "; USING "$$######,#"; TBRECust(1).LandValue
     PRINT #10, TAB(12); TBRECust(1).PropDesc2; TAB(50); "Bldg Value: "; USING "$$######,#"; TBRECust(1).BldgValue
     PRINT #10, TAB(12); TBRECust(1).PropDesc3; TAB(50); "      Size: "; USING "#####.####"; VAL(TBRECust(1).Acres)
     PRINT #10, TAB(12); TBRECust(1).PropDesc4; TAB(50); "  District: "; TBRECust(1).District
     PRINT #10, "Map Numb:"; TAB(12); TBRECust(1).MapNumber; TAB(50); "     Class: "; TBRECust(1).ClassCode
     lc = lc + 5
     END IF
     PRINT #10, STRING$(79, "-"): lc = lc + 5
     CustCount = CustCount + 1
     IF devspec$ = "S" THEN 1210 ELSE 1220
1210 IF lc >= 14 THEN
      LOCATE 23, 1: COLOR 11
      PRINT "Press any key for next page (ESC-Quit)!"; : fl = 1
      GOSUB 62000
        IF cf = 1 THEN
         CLOSE
         GOTO 10
        ELSE
         CLS
         GOSUB 1900
         GOTO 1230
        END IF
      END IF
1220 IF lc >= 55 THEN PRINT #10, CHR$(12); : GOSUB 1900
1230 NEXT X: PRINT #10, "Total Customers .. "; USING "######,#"; CustCount
     IF devspec$ = "S" THEN LOCATE 23, 1: COLOR 11: PRINT "All done, press any key for menu "; : fl = 1: GOSUB 62000 ELSE PRINT #10, CHR$(12);
1240 CLOSE
1250 GOTO 10

1900 p = p + 1
     PRINT #10, TAB(27); "Tax Customer Account Listing": PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "####"; p
    IF ttyp$ = "R" THEN PRINT #10, "Real Estate Customers"
     IF ttyp$ = "P" THEN PRINT #10, "Personal Property Customers"
1910 PRINT #10, "Acct. #"
     PRINT #10, STRING$(79, "="): lc = 7: RETURN

2000 SCREN$ = "2.02.00": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 32: PRINT "Customer Account Listing"
     CustCount = 0
     page = 0
2030 LOCATE 4, 1: COLOR 11: PRINT "Sort Specification: Sort by (A)ccount # or (L)ast Name (A,L  or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
2040 IF cf = 1 THEN 10
2050 IF NOT (in$ = "A" OR in$ = "L") THEN 2030
2060 SortSpec$ = in$: LOCATE 4, 79: PRINT SortSpec$
2070 LOCATE 6, 1: COLOR 11
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-End) ..... "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "P" OR in$ = "A") THEN 2070
     TaxType$ = UCASE$(in$)
2080 COLOR 11: LOCATE 8, 1
     PRINT "Output Specification: Output to (S)creen or (P)rinter (S,P or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
2090 IF cf = 1 THEN 10
2100 IF NOT (in$ = "S" OR in$ = "P") THEN 2080
2105 devspec$ = in$: LOCATE 8, 77: PRINT devspec$
2110 LOCATE 10, 1: COLOR 11: PRINT STRING$(79, "="): COLOR 10: LOCATE 13, 23: PRINT "Now Processing : Please be Patient!": COLOR 14
2120 IF devspec$ = "S" THEN dev$ = "SCRN:" ELSE dev$ = "lpt1:"
2130 OPEN "O", 10, dev$: WIDTH #10, 255
2140 GOSUB 60000
     IF devspec$ = "S" THEN CLS : GOSUB 2900 ELSE GOSUB 2900
2160 IF SortSpec$ = "A" THEN
        IF TaxType$ = "R" THEN
         EndofFile = LOF(TBFile) / LEN(TBRECust(1))
        ELSE
         EndofFile = LOF(TBFile) / LEN(TBPPCust(1))
        END IF
     END IF

     IF SortSpec$ = "L" THEN
      IF TaxType$ = "R" THEN
        OPEN TBpath$ + "TBREname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS mrec$: EndofFile = LOF(4) / 24
        ELSE
        OPEN TBpath$ + "TBPPname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS mrec$: EndofFile = LOF(4) / 24
      END IF
     END IF
2170 FOR X = 1 TO EndofFile
     IF SortSpec$ = "L" THEN
      GET 4, X
      IF TaxType$ = "R" THEN
       GET 1, CVS(mrec$), TBRECust(1)
       ELSE
       GET 1, CVS(mrec$), TBPPCust(1)
      END IF
      Account! = CVS(mrec$)
      ELSE
      IF TaxType$ = "R" THEN
      GET 1, X, TBRECust(1)
      ELSE
      GET 1, X, TBPPCust(1)
      END IF
      Account! = X
     END IF

2191 IF TaxType$ = "R" THEN
     IF TBRECust(1).Deleted = "Y" THEN 2230
     PRINT #10, Account!;
      PRINT #10, TAB(10); LEFT$(TBRECust(1).CustName, 35);
      PRINT #10, TAB(48); USING "#######,#"; TBRECust(1).LandValue;
      PRINT #10, TAB(58); USING "#######,#"; TBRECust(1).BldgValue;
      PRINT #10, TAB(68); USING "#######,#"; TBRECust(1).LandValue + TBRECust(1).BldgValue
      TotalLand# = TotalLand# + TBRECust(1).LandValue
      TotalBldg# = TotalBldg# + TBRECust(1).BldgValue
      lc = lc + 1
     ELSE
      IF TBPPCust(1).Deleted = "Y" THEN 2230
      PRINT #10, Account!;
       PRINT #10, TAB(10); TBPPCust(1).CustName;
       PRINT #10, TAB(68); USING "#######,#"; TBPPCust(1).PersProp + TBPPCust(1).MachTools + TBPPCust(1).MobileHome + TBPPCust(1).MerchCap + TBPPCust(1).FarmEquip
      PRINT #10, USING "########,#"; TBPPCust(1).PersProp;
       PRINT #10, TAB(15); USING "########,#"; TBPPCust(1).MachTools;
       PRINT #10, TAB(30); USING "########,#"; TBPPCust(1).MobileHome;
       PRINT #10, TAB(45); USING "########,#"; TBPPCust(1).MerchCap;
       PRINT #10, TAB(60); USING "########,#"; TBPPCust(1).FarmEquip
       TotalPers# = TotalPers# + TBPPCust(1).PersProp
       TotalTools# = TotalTools# + TBPPCust(1).MachTools
       TotalHome# = TotalHome# + TBPPCust(1).MobileHome
       TotalMerch# = TotalMerch# + TBPPCust(1).MerchCap
       TotalFarm# = TotalFarm# + TBPPCust(1).FarmEquip
       PRINT #10, STRING$(79, "-"): lc = lc + 3
     END IF
     
     
     CustCount = CustCount + 1
     IF devspec$ = "S" THEN 2210 ELSE 2220
2210 IF lc >= 17 THEN LOCATE 23, 1: COLOR 11: PRINT "Press any key for next page (ESC-Quit)!"; : fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10 ELSE CLS : GOSUB 2900: GOTO 2230
2220 IF lc >= 55 THEN PRINT #10, CHR$(12); : GOSUB 2900
2230 NEXT X
     IF TaxType$ = "R" THEN
      PRINT #10, "Total Customers .. "; USING "##########,#"; CustCount
      PRINT #10, "Total Land Value.. "; USING "##########,#"; TotalLand#
      PRINT #10, "Total Bldg Value.. "; USING "##########,#"; TotalBldg#
      PRINT #10, "Grand Total Value. "; USING "##########,#"; TotalLand# + TotalBldg#
      ELSE
      PRINT #10, "Total Customers .. "; USING "##########,#"; CustCount
      PRINT #10, "Total Pers Prop .. "; USING "##########,#"; TotalPers#
      PRINT #10, "Total Mach/Tools . "; USING "##########,#"; TotalTools#
      PRINT #10, "Total Mobile Home. "; USING "##########,#"; TotalHome#
      PRINT #10, "Total Merch Cap... "; USING "##########,#"; TotalMerch#
      PRINT #10, "Total Farm Equip.. "; USING "##########,#"; TotalFarm#
      PRINT #10, "Grand Total Value. "; USING "##########,#"; TotalPers# + TotalTools# + TotalHome# + TotalMerch# + TotalFarm#
     END IF



     IF devspec$ = "S" THEN
       LOCATE 23, 1: COLOR 11: PRINT "All done, press any key for menu "; : fl = 1: GOSUB 62000
       ELSE
       PRINT #10, CHR$(12);
     END IF
     CLOSE
     GOTO 10

2900 p = p + 1
     PRINT #10, TAB(27); "Tax Customer Valuation Listing"
     PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "####"; p
     PRINT #10, ""
    IF TaxType$ = "R" THEN
      PRINT #10, "Real Estate Customers"
      PRINT #10, "Acct #"; TAB(10); "Customer Name"; TAB(47); "Land Val"; TAB(58); "Bldg Val"; TAB(69); "Total Val"
      PRINT #10, STRING$(79, "="): lc = 2
     ELSE
      PRINT #10, "Personal Property Customers"
      PRINT #10, "Acct #"; TAB(10); "Customer Name"; TAB(69); "Total Value"
       PRINT #10, "Pers Prop"; TAB(15); "Mach/Tools"; TAB(30); "Mobile Home"; TAB(45); "Merch Cap"; TAB(60); "Farm Eqmt"
      PRINT STRING$(79, "="): lc = 4
     END IF
     RETURN
3000 SCREN$ = "2.03.00": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 32: PRINT "Customer Account Listing"
3020 LOCATE 3, 1
3030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount # or (L)ast Name (A,L  or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
3040 IF cf = 1 THEN 10
3050 IF NOT (in$ = "A" OR in$ = "L" OR in$ = "M") THEN 3020
3060 SortSpec$ = in$: LOCATE 4, 79: PRINT SortSpec$
3070 LOCATE 6, 1: COLOR 11
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-End) ..... "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "P" OR in$ = "A") THEN 3070
     TaxType$ = UCASE$(in$)


3080 COLOR 11: LOCATE 8, 1
     PRINT "Output Specification: Output to (S)creen or (P)rinter (S,P or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
3090 IF cf = 1 THEN 10
3100 IF NOT (in$ = "S" OR in$ = "P") THEN 3080
3105 devspec$ = in$: LOCATE 8, 77: PRINT devspec$
3110 LOCATE 10, 1: COLOR 11: PRINT STRING$(79, "="): COLOR 10: LOCATE 13, 23: PRINT "Now Processing : Please be Patient!": COLOR 14
3120 IF devspec$ = "S" THEN dev$ = "SCRN:" ELSE dev$ = "lpt1:"
3130 OPEN "O", 10, dev$: WIDTH #10, 255
3140 GOSUB 60000: IF devspec$ = "S" THEN CLS
     GOSUB Alignment
3160 IF SortSpec$ = "A" THEN
        IF TaxType$ = "R" THEN
         EndofFile = LOF(TBFile) / LEN(TBRECust(1))
        ELSE
         EndofFile = LOF(TBFile) / LEN(TBPPCust(1))
        END IF
     END IF

     IF SortSpec$ = "L" THEN
      IF TaxType$ = "R" THEN
        OPEN TBpath$ + "TBREname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS mrec$: EndofFile = LOF(4) / 24
        ELSE
        OPEN TBpath$ + "TBPPname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS mrec$: EndofFile = LOF(4) / 24
      END IF
     END IF
3170 FOR X = 1 TO EndofFile
     IF SortSpec$ = "L" THEN
      GET 4, X
      IF TaxType$ = "R" THEN
       GET 1, CVS(mrec$), TBRECust(1)
       ELSE
       GET 1, CVS(mrec$), TBPPCust(1)
      END IF
      Account! = CVS(mrec$)
      ELSE
      IF TaxType$ = "R" THEN
      GET 1, X, TBRECust(1)
      ELSE
      GET 1, X, TBPPCust(1)
      END IF
      Account! = X
     END IF

3191 IF TaxType$ = "R" THEN
     IF TBRECust(1).Deleted = "Y" THEN 3230
     PRINT #10, ""
     PRINT #10, TBRECust(1).CustName
     PRINT #10, TBRECust(1).Address1
     PRINT #10, TBRECust(1).Address2
     PRINT #10, RTRIM$(TBRECust(1).Address3)
     PRINT #10, ""
     ELSE
     IF TBPPCust(1).Deleted = "Y" THEN 3230
     PRINT #10, ""
     PRINT #10, TBPPCust(1).CustName
     PRINT #10, TBPPCust(1).Address1
     PRINT #10, TBPPCust(1).Address2
     PRINT #10, TBPPCust(1).City; ", "; TBPPCust(1).State; " "; TBPPCust(1).Zip
     PRINT #10, ""
     END IF
     lc = lc + 6
     IF devspec$ = "S" THEN 3210 ELSE 3230
3210 IF lc >= 18 THEN LOCATE 23, 1: COLOR 11: PRINT "Press any key for next page (ESC-Quit)!"; : fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10 ELSE CLS : lc = 0: GOTO 3230
3230 NEXT X:
     IF devspec$ = "S" THEN LOCATE 23, 1: COLOR 11: PRINT "All done, press any key for menu "; : fl = 1: GOSUB 62000 ELSE PRINT #10, CHR$(12);
3240 CLOSE
3250 GOTO 10
Alignment:
LOCATE 20, 1: COLOR 11
PRINT "Print an Alignment Mask (Y/N)(ESC-Cancel Label Print) .. "; : fl = 1
GOSUB 62210
IF cf = 1 THEN CLOSE : GOTO 10
IF in$ = "Y" THEN
 PRINT #10, ""
 PRINT #10, STRING$(33, "X")
 PRINT #10, STRING$(33, "X")
 PRINT #10, STRING$(33, "X")
 PRINT #10, STRING$(33, "X")
 PRINT #10, ""
 GOTO Alignment
ELSE
 LOCATE 20, 1: PRINT STRING$(79, 32):
 LOCATE 1, 1: RETURN
END IF

4000 RUN "tbrep06"
5000 RUN "TBRep09"
6000 RUN "TBRep11"
7000 RUN "TBRep12"
8000 RUN "TBRep13"
9000 RUN "TBMPSORT"
10000 RUN "TBRep15"
11000 RUN "TBRep16"
12000 RUN "TBRep17"
13000 RUN "TBCRep13"
14000 RUN "TBVEHFIL"
15000 RUN "TBVEHREP"
16000 RUN "TBBoats"
17000 RUN "tbtrail"

18000 RUN "tbrepm2"

60000  IF TaxType$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBpath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBpath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         END IF
        RETURN

60025 TBBalFile = FREEFILE
      IF TaxType$ = "R" THEN
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
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 61 THEN cf = 13: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 62 THEN cf = 14: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 63 THEN cf = 15: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 64 THEN cf = 16: RETURN

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
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN

Namesearch:
 LOCATE 8, 1:
 COLOR 11
 PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : fl = 25
 GOSUB 62000
 IF cf = 1 THEN
  ggo = 0
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  RETURN
 END IF
 nn$ = in$
 nn = LEN(nn$)
 IF nn > 20 THEN nn = 20
 nn$ = LEFT$(nn$, nn)

 IF TaxType$ = "R" THEN
  TBIDXFile = FREEFILE
  OPEN TBpath$ + "TBRENAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS mrec$
  GOSUB 60000
  FOR ll = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, ll
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO n1
n3:
  NEXT ll
  FOR LLL = 8 TO 13: PRINT STRING$(79, 32): NEXT LLL
  ggo = 0
  RETURN

n1:
GET TBFile, CVS(mrec$), TBRECust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBRECust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(mrec$)
n2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO n2
IF in$ = "Q" THEN
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  ggo = 0
  RETURN
END IF
IF in$ = "Y" THEN
  ggo = 1: Record! = CVS(mrec$)
  FOR LLL = 8 TO 13:  LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  RETURN
END IF
IF in$ = "N" THEN GOTO n3
GOTO n2
END IF


  TBIDXFile = FREEFILE
  OPEN TBpath$ + "TBPPNAME.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBIDXFile LEN = 24
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS mrec$
  GOSUB 60000
  FOR ll = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, ll
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO PPn1
PPn3:
  NEXT ll
  LOCATE 8, 1: PRINT STRING$(79, 32)
  ggo = 0
  RETURN

PPn1:
GET TBFile, CVS(mrec$), TBPPCust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBPPCust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(mrec$)
PPn2:
LOCATE 13, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO PPn2
IF in$ = "Q" THEN
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  ggo = 0
  RETURN
END IF
IF in$ = "Y" THEN
  ggo = 1: Record! = CVS(mrec$)
  FOR LLL = 8 TO 13: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
  CLOSE
  RETURN
END IF
IF in$ = "N" THEN GOTO PPn3
GOTO PPn2

FigureBalance:
     CurBal# = 0: PastDue# = 0
     IF TaxType$ = "R" THEN
     CurBal# = TBRECust(1).CurYrTaxAmount + TBRECust(1).CurYrIntAmount + TBRECust(1).CurYrPenAmount + TBRECust(1).CurYrOthAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR ll = 1 TO 10
     TBRECustBal(1).OthAmt(ll) = 0
     PastDue# = PastDue# + TBRECustBal(1).TaxAmt(ll) + TBRECustBal(1).IntAmt(ll) + TBRECustBal(1).PenAmt(ll) + TBRECustBal(1).OthAmt(ll) + TBRECustBal(1).LateAmt(ll)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT ll
     ELSE
     CurBal# = TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax + TBPPCust(1).CurYrIntAmount + TBPPCust(1). _
CurYrPenAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR ll = 1 TO 10
     PastDue# = PastDue# + TBPPCustBal(1).PerPropTaxAmt(ll) + TBPPCustBal(1).FarmEquipAmt(ll) + TBPPCustBal(1).MachToolsAmt(ll) + TBPPCustBal(1).MerchCapAmt(ll) + TBPPCustBal(1).MobileHomeAmt(ll) + TBPPCustBal(1).LateFeeAmt(ll) + TBPPCustBal(1). _
IntAmt(ll) + TBPPCustBal(1).PenAmt(ll)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT ll


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

