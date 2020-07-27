10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbcustrep
30 REM Customer Report Menu
40 REM Last Revision Date: 8-24-87
50 REM by Nick Picerno
 
   DEFINT A-Z
   DECLARE FUNCTION QPTrim$ (A$)
   DECLARE FUNCTION Num2Date$ (Num%)
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
'   TBpath$ = "F:\"                    ' change to "F:\" Before compiling
  'TBpath$ = "c:\qb45\lc\"                    ' change to "F:\" Before compiling
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
200 PRINT TAB(28); "Tax Billing System Reports "
220 PRINT
230 PRINT TAB(5); "1. Print Master Customer Listing";
240  PRINT TAB(44); "10. Print Daily Collection Journal"
        PRINT
250 PRINT TAB(5); "2. Print Customer Balances";
     PRINT TAB(44); "11. Print Land Book"
260 PRINT
270 PRINT TAB(5); "3. Customer Account Inquiry"; TAB(44); "12. Print Personal Prop Book"

280 PRINT
290 PRINT TAB(5); "4. Print Account Valuation Listing"; TAB(44); "13. Print Alpha Customer Listing"

300 PRINT
302 PRINT TAB(5); "5. Print Mailing Labels"; TAB(44); "14. Print Map Order Customer Listing"
    PRINT
    PRINT TAB(5); "6. Print Property Record Sheets"; TAB(44); "15. Print Reassessment Proof"
    PRINT
    PRINT TAB(5); "7. Balance by Specific Tax Year"; TAB(44); "16. Print Reassessment Notices"
    PRINT
    PRINT TAB(5); "8. Collection System"; TAB(44); "17. Print Reassessment Book"
    PRINT
    PRINT TAB(5); "9. Print Personal Property Sheets"; TAB(44); "18. 2nd Page"
310 LOCATE 23, 1: COLOR 11: PRINT TAB(25); "Selection No. (ESC-Master Menu) : "; : fl = -2: GOSUB 62000
320 IF cf = 1 THEN CLOSE : RUN "tbmenu"
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

2000 REM
2001 Tot1# = 0: Tot2# = 0: TOT3# = 0
     FOR ll = 0 TO 9: TOTTAX#(ll) = 0: NEXT ll
2010 SCREN$ = "2.02.00": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 32: PRINT "Customer Balance Listing"
2020 LOCATE 4, 1
2023 LOCATE 4, 1: PRINT STRING$(79, 32): LOCATE 4, 1
2030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount or (L)ast Name (A,L or ESC-ABORT) ..... "; : fl = 1: GOSUB 62210
2040 IF cf = 1 THEN 10
2050 IF NOT (in$ = "A" OR in$ = "L") THEN 2023
2060 SortSpec$ = in$: LOCATE 4, 79: PRINT SortSpec$
2065 LOCATE 5, 1: COLOR 11
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-Cancel) ... "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "P" OR in$ = "A") THEN 2065
     TaxType$ = in$
     GOSUB 60000
     GOSUB 60025
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

2110 LOCATE 10, 1: COLOR 11: PRINT "Print Past Due by Year (Y/N) .. "; : fl = 1: GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN 2110
     brkdwn$ = in$
     LOCATE 11, 1: COLOR 11: PRINT STRING$(79, "="): LOCATE 13, 23: PRINT "Now Processing : Please be Patient!": COLOR 14
2120 ReportFile$ = "TBREBAL.PRN"
     RptHandle% = 10
     OPEN ReportFile$ FOR OUTPUT AS #RptHandle%
     GOSUB 2900
2170 FOR Cnt! = 1 TO EndofFile
     IF SortSpec$ = "L" THEN
      GET 4, Cnt!
      IF TaxType$ = "R" THEN
       GET TBFile, CVS(mrec$), TBRECust(1)
       GET TBBalFile, CVS(mrec$), TBRECustBal(1)
       ELSE
       GET TBFile, CVS(mrec$), TBPPCust(1)
       GET TBBalFile, CVS(mrec$), TBPPCustBal(1)
      END IF
       Account! = CVS(mrec$)
       GOTO 2190
     END IF

      IF TaxType$ = "R" THEN
      GET TBFile, Cnt!, TBRECust(1)
      GET TBBalFile, Cnt!, TBRECustBal(1)
      ELSE
      GET TBFile, Cnt!, TBPPCust(1)
      GET TBBalFile, Cnt!, TBPPCustBal(1)
      END IF
      Account! = Cnt!
2190 PastDue# = 0
     CurBal# = 0
     IF TaxType$ = "R" THEN
      IF TBRECust(1).Deleted = "Y" THEN 2230
      IF LEFT$(TBRECust(1).ClassCode, 1) = "7" THEN 2230
      ELSE
      IF TBPPCust(1).Deleted = "Y" THEN 2230
      END IF
     
     GOSUB FigureBalance
     IF PastDue# + CurBal# = 0 THEN 2230
     
   IF TaxType$ = "R" THEN
      PRINT #10, Account!; TAB(10); TBRECust(1).CustName;
       PRINT #10, TAB(54); USING "######.##"; CurBal#;
        PRINT #10, TAB(68); USING "######.##"; PastDue#
      lc = lc + 1
     IF brkdwn$ = "Y" THEN
     FOR ll = 1 TO 5
      TOTTAX#(ll) = TOTTAX#(ll) + TBRECustBal(1).TaxAmt(ll) + TBRECustBal(1).IntAmt(ll) + TBRECustBal(1).PenAmt(ll) + TBRECustBal(1).OthAmt(ll) + TBRECustBal(1).LateAmt(ll)
      PRINT #10, TAB(((ll - 1) * 15) + 1); USING "######.##"; TBRECustBal(1).TaxAmt(ll) + TBRECustBal(1).IntAmt(ll) + TBRECustBal(1).PenAmt(ll) + TBRECustBal(1).OthAmt(ll) + TBRECustBal(1).LateAmt(ll); : NEXT ll
     FOR ll = 6 TO 10
      TOTTAX#(ll) = TOTTAX#(ll) + TBRECustBal(1).TaxAmt(ll) + TBRECustBal(1).IntAmt(ll) + TBRECustBal(1).PenAmt(ll) + TBRECustBal(1).OthAmt(ll) + TBRECustBal(1).LateAmt(ll)
      PRINT #10, TAB(((ll - 6) * 15) + 1); USING "######.##"; (TBRECustBal(1).TaxAmt(ll) + TBRECustBal(1).IntAmt(ll) + TBRECustBal(1).PenAmt(ll) + TBRECustBal(1).OthAmt(ll) + TBRECustBal(1).LateAmt(ll)); : NEXT ll
      PRINT #10, "": lc = lc + 2
     END IF
     Tot1# = Tot1# + CurBal#: Tot2# = Tot2# + PastDue#
     TotCust! = TotCust! + 1
   ELSE
      PRINT #10, Account!; TAB(10); TBPPCust(1).CustName;
       PRINT #10, TAB(54); USING "######.##"; CurBal#;
        PRINT #10, TAB(68); USING "######.##"; PastDue#
      lc = lc + 1
     IF brkdwn$ = "Y" THEN
     FOR ll = 1 TO 5
      TOTTAX#(ll) = TOTTAX#(ll) + TBPPCustBal(1).PerPropTaxAmt(ll) + TBPPCustBal(1).FarmEquipAmt(ll) + TBPPCustBal(1).MachToolsAmt(ll) + TBPPCustBal(1).MerchCapAmt(ll) + TBPPCustBal(1).MobileHomeAmt(ll) + TBPPCustBal(1).LateFeeAmt(ll) + TBPPCustBal( _
1).IntAmt(ll) + TBPPCustBal(1).PenAmt(ll)
      PRINT #10, TAB(((ll - 1) * 15) + 1); USING "######.##"; TBPPCustBal(1).PerPropTaxAmt(ll) + TBPPCustBal(1).FarmEquipAmt(ll) + TBPPCustBal(1).MachToolsAmt(ll) + TBPPCustBal(1).MerchCapAmt(ll) + TBPPCustBal(1).MobileHomeAmt(ll) + TBPPCustBal(1). _
LateFeeAmt(ll) + TBPPCustBal(1).IntAmt(ll) + TBPPCustBal(1).PenAmt(ll); : NEXT ll
     FOR ll = 6 TO 10
      TOTTAX#(ll) = TOTTAX#(ll) + TBPPCustBal(1).PerPropTaxAmt(ll) + TBPPCustBal(1).FarmEquipAmt(ll) + TBPPCustBal(1).MachToolsAmt(ll) + TBPPCustBal(1).MerchCapAmt(ll) + TBPPCustBal(1).MobileHomeAmt(ll) + TBPPCustBal(1).LateFeeAmt(ll) + TBPPCustBal( _
1).IntAmt(ll) + TBPPCustBal(1).PenAmt(ll)
      PRINT #10, TAB(((ll - 6) * 15) + 1); USING "######.##"; (TBPPCustBal(1).PerPropTaxAmt(ll) + TBPPCustBal(1).FarmEquipAmt(ll) + TBPPCustBal(1).MachToolsAmt(ll) + TBPPCustBal(1).MerchCapAmt(ll) + TBPPCustBal(1).MobileHomeAmt(ll) + TBPPCustBal(1). _
LateFeeAmt(ll)) + TBPPCustBal(1).IntAmt(ll) + TBPPCustBal(1).PenAmt(ll); : NEXT ll
      PRINT #10, "": lc = lc + 2
     END IF
    T1# = T1# + TBPPCust(1).CurYrPersPropTax
    t2# = t2# + TBPPCust(1).CurYrFarmEquipTax
    t3# = t3# + TBPPCust(1).CurYrMachToolsTax
    t4# = t4# + TBPPCust(1).CurYrMerchCapTax
    t5# = t5# + TBPPCust(1).CurYrMobileHomeTax
    t6# = t6# + TBPPCust(1).CurYrLateFeeTax
    t7# = t7# + TBPPCust(1).CurYrIntAmount
    t8# = t8# + TBPPCust(1).CurYrPenAmount
     Tot1# = Tot1# + CurBal#: Tot2# = Tot2# + PastDue#
     TotCust! = TotCust! + 1
     END IF
2220 IF lc >= 55 THEN PRINT #10, CHR$(12); : GOSUB 2900
2230 NEXT Cnt!
       PRINT #10, STRING$(79, "-"):
        PRINT #10, "Totals   Customers:"; USING "#####,#"; TotCust!;
         PRINT #10, TAB(52); USING "#######,#.##"; Tot1#;
          PRINT #10, TAB(66); USING "#######,#.##"; Tot2#
        IF brkdwn$ = "Y" THEN 2232 ELSE 2235
2232 FOR ll = 1 TO 5: PRINT #10, TAB(((ll - 1) * 15) + 1); USING "######.##"; TOTTAX#(ll); : NEXT ll
     FOR ll = 6 TO 10: PRINT #10, TAB(((ll - 6) * 15) + 1); USING "######.##"; TOTTAX#(ll); : NEXT ll: PRINT #10, ""
2235 PRINT CHR$(12);
2240 CLOSE
      CLS
      LOCATE 1, 1
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      Header$ = "Tax Balance Report"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$

2250 GOTO 10

2900 p = p + 1
     PRINT #10, TAB(27); "Tax Customer Balance Listing": PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "####"; p
     PRINT #10, "":
     IF TaxType$ = "R" THEN
      PRINT #10, "Real Estate Customers"
      ELSE
      PRINT #10, "Personal Property Customers"
     END IF
2910 PRINT #10, "Acct #"; TAB(10); "Account Name"; TAB(54); "Current"; TAB(68); "Previous"
     IF brkdwn$ = "Y" THEN 2920 ELSE PRINT #10, STRING$(79, "="): lc = 2: RETURN
2920 FOR ll = 1 TO 5: PRINT #10, TAB(((ll - 1) * 15) + 1); "Year: "; VAL(CurrentYear$) - (ll); : NEXT ll
2930 FOR ll = 6 TO 10: PRINT #10, TAB(((ll - 6) * 15) + 1); "Year: "; VAL(CurrentYear$) - (ll); : NEXT ll
2940 PRINT #10, "": PRINT #10, STRING$(79, "="): lc = 4: RETURN


3000 REM
3010 SCREN$ = "2.03.00": CLS : GOSUB 63500
3020 COLOR 11: LOCATE 1, 32: PRINT "Customer Account Inquiry";
3025 LOCATE 3, 1: COLOR 11
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-End) ..... "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "P" OR in$ = "A") THEN 3025
     TaxType$ = UCASE$(in$)
3030 LOCATE 4, 1: COLOR 11
     PRINT "Enter Customer Account Number (TAB-Last Name : ESC-Menu) .. "; : fl = -6
     GOSUB 62000
     IF cf = 1 THEN
      CLOSE : GOTO 10
     END IF
     IF cf = 2 THEN
      GOSUB Namesearch
      IF ggo = 1 THEN 3075
      IF ggo = 0 THEN LOCATE 10, 1: PRINT "Account Not Found : Press Any Key to Continue "; : fl = 1: GOSUB 62000: GOTO 3000
     END IF
     Record! = VAL(in$)
3075 CLOSE : GOSUB 60000: GOSUB 60025
     IF TaxType$ = "R" THEN
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBRECust(1)) THEN 3030
       LOCK #TBFile, Record!
       GET TBFile, Record!, TBRECust(1)
       GET TBBalFile, Record!, TBRECustBal(1)
       IF TBRECust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 3000
       ELSE
       IF Record! < 1 OR Record! > LOF(TBFile) / LEN(TBPPCust(1)) THEN 3030
       LOCK #TBFile, Record!
       GET TBFile, Record!, TBPPCust(1)
       GET TBBalFile, Record!, TBPPCustBal(1)
       IF TBPPCust(1).Deleted = "Y" THEN GOSUB Deleted: GOTO 3000
     END IF

     IF TaxType$ = "R" THEN
       GOSUB DisplayReal
        ELSE
       GOSUB DisplayPersonal
     END IF
     CLOSE
     GOTO 3000
     
DisplayReal:
FOR ll = 3 TO 23: LOCATE ll, 1: PRINT STRING$(79, 32): NEXT ll
GOSUB FigureBalance
COLOR 11
LOCATE 3, 1: PRINT "Real Estate Customer"
PRINT "Account of ..."
COLOR 15
PRINT TBRECust(1).CustName; : COLOR 11: PRINT TAB(60); "Acct # "; : COLOR 15: PRINT Record!
PRINT TBRECust(1).Address1; TAB(40); TBRECust(1).Comment1
PRINT TBRECust(1).Address2; TAB(40); TBRECust(1).Comment2
PRINT TBRECust(1).Address3; TAB(40); TBRECust(1).Comment3
COLOR 11
PRINT STRING$(79, "-")
PRINT "      Map #: "; : COLOR 15: PRINT TBRECust(1).MapNumber
COLOR 11
PRINT " Land Value: "; : COLOR 15: PRINT USING "$$#######,#"; TBRECust(1).LandValue;
 COLOR 11
 PRINT TAB(40); "Reassessed Value: "; : COLOR 10: PRINT USING "$$######,#.##"; TBRECust(1).RELandValue
 COLOR 11
PRINT " Bldg Value: "; : COLOR 15: PRINT USING "$$#######,#"; TBRECust(1).BldgValue;
COLOR 11
 PRINT , TAB(40); "Reassessed Value: "; : COLOR 10: PRINT USING "$$######,#.##"; TBRECust(1).REBldgValue
COLOR 11
PRINT "Description: ": COLOR 15
PRINT TBRECust(1).PropDesc1; TAB(40); TBRECust(1).PropDesc2
PRINT TBRECust(1).PropDesc3; TAB(40); TBRECust(1).PropDesc4
PRINT
COLOR 11
LOCATE 16, 1: PRINT STRING$(79, "-")
PRINT "Total Account Balance: "; : COLOR 15: PRINT USING "$$#######,#.##"; CurBal# + PastDue#
COLOR 11
PRINT "          Current Due: "; : COLOR 15: PRINT USING "$$#######,#.##"; CurBal#
COLOR 11
PRINT "             Past Due: "; : COLOR 15: PRINT USING "$$#######,#.##"; PastDue#
COLOR 11
LOCATE 20, 1: PRINT STRING$(79, "-")
LOCATE 21, 1: PRINT "Press <F1>-Transaction History   <F2>-Tax Balance by Year   <ESC>-End  "; : fl = 1
GOSUB 62000
IF cf = 11 THEN GOSUB RealHistory: GOTO DisplayReal
IF cf = 12 THEN GOSUB RealBalance: GOTO DisplayReal
CLOSE
RETURN
RealBalance:
FOR ll = 3 TO 23: LOCATE ll, 1: PRINT STRING$(79, 32): NEXT ll
GOSUB 60025
GET TBBalFile, Record!, TBRECustBal(1)
LOCATE 3, 1: COLOR 11
PRINT "Account of ... "; TBRECust(1).CustName
PRINT
PRINT " Year"; TAB(10); "Tax Amount"; TAB(26); "Int't Amt"; TAB(40); "Pen'ty Amt"; TAB(63); " Balance"
PRINT STRING$(79, "-")
COLOR 15
PRINT VAL(CurrentYear$); TAB(10); USING "#####,#.##"; TBRECust(1).CurYrTaxAmount;
 PRINT TAB(25); USING "#####,#.##"; TBRECust(1).CurYrIntAmount;
  PRINT TAB(40); USING "#####,#.##"; TBRECust(1).CurYrPenAmount;
   PRINT TAB(60); USING "$$#####,#.##"; TBRECust(1).CurYrTaxAmount + TBRECust(1).CurYrIntAmount + TBRECust(1).CurYrPenAmount
FOR Pyr = 1 TO 10
PRINT VAL(CurrentYear$) - Pyr; TAB(10); USING "#####,#.##"; TBRECustBal(1).TaxAmt(Pyr);
 PRINT TAB(25); USING "#####,#.##"; TBRECustBal(1).IntAmt(Pyr);
  PRINT TAB(40); USING "#####,#.##"; TBRECustBal(1).PenAmt(Pyr);
   PRINT TAB(60); USING "$$#####,#.##"; TBRECustBal(1).TaxAmt(Pyr) + TBRECustBal(1).IntAmt(Pyr) + TBRECustBal(1).PenAmt(Pyr)
NEXT Pyr



LOCATE 23, 1: PRINT "Press <ENTER> to Continue "; : fl = 1
GOSUB 62000
CLOSE TBBalFile
RETURN

DisplayPersonal:
FOR ll = 3 TO 23: LOCATE ll, 1: PRINT STRING$(79, 32): NEXT ll
GOSUB FigureBalance
COLOR 11
LOCATE 3, 1: PRINT "Personal Property Customer"
PRINT "Account of ..."
COLOR 15
PRINT TBPPCust(1).CustName; : COLOR 11: PRINT TAB(60); "Acct # "; : COLOR 15: PRINT Record!
PRINT TBPPCust(1).Address1
PRINT TBPPCust(1).Address2
PRINT RTRIM$(TBPPCust(1).City); ", "; TBPPCust(1).State; "  "; TBPPCust(1).Zip
PRINT
COLOR 11: PRINT "Personal Property Value: "; : COLOR 15: PRINT USING "$$#######,#"; TBPPCust(1).PersProp
COLOR 11: PRINT "  Machinery/Tools Value: "; : COLOR 15: PRINT USING "$$#######,#"; TBPPCust(1).MachTools
COLOR 11: PRINT "      Mobile Home Value: "; : COLOR 15: PRINT USING "$$#######,#"; TBPPCust(1).MobileHome
COLOR 11: PRINT " Merchant Capital Value: "; : COLOR 15: PRINT USING "$$#######,#"; TBPPCust(1).MerchCap
COLOR 11: PRINT "   Farm Equipment Value: "; : COLOR 15: PRINT USING "$$#######,#"; TBPPCust(1).FarmEquip
COLOR 11
LOCATE 16, 1: PRINT STRING$(79, "-")
PRINT "Total Account Balance: "; : COLOR 15: PRINT USING "$$#######,#.##"; CurBal# + PastDue#
COLOR 11
PRINT "          Current Due: "; : COLOR 15: PRINT USING "$$#######,#.##"; CurBal#
COLOR 11
PRINT "             Past Due: "; : COLOR 15: PRINT USING "$$#######,#.##"; PastDue#
COLOR 11
LOCATE 20, 1: PRINT STRING$(79, "-")
LOCATE 21, 1: PRINT "Press <F1>-Transaction History   <F2>-Tax Balance by Year   <ESC>-End  "; : fl = 1
GOSUB 62000
IF cf = 11 THEN GOSUB PersHistory: GOTO DisplayPersonal
IF cf = 12 THEN GOSUB PersBalance: GOTO DisplayPersonal
CLOSE
RETURN
PersBalance:
FOR ll = 3 TO 23: LOCATE ll, 1: PRINT STRING$(79, 32): NEXT ll
GOSUB 60025
GET TBBalFile, Record!, TBPPCustBal(1)
LOCATE 3, 1: COLOR 11
PRINT "Account of ... "; TBPPCust(1).CustName
PRINT
PRINT " Year"; TAB(10); "Tax Amount"; TAB(26); "Int't Amt"; TAB(40); "Pen'ty Amt"; TAB(63); " Balance"
PRINT STRING$(79, "-")
COLOR 15
PRINT VAL(CurrentYear$); TAB(10); USING "#####,#.##"; TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax;
 PRINT TAB(25); USING "#####,#.##"; TBPPCust(1).CurYrIntAmount;
  PRINT TAB(40); USING "#####,#.##"; TBPPCust(1).CurYrPenAmount;
   PRINT TAB(60); USING "$$#####,#.##"; TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax + TBPPCust(1). _
CurYrIntAmount + TBPPCust(1).CurYrPenAmount
FOR Pyr = 1 TO 10
PRINT VAL(CurrentYear$) - Pyr; TAB(10); USING "#####,#.##"; TBPPCustBal(1).PerPropTaxAmt(Pyr) + TBPPCustBal(1).FarmEquipAmt(Pyr) + TBPPCustBal(1).MachToolsAmt(Pyr) + TBPPCustBal(1).MerchCapAmt(Pyr) + TBPPCustBal(1).MobileHomeAmt(Pyr) + TBPPCustBal(1 _
).LateFeeAmt(Pyr);
 PRINT TAB(25); USING "#####,#.##"; TBPPCustBal(1).IntAmt(Pyr);
  PRINT TAB(40); USING "#####,#.##"; TBPPCustBal(1).PenAmt(Pyr);
   PRINT TAB(60); USING "$$#####,#.##"; TBPPCustBal(1).PerPropTaxAmt(Pyr) + TBPPCustBal(1).FarmEquipAmt(Pyr) + TBPPCustBal(1).MachToolsAmt(Pyr) + TBPPCustBal(1).MerchCapAmt(Pyr) + TBPPCustBal(1).MobileHomeAmt(Pyr) + TBPPCustBal(1).LateFeeAmt(Pyr) +  _
TBPPCustBal(1).IntAmt(Pyr) + TBPPCustBal(1).PenAmt(Pyr)
NEXT Pyr
LOCATE 23, 1: PRINT "Press <ENTER> to Continue or <F5>-Down : <F6>-Up "; : fl = 1
GOSUB 62000
IF cf = 15 THEN
 LOCATE 23, 1: PRINT STRING$(79, 32)
 LOCATE 23, 1: COLOR 11
 PRINT "Enter Your Initials Please .. "; : fl = 3
 GOSUB 62000
 Initials$ = UCASE$(in$)
 OPEN "TBCLOG" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = 64
 FIELD 14, 3 AS init$, 10 AS cdate$, 30 AS who$, 4 AS ARecord$
 LSET cdate$ = DATE$
 LSET init$ = Initials$
 LSET who$ = TBPPCust(1).CustName
 LSET ARecord$ = MKS$(Record!)
 PUT 14, LOF(14) / 64 + 1
 CLOSE 14
     
 FOR ll = 10 TO 2 STEP -1
  TBPPCustBal(1).PerPropTaxAmt(ll) = TBPPCustBal(1).PerPropTaxAmt(ll - 1)
  TBPPCustBal(1).FarmEquipAmt(ll) = TBPPCustBal(1).FarmEquipAmt(ll - 1)
  TBPPCustBal(1).MachToolsAmt(ll) = TBPPCustBal(1).MachToolsAmt(ll - 1)
  TBPPCustBal(1).MerchCapAmt(ll) = TBPPCustBal(1).MerchCapAmt(ll - 1)
  TBPPCustBal(1).MobileHomeAmt(ll) = TBPPCustBal(1).MobileHomeAmt(ll - 1)
  TBPPCustBal(1).LateFeeAmt(ll) = TBPPCustBal(1).LateFeeAmt(ll - 1)
  TBPPCustBal(1).IntAmt(ll) = TBPPCustBal(1).IntAmt(ll - 1)
  TBPPCustBal(1).PenAmt(ll) = TBPPCustBal(1).PenAmt(ll - 1)
 NEXT ll
 ll = 1
  TBPPCustBal(1).PerPropTaxAmt(ll) = TBPPCust(1).CurYrPersPropTax
  TBPPCustBal(1).FarmEquipAmt(ll) = TBPPCust(1).CurYrFarmEquipTax
  TBPPCustBal(1).MachToolsAmt(ll) = TBPPCust(1).CurYrMachToolsTax
  TBPPCustBal(1).MerchCapAmt(ll) = TBPPCust(1).CurYrMerchCapTax
  TBPPCustBal(1).MobileHomeAmt(ll) = TBPPCust(1).CurYrMobileHomeTax
  TBPPCustBal(1).LateFeeAmt(ll) = TBPPCust(1).CurYrLateFeeTax
  TBPPCustBal(1).IntAmt(ll) = TBPPCust(1).CurYrIntAmount
  TBPPCustBal(1).PenAmt(ll) = TBPPCust(1).CurYrPenAmount
  TBPPCust(1).CurYrPersPropTax = 0
  TBPPCust(1).CurYrFarmEquipTax = 0
  TBPPCust(1).CurYrMachToolsTax = 0
  TBPPCust(1).CurYrMerchCapTax = 0
  TBPPCust(1).CurYrMobileHomeTax = 0
  TBPPCust(1).CurYrLateFeeTax = 0
  TBPPCust(1).CurYrIntAmount = 0
  TBPPCust(1).CurYrPenAmount = 0
  PUT TBFile, Record!, TBPPCust(1)
  PUT TBBalFile, Record!, TBPPCustBal(1)
  GOTO PersBalance

END IF
IF cf = 16 THEN
 LOCATE 23, 1: PRINT STRING$(79, 32)
 LOCATE 23, 1: COLOR 11
 PRINT "Enter Your Initials Please .. "; : fl = 3
 GOSUB 62000
 Initials$ = UCASE$(in$)
 OPEN "TBCLOG" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = 64
 FIELD 14, 3 AS init$, 10 AS cdate$, 30 AS who$, 4 AS ARecord$
 LSET cdate$ = DATE$
 LSET init$ = Initials$
 LSET who$ = TBPPCust(1).CustName
 LSET ARecord$ = MKS$(Record!)
 PUT 14, LOF(14) / 64 + 1
 CLOSE 14
  ll = 1
  TBPPCust(1).CurYrPersPropTax = TBPPCustBal(1).PerPropTaxAmt(ll)
  TBPPCust(1).CurYrFarmEquipTax = TBPPCustBal(1).FarmEquipAmt(ll)
  TBPPCust(1).CurYrMachToolsTax = TBPPCustBal(1).MachToolsAmt(ll)
  TBPPCust(1).CurYrMerchCapTax = TBPPCustBal(1).MerchCapAmt(ll)
  TBPPCust(1).CurYrMobileHomeTax = TBPPCustBal(1).MobileHomeAmt(ll)
  TBPPCust(1).CurYrLateFeeTax = TBPPCustBal(1).LateFeeAmt(ll)
  TBPPCust(1).CurYrIntAmount = TBPPCustBal(1).IntAmt(ll)
  TBPPCust(1).CurYrPenAmount = TBPPCustBal(1).PenAmt(ll)
  FOR ll = 1 TO 9
  TBPPCustBal(1).PerPropTaxAmt(ll) = TBPPCustBal(1).PerPropTaxAmt(ll + 1)
  TBPPCustBal(1).FarmEquipAmt(ll) = TBPPCustBal(1).FarmEquipAmt(ll + 1)
  TBPPCustBal(1).MachToolsAmt(ll) = TBPPCustBal(1).MachToolsAmt(ll + 1)
  TBPPCustBal(1).MerchCapAmt(ll) = TBPPCustBal(1).MerchCapAmt(ll + 1)
  TBPPCustBal(1).MobileHomeAmt(ll) = TBPPCustBal(1).MobileHomeAmt(ll + 1)
  TBPPCustBal(1).LateFeeAmt(ll) = TBPPCustBal(1).LateFeeAmt(ll + 1)
  TBPPCustBal(1).IntAmt(ll) = TBPPCustBal(1).IntAmt(ll + 1)
  TBPPCustBal(1).PenAmt(ll) = TBPPCustBal(1).PenAmt(ll + 1)
 NEXT ll
 PUT TBFile, Record!, TBPPCust(1)
  PUT TBBalFile, Record!, TBPPCustBal(1)
  GOTO PersBalance

END IF
CLOSE TBBalFile
RETURN

     
RealHistory:
GOSUB 61000
GOSUB RealHeading
TR! = TBRECust(1).FirstTrans
GOSUB RealTransLink
IF Counter = 0 THEN
 LOCATE 10, 1: COLOR 15
 PRINT "No History For This Account : Press <ENTER> to Continue "; : fl = 1
 GOSUB 62000
 CLOSE TBTransFile
 RETURN
END IF
COLOR 15
FOR Cnt! = Counter TO 1 STEP -1
GET TBTransFile, TransLink!(Cnt!), TBTrans(1)
PRINT Num2Date$(TBTrans(1).TransDate);
 PRINT TAB(12); TBTrans(1).TrDesc;
  PRINT TAB(30); "";
  IF TBTrans(1).TrType = 1 THEN PRINT "Tax Bill";
  IF TBTrans(1).TrType = 2 THEN PRINT "Payment";
  IF TBTrans(1).TrType = 3 THEN PRINT "Interest";
  IF TBTrans(1).TrType = 4 THEN PRINT "Penalty";
  IF TBTrans(1).TrType = 5 THEN PRINT "Release";
  IF TBTrans(1).TrType = 6 THEN PRINT "Beg Bal";
  IF TBTrans(1).TrType < 1 OR TBTrans(1).TrType > 6 THEN PRINT "Undef";
 PRINT TAB(40); USING "$$#####,#.##"; TBTrans(1).TransAmt;
  PRINT TAB(55); USING "$$#####,#.##"; TBTrans(1).TrRunningBal
  NEXT Cnt!
  LOCATE 23, 1: COLOR 11
  PRINT "Press <ENTER> to Continue! "; : fl = 1
  GOSUB 62000
  CLOSE TBTransFile
  RETURN

RealHeading:
COLOR 11
CLS
PRINT TAB(26); "Tax Customer - Account History"
PRINT TAB(26); "  * Real Estate Customer *"; TAB(60); "Date: "; DATE$
PRINT
PRINT " Account of: "; : COLOR 15: PRINT TBRECust(1).CustName
COLOR 11
PRINT "Account Bal: "; : COLOR 15: PRINT USING "$$######,#.##"; CurBal# + PastDue#
COLOR 11
PRINT
PRINT "Tr Date"; TAB(12); "Description"; TAB(30); "Tr Type"; TAB(43); "Tr Amount"; TAB(60); "Run Balance"
PRINT STRING$(79, "-")
RETURN

RealTransLink:
 Counter = 0
Realtrlink1:
 IF TR! = 0 THEN RETURN
 Counter = Counter + 1
 TransLink!(Counter) = TR!
 GET TBTransFile, TR!, TBTrans(1)
 TR! = TBTrans(1).TrNextTrans
 GOTO Realtrlink1

PersHistory:
GOSUB 61000
GOSUB PersHeading
TR! = TBPPCust(1).FirstTrans
GOSUB PersTransLink
IF Counter = 0 THEN
 LOCATE 10, 1: COLOR 15
 PRINT "No History For This Account : Press <ENTER> to Continue "; : fl = 1
 GOSUB 62000
 CLOSE TBTransFile
 RETURN
END IF
COLOR 15
FOR Cnt! = Counter TO 1 STEP -1
GET TBTransFile, TransLink!(Cnt!), TBTrans(1)
PRINT Num2Date$(TBTrans(1).TransDate);
 PRINT TAB(12); TBTrans(1).TrDesc;
  PRINT TAB(30); "";
  IF TBTrans(1).TrType = 1 THEN PRINT "Tax Bill";
  IF TBTrans(1).TrType = 2 THEN PRINT "Payment";
  IF TBTrans(1).TrType = 3 THEN PRINT "Interest";
  IF TBTrans(1).TrType = 4 THEN PRINT "Penalty";
  IF TBTrans(1).TrType = 5 THEN PRINT "Release";
  IF TBTrans(1).TrType = 6 THEN PRINT "Beg Bal";
  IF TBTrans(1).TrType < 1 OR TBTrans(1).TrType > 6 THEN PRINT "Undef";
 PRINT TAB(40); USING "$$#####,#.##"; TBTrans(1).TransAmt;
  PRINT TAB(59); USING "**$#####,#.##"; TBTrans(1).TrRunningBal
  NEXT Cnt!
  LOCATE 23, 1: COLOR 11
  PRINT "Press <ENTER> to Continue! "; : fl = 1
  GOSUB 62000
  CLOSE TBTransFile
  RETURN
  RETURN

PersHeading:
COLOR 11
CLS
PRINT TAB(26); "Tax Customer - Account History"
PRINT TAB(26); "  * Pers Property Customer *"; TAB(60); "Date: "; DATE$
PRINT
PRINT " Account of: "; : COLOR 15: PRINT TBPPCust(1).CustName
COLOR 11
PRINT "Account Bal: "; : COLOR 15: PRINT USING "$$######,#.##"; CurBal# + PastDue#
COLOR 11
PRINT
PRINT "Tr Date"; TAB(12); "Description"; TAB(30); "Tr Type"; TAB(43); "Tr Amount"; TAB(60); "Run Balance"
PRINT STRING$(79, "-")
RETURN

PersTransLink:
 Counter = 0
Perstrlink1:
 IF TR! = 0 THEN RETURN
 Counter = Counter + 1
 TransLink!(Counter) = TR!
 GET TBTransFile, TR!, TBTrans(1)
 TR! = TBTrans(1).TrNextTrans
 GOTO Perstrlink1

4000 SCREN$ = "2.01.00": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 32: PRINT "Customer Account Listing"
     CustCount = 0
     page = 0
4030 LOCATE 4, 1: COLOR 11: PRINT "Sort Specification: Sort by (A)ccount # or (L)ast Name (A,L  or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
4040 IF cf = 1 THEN 10
4050 IF NOT (in$ = "A" OR in$ = "L") THEN 4030
4060 SortSpec$ = in$: LOCATE 4, 79: PRINT SortSpec$
4070 LOCATE 6, 1: COLOR 11
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-End) ..... "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "P" OR in$ = "A") THEN 4070
     TaxType$ = UCASE$(in$)
4080 COLOR 11: LOCATE 8, 1
     PRINT "Output Specification: Output to (S)creen or (P)rinter (S,P or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
4090 IF cf = 1 THEN 10
4100 IF NOT (in$ = "S" OR in$ = "P") THEN 4080
4105 devspec$ = in$: LOCATE 8, 77: PRINT devspec$
4110 LOCATE 10, 1: COLOR 11: PRINT STRING$(79, "="): COLOR 10: LOCATE 13, 23: PRINT "Now Processing : Please be Patient!": COLOR 14
4120 IF devspec$ = "S" THEN dev$ = "SCRN:" ELSE dev$ = "lpt1:"
4130 OPEN "O", 10, dev$: WIDTH #10, 255
4140 GOSUB 60000
     IF devspec$ = "S" THEN CLS : GOSUB 4900 ELSE GOSUB 4900
4160 IF SortSpec$ = "A" THEN
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
4170 FOR X = 1 TO EndofFile
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

4191 IF TaxType$ = "R" THEN
     IF TBRECust(1).Deleted = "Y" THEN 4230
     PRINT #10, Account!;
      PRINT #10, TAB(10); LEFT$(TBRECust(1).CustName, 35);
      PRINT #10, TAB(48); USING "#######,#"; TBRECust(1).LandValue;
      PRINT #10, TAB(58); USING "#######,#"; TBRECust(1).BldgValue;
      PRINT #10, TAB(68); USING "#######,#"; TBRECust(1).LandValue + TBRECust(1).BldgValue
      TotalLand# = TotalLand# + TBRECust(1).LandValue
      TotalBldg# = TotalBldg# + TBRECust(1).BldgValue
      lc = lc + 1
     ELSE
      IF TBPPCust(1).Deleted = "Y" THEN 4230
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
     IF devspec$ = "S" THEN 4210 ELSE 4220
4210 IF lc >= 17 THEN LOCATE 23, 1: COLOR 11: PRINT "Press any key for next page (ESC-Quit)!"; : fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10 ELSE CLS : GOSUB 4900: GOTO 4230
4220 IF lc >= 55 THEN PRINT #10, CHR$(12); : GOSUB 4900
4230 NEXT X
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

4900 p = p + 1
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
5000 SCREN$ = "2.05.00": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 32: PRINT "Customer Account Listing"
5020 LOCATE 3, 1
5030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount # or (L)ast Name (A,L  or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
5040 IF cf = 1 THEN 10
5050 IF NOT (in$ = "A" OR in$ = "L" OR in$ = "M") THEN 5020
5060 SortSpec$ = in$: LOCATE 4, 79: PRINT SortSpec$
5070 LOCATE 6, 1: COLOR 11
     PRINT "(R)eal or (P)ersonal (R,P or <ESC>-End) ..... "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "P" OR in$ = "A") THEN 5070
     TaxType$ = UCASE$(in$)


5080 COLOR 11: LOCATE 8, 1
     PRINT "Output Specification: Output to (S)creen or (P)rinter (S,P or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
5090 IF cf = 1 THEN 10
5100 IF NOT (in$ = "S" OR in$ = "P") THEN 5080
5105 devspec$ = in$: LOCATE 8, 77: PRINT devspec$
5110 LOCATE 10, 1: COLOR 11: PRINT STRING$(79, "="): COLOR 10: LOCATE 13, 23: PRINT "Now Processing : Please be Patient!": COLOR 14
5120 IF devspec$ = "S" THEN dev$ = "SCRN:" ELSE dev$ = "lpt1:"
5130 OPEN "O", 10, dev$: WIDTH #10, 255
5140 GOSUB 60000: IF devspec$ = "S" THEN CLS
     GOSUB Alignment
5160 IF SortSpec$ = "A" THEN
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
5170 FOR X = 1 TO EndofFile
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

5191 IF TaxType$ = "R" THEN
     IF TBRECust(1).Deleted = "Y" THEN 5230
     PRINT #10, ""
     PRINT #10, TBRECust(1).CustName
     PRINT #10, TBRECust(1).Address1
     PRINT #10, TBRECust(1).Address2
     PRINT #10, RTRIM$(TBRECust(1).Address3)
     PRINT #10, ""
     ELSE
     IF TBPPCust(1).Deleted = "Y" THEN 5230
     PRINT #10, ""
     PRINT #10, TBPPCust(1).CustName
     PRINT #10, TBPPCust(1).Address1
     PRINT #10, TBPPCust(1).Address2
     PRINT #10, TBPPCust(1).City; ", "; TBPPCust(1).State; " "; TBPPCust(1).Zip
     PRINT #10, ""
     END IF
     lc = lc + 6
     IF devspec$ = "S" THEN 5210 ELSE 5230
5210 IF lc >= 18 THEN LOCATE 23, 1: COLOR 11: PRINT "Press any key for next page (ESC-Quit)!"; : fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10 ELSE CLS : lc = 0: GOTO 5230
5230 NEXT X:
     IF devspec$ = "S" THEN LOCATE 23, 1: COLOR 11: PRINT "All done, press any key for menu "; : fl = 1: GOSUB 62000 ELSE PRINT #10, CHR$(12);
5240 CLOSE
5250 GOTO 10
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

6000 RUN "tbrep06"

7000 REM
     Tot1# = 0: Tot2# = 0
7010 SCREN$ = "2.07.00": CLS : GOSUB 63500: COLOR 11
     LOCATE 1, 32: PRINT "Specific Tax Yr Balance"
7020 LOCATE 3, 1: PRINT STRING$(79, 32): LOCATE 3, 1
7030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount or (L)ast Name (A,L or ESC-ABORT) ..... "; : fl = 1: GOSUB 62210
7040 IF cf = 1 THEN 10
7050 IF NOT (in$ = "A" OR in$ = "L") THEN 7020
7060 SortSpec$ = in$: LOCATE 3, 79: PRINT SortSpec$
7065 LOCATE 5, 1: COLOR 11
    PRINT "(R)eal or (P)ersonal (R,P or <ESC>-Cancel) ... "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "P" OR in$ = "A") THEN 7065
     TaxType$ = in$
     GOSUB 60000
7160 IF SortSpec$ = "A" THEN
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

7070 LOCATE 6, 1

7110 LOCATE 10, 1: COLOR 11: PRINT "Enter Year to Print (XXXX) .. "; : fl = -4
     GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10
     IF LEN(in$) <> 4 THEN 7110
     YearToPrint$ = in$
     CurrentYear = VAL(CurrentYear$)
     YearToPrint = VAL(YearToPrint$)
     IF YearToPrint > CurrentYear THEN
      LOCATE 12, 1: COLOR 15
      PRINT CHR$(7);
      PRINT "Year Entered Exceeds Current Year on System : Current is "; CurrentYear$
      PRINT "Press <ENTER> to Try Again : <ESC> to End: ";
      fl = 1
      GOSUB 62000
      IF cf = 1 THEN CLOSE : GOTO 10
      GOTO 7000
     END IF
     GrabYear = CurrentYear - YearToPrint
     LOCATE 11, 1: COLOR 11: PRINT STRING$(79, "="): LOCATE 13, 23: PRINT "Now Processing : Please be Patient!": COLOR 14
     ReportFile$ = "TBREBAL.PRN"
     RptHandle% = 10
     OPEN ReportFile$ FOR OUTPUT AS #RptHandle%
     GOSUB 7900

7140 GOSUB 60000
     GOSUB 60025
7170 FOR X = 1 TO EndofFile
     IF SortSpec$ = "L" THEN
      GET 4, X
      IF TaxType$ = "R" THEN
       GET TBFile, CVS(mrec$), TBRECust(1)
       GET TBBalFile, CVS(mrec$), TBRECustBal(1)
       ELSE
       GET TBFile, CVS(mrec$), TBPPCust(1)
       GET TBBalFile, CVS(mrec$), TBPPCustBal(1)
      END IF
       Account! = CVS(mrec$)
      ELSE
      IF TaxType$ = "R" THEN
      GET TBFile, X, TBRECust(1)
      GET TBBalFile, X, TBRECustBal(1)
      ELSE
      GET TBFile, X, TBPPCust(1)
      GET TBBalFile, X, TBPPCustBal(1)
      END IF
      Account! = X
     END IF

7190 PastDue# = 0: CurBal# = 0: Tax# = 0: Interest# = 0
     GOSUB FigureBalance
     IF TaxType$ = "R" THEN
      IF TBRECust(1).Deleted = "Y" THEN 7230
      ELSE
      IF TBPPCust(1).Deleted = "Y" THEN 7230
     END IF
     IF PastDue# + CurBal# = 0 THEN 7230
     IF TaxType$ = "R" THEN
      IF LEFT$(TBRECust(1).ClassCode, 1) = "7" THEN 7230
     END IF

   IF TaxType$ = "R" THEN
      IF GrabYear > 0 THEN
        Tax# = TBRECustBal(1).TaxAmt(GrabYear) + TBRECustBal(1).LateAmt(GrabYear)
        Interest# = TBRECustBal(1).IntAmt(GrabYear) + TBRECustBal(1).PenAmt(GrabYear) + TBRECustBal(1).OthAmt(GrabYear)
      ELSE
        Tax# = TBRECust(1).CurYrTaxAmount
        Interest# = TBRECust(1).CurYrPenAmount + TBRECust(1).CurYrOthAmount + TBRECust(1).CurYrIntAmount
      END IF
       IF Tax# <> 0 OR Interest# <> 0 THEN
       PRINT #10, Account!; TAB(10); TBRECust(1).CustName;
        PRINT #10, TAB(50); USING "#####.##"; Tax#;
         PRINT #10, TAB(60); USING "#####.##"; Interest#;
          PRINT #10, TAB(70); USING "#####.##"; Tax# + Interest#
         lc = lc + 1
         Tot1# = Tot1# + Tax#: Tot2# = Tot2# + Interest#
         END IF
   ELSE
      
      IF GrabYear > 0 THEN
       Tax# = TBPPCustBal(1).PerPropTaxAmt(GrabYear) + TBPPCustBal(1).FarmEquipAmt(GrabYear) + TBPPCustBal(1).MerchCapAmt(GrabYear) + TBPPCustBal(1).MobileHomeAmt(GrabYear) + TBPPCustBal(1).MachToolsAmt(GrabYear) + TBPPCustBal(1).LateFeeAmt(GrabYear _
)
       Interest# = TBPPCustBal(1).PenAmt(GrabYear)

      ELSE
       Tax# = TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMobileHomeTax
       Interest# = TBPPCust(1).CurYrLateFeeTax + TBPPCust(1).CurYrPenAmount + TBPPCust(1).CurYrIntAmount
      END IF
      IF Tax# <> 0 OR Interest# <> 0 THEN
       PRINT #10, Account!; TAB(10); TBPPCust(1).CustName;
        PRINT #10, TAB(50); USING "#####.##"; Tax#;
         PRINT #10, TAB(60); USING "#####.##"; Interest#;
          PRINT #10, TAB(70); USING "#####.##"; Tax# + Interest#
        lc = lc + 1
        Tot1# = Tot1# + Tax#: Tot2# = Tot2# + Interest#
       END IF
    END IF
7200 REM
7220 IF lc >= 55 THEN PRINT #10, CHR$(12); : GOSUB 7900
7230 NEXT X
       PRINT #10, STRING$(79, "-"):
        PRINT #10, "Total Tax ...... "; USING "$$#######,#.##"; Tot1#
        PRINT #10, "Total Int/Pen .. "; USING "$$#######,#.##"; Tot2#
        PRINT #10, "Grand Total .... "; USING "$$#######,#.##"; Tot1# + Tot2#
       PRINT #10, CHR$(12);
7240 CLOSE
      CLS
      LOCATE 1, 1
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      Header$ = "Tax Balance Report"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
7250 GOTO 10

7900 p = p + 1
     PRINT #10, TAB(23); "Tax Customer Balance Listing by Year": PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "####"; p
     PRINT #10, "Year: "; YearToPrint$
     IF TaxType$ = "R" THEN
      PRINT #10, "Real Estate Customers"
      ELSE
      PRINT #10, "Personal Property Customers"
     END IF
7910 PRINT #10, "Acct #"; TAB(10); "Account Name"; TAB(50); "Tax Owed"; TAB(60); "Int/Pen"; TAB(70); "Total Due"
7940 PRINT #10, STRING$(79, "="): lc = 5: RETURN
8000 RUN "TBDELNOT"
9000 RUN "TBRep09"
10000 RUN "TBRep10"
11000 RUN "TBRep11"
12000 RUN "TBRep12"
13000 RUN "TBRep13"
14000 RUN "TBMPSORT"
15000 RUN "TBRep15"
16000 RUN "TBRep16"
17000 RUN "TBRep17"
18000 RUN "tbRepM2"
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

