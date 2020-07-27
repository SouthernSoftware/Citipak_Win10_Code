10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbccustrp
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
   TBpath$ = "F:\"                    ' change to "F:\" Before compiling
   
   CLOSE
   p = 0: Page = 0
   OPEN "R", 1, TBpath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
   CurrentYear$ = year$
   CLOSE 1
   'ON ERROR GOTO networkerror
   WIDTH LPRINT 136

     REDIM TransLink!(500)
     REDIM TBRECust(1) AS TBRECust
     REDIM TBRECustBal(1) AS TBREBAL
     REDIM TBPPCust(1) AS TBPPCust
     REDIM TBPPCustBal(1) AS TBPPBal
     REDIM TBTrans(1) AS TBTrans

     
100 REM begin program
110 scren$ = "2.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500
130 LOCATE 4, 1: COLOR 11
200 PRINT TAB(21); "Tax Billing System Reports 2nd Page Menu"
220 PRINT : PRINT
230 PRINT TAB(5); "1. Print Business Listing";
     PRINT TAB(42); "7. Print New Bldg Book (Pre-Bill)"
    PRINT
250 PRINT TAB(5); "2. Print Pers Prop - Farm Equip";
     PRINT TAB(42); "8. High Mileage Vehicle Report"
260 PRINT
270 PRINT TAB(5); "3. Print Pers Prop - Mach/Tools";
     PRINT TAB(42); "9. Print Real Estate Edited Accts "
280 PRINT
290 PRINT TAB(5); "4. Print Pers Prop - Merch Cap";
     PRINT TAB(41); "10. Print Pers Prop Cust List"
300 PRINT
302 PRINT TAB(5); "5. Print Pers Prop - Mobile Homes";
     PRINT TAB(41); "11. Print Mortgage Customer Report"
     PRINT
    PRINT TAB(5); "6. Print New Bldg Listing";
     PRINT TAB(41); "12. Print Real Estate Agent Report"
    
310 LOCATE 20, 1: COLOR 11
     PRINT TAB(21); "Selection No. (ESC-Main Reports Menu) : "; : fl = -2: GOSUB 62000
320 IF cf = 1 THEN CLOSE : RUN "tbccustr"
325 IF LEN(in$) = 0 THEN 310
330 IF VAL(in$) < 1 OR VAL(in$) > 12 THEN 310
340 ON VAL(in$) GOTO 1000, 2000, 3000, 4000, 5000, 6000, 7000, 8000, 9000, 10000, 11000, 12000

1000 REM
     CustCount = 0
1010 scren$ = "2.02.01": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 28: PRINT "Customer Listing : Business Accts"
1020 LOCATE 4, 1
1030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount # or (L)ast Name (A,L  or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
1040 IF cf = 1 THEN 10
1050 IF NOT (in$ = "A" OR in$ = "L" OR in$ = "M") THEN 1020
1060 SortSpec$ = in$: LOCATE 4, 79: PRINT SortSpec$
1070 TaxType$ = "P"
     ReportFile$ = "TBBUSLST.PRN"
     RptHandle% = 10
     OPEN ReportFile$ FOR OUTPUT AS #RptHandle%
     GOSUB 1900
1140 GOSUB 60000
1150 IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10
1160 IF SortSpec$ = "A" THEN
       EndofFile = LOF(TBFile) / LEN(TBPPCust(1))
     END IF

     IF SortSpec$ = "L" THEN
        OPEN TBpath$ + "TBPPname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS MREC$: EndofFile = LOF(4) / 24
     END IF
1170 FOR x = 1 TO EndofFile
     IF SortSpec$ = "L" THEN
      GET 4, x
      GET 1, CVS(MREC$), TBPPCust(1)
      Account! = CVS(MREC$)
        ELSE
      GET 1, x, TBPPCust(1)
     END IF
      Account! = x
     
     
1191 IF TBPPCust(1).Business = "Y" THEN
     IF TBPPCust(1).Deleted = "Y" THEN 1230
     Address$ = RTRIM$(TBPPCust(1).Address1) + "  " + RTRIM$(TBPPCust(1).Address2)

     Address2$ = RTRIM$(TBPPCust(1).City) + " " + RTRIM$(TBPPCust(1).State) + "  " + RTRIM$(TBPPCust(1).Zip)
     PRINT #10, Account!;
     PRINT #10, TAB(10); TBPPCust(1).CustName;
     PRINT #10, TAB(40); LEFT$(LTRIM$(Address$), 43); TAB(95); Address2$
     lc = lc + 1
     CustCount = CustCount + 1
     END IF
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
1220 IF lc >= 58 THEN PRINT #10, CHR$(12); : GOSUB 1900
1230 NEXT x
     PRINT #10,
     PRINT #10, "Total Customers .. "; USING "######,#"; CustCount
1240 CLOSE
      CLS
      LOCATE 1, 1
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      Header$ = "Tax Account Listing"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$

1250 GOTO 10

1900 p = p + 1
     PRINT #10, TAB(27); "Tax Customer Business Account Listing"
     PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "####"; p
     
     
1910 PRINT #10, "Acct. #"; TAB(10); "Name"; TAB(40); "Address"
     PRINT #10, STRING$(130, "="): lc = 5: RETURN
2000 scren$ = "2.02.02": CLS : GOSUB 63500: COLOR 11
     CustCount = 0: TotalFarm# = 0
       LOCATE 1, 32: PRINT "Customer Listing : Farm Equip"
2020 LOCATE 4, 1
2030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount # or (L)ast Name (A,L  or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
2040 IF cf = 1 THEN 10
2050 IF NOT (in$ = "A" OR in$ = "L" OR in$ = "M") THEN 2020
2060 SortSpec$ = in$: LOCATE 4, 79: PRINT SortSpec$
2070 TaxType$ = "P"
     ReportFile$ = "TBPPFE.PRN"
     RptHandle% = 10
     OPEN ReportFile$ FOR OUTPUT AS #RptHandle%
     GOSUB 2900
2140 GOSUB 60000
2150 IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10
2160 IF SortSpec$ = "A" THEN
       EndofFile = LOF(TBFile) / LEN(TBPPCust(1))
     END IF

     IF SortSpec$ = "L" THEN
        OPEN TBpath$ + "TBPPname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS MREC$: EndofFile = LOF(4) / 24
     END IF
2170 FOR x = 1 TO EndofFile
     IF SortSpec$ = "L" THEN
      GET 4, x
      GET 1, CVS(MREC$), TBPPCust(1)
      Account! = CVS(MREC$)
        ELSE
      GET 1, x, TBPPCust(1)
     END IF
      Account! = x


2191 IF TBPPCust(1).FarmEquip > 0 THEN
     IF TBPPCust(1).Deleted = "Y" THEN 2230
     Address$ = RTRIM$(TBPPCust(1).Address1) + "  " + RTRIM$(TBPPCust(1).Address2)

     Address2$ = RTRIM$(TBPPCust(1).City) + " " + RTRIM$(TBPPCust(1).State) + "  " + RTRIM$(TBPPCust(1).Zip)
     PRINT #10, Account!;
     PRINT #10, TAB(10); TBPPCust(1).CustName;
     PRINT #10, TAB(40); LEFT$(LTRIM$(Address$), 43); TAB(95); LEFT$(Address2$, 28);
     PRINT #10, TAB(125); USING "######"; TBPPCust(1).FarmEquip
     lc = lc + 1
     CustCount = CustCount + 1
     TotalFarm# = TotalFarm# + TBPPCust(1).FarmEquip
     END IF

2220 IF lc >= 58 THEN PRINT #10, CHR$(12); : GOSUB 2900
2230 NEXT x
     PRINT #10,
     PRINT #10, "Total Customers .. "; USING "######,#"; CustCount
     PRINT #10, "Total Value : Farm Equip .. "; USING "$$######,#"; TotalFarm#
2240 CLOSE
      CLS
      LOCATE 1, 1
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      Header$ = "Tax Account Listing"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$

2250 GOTO 10

2900 p = p + 1
     PRINT #10, TAB(23); "Tax Customer Account Listing : Farm Equipment"
     PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "####"; p
     PRINT #10, "Real Estate Customers"

2910 PRINT #10, "Acct. #"; TAB(10); "Name"; TAB(40); "Address"; TAB(126); "Value"
     PRINT #10, STRING$(132, "="): lc = 5: RETURN

3000 REM
     CustCount = 0: TotalMT# = 0
3010 scren$ = "2.02.03": CLS : GOSUB 63500
    COLOR 11: LOCATE 1, 32: PRINT "Customer Listing : Mach/Tools"
3020 LOCATE 4, 1
3030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount # or (L)ast Name (A,L  or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
3040 IF cf = 1 THEN 10
3050 IF NOT (in$ = "A" OR in$ = "L" OR in$ = "M") THEN 3020
3060 SortSpec$ = in$: LOCATE 4, 79: PRINT SortSpec$
3070 TaxType$ = "P"
     ReportFile$ = "TBBUSLST.PRN"
     RptHandle% = 10
     OPEN ReportFile$ FOR OUTPUT AS #RptHandle%
     GOSUB 3900
3140 GOSUB 60000
3150 IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10
3160 IF SortSpec$ = "A" THEN
       EndofFile = LOF(TBFile) / LEN(TBPPCust(1))
     END IF

     IF SortSpec$ = "L" THEN
        OPEN TBpath$ + "TBPPname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS MREC$: EndofFile = LOF(4) / 24
     END IF
3170 FOR x = 1 TO EndofFile
     IF SortSpec$ = "L" THEN
      GET 4, x
      GET 1, CVS(MREC$), TBPPCust(1)
      Account! = CVS(MREC$)
        ELSE
      GET 1, x, TBPPCust(1)
     END IF
      Account! = x


3191 IF TBPPCust(1).MachTools > 0 THEN
     IF TBPPCust(1).Deleted = "Y" THEN 3230
     Address$ = RTRIM$(TBPPCust(1).Address1) + "  " + RTRIM$(TBPPCust(1).Address2)

     Address2$ = RTRIM$(TBPPCust(1).City) + " " + RTRIM$(TBPPCust(1).State) + "  " + RTRIM$(TBPPCust(1).Zip)
     PRINT #10, Account!;
     PRINT #10, TAB(10); TBPPCust(1).CustName;
     PRINT #10, TAB(40); LEFT$(LTRIM$(Address$), 43); TAB(95); LEFT$(Address2$, 28);
     PRINT #10, TAB(125); USING "######"; TBPPCust(1).MachTools
     lc = lc + 1
     CustCount = CustCount + 1
     TotalMT# = TotalMT# + TBPPCust(1).MachTools
     END IF
     
3220 IF lc >= 58 THEN PRINT #10, CHR$(12); : GOSUB 3900
3230 NEXT x
     PRINT #10,
     PRINT #10, "Total Customers .. "; USING "######,#"; CustCount
     PRINT #10, "Total Value : Mach Tools .. "; USING "$$######,#"; TotalMT#
     
3240 CLOSE
      CLS
      LOCATE 1, 1
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      Header$ = "Tax Account Listing"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$

3250 GOTO 10

3900 p = p + 1
     PRINT #10, TAB(23); "Tax Customer Account Listing : Machine/Tools"
     PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "####"; p
     PRINT #10, "Real Estate Customers"

3910 PRINT #10, "Acct. #"; TAB(10); "Name"; TAB(40); "Address"; TAB(126); "Value"
     PRINT #10, STRING$(130, "="): lc = 5: RETURN
4000 REM
     CustCount = 0: TotalMT# = 0
4010 scren$ = "2.02.04": CLS : GOSUB 63500
    COLOR 11: LOCATE 1, 32: PRINT "Customer Listing : Merchant Cap"
4020 LOCATE 4, 1
4030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount # or (L)ast Name (A,L  or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
4040 IF cf = 1 THEN 10
4050 IF NOT (in$ = "A" OR in$ = "L" OR in$ = "M") THEN 4020
4060 SortSpec$ = in$: LOCATE 4, 79: PRINT SortSpec$
4070 TaxType$ = "P"
     ReportFile$ = "TBMC.PRN"
     RptHandle% = 10
     OPEN ReportFile$ FOR OUTPUT AS #RptHandle%
     GOSUB 4900
4140 GOSUB 60000
4150 IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10
4160 IF SortSpec$ = "A" THEN
       EndofFile = LOF(TBFile) / LEN(TBPPCust(1))
     END IF

     IF SortSpec$ = "L" THEN
        OPEN TBpath$ + "TBPPname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS MREC$: EndofFile = LOF(4) / 24
     END IF
4170 FOR x = 1 TO EndofFile
     IF SortSpec$ = "L" THEN
      GET 4, x
      GET 1, CVS(MREC$), TBPPCust(1)
      Account! = CVS(MREC$)
        ELSE
      GET 1, x, TBPPCust(1)
     END IF
      Account! = x


4191 IF TBPPCust(1).MerchCap > 0 THEN
     IF TBPPCust(1).Deleted = "Y" THEN 4230
     Address$ = RTRIM$(TBPPCust(1).Address1) + "  " + RTRIM$(TBPPCust(1).Address2)

     Address2$ = RTRIM$(TBPPCust(1).City) + " " + RTRIM$(TBPPCust(1).State) + "  " + RTRIM$(TBPPCust(1).Zip)
     PRINT #10, Account!;
     PRINT #10, TAB(10); TBPPCust(1).CustName;
     PRINT #10, TAB(40); LEFT$(LTRIM$(Address$), 43); TAB(95); LEFT$(Address2$, 28);
     PRINT #10, TAB(125); USING "######"; TBPPCust(1).MerchCap
     lc = lc + 1
     CustCount = CustCount + 1
     TotalMT# = TotalMT# + TBPPCust(1).MerchCap
     END IF

4220 IF lc >= 58 THEN PRINT #10, CHR$(12); : GOSUB 4900
4230 NEXT x
     PRINT #10, "Total Customers .. "; USING "######,#"; CustCount
     PRINT #10, "Total Value : Merchant Capital .. "; USING "$$######,#"; TotalMT#

4240 CLOSE
      CLS
      LOCATE 1, 1
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      Header$ = "Tax Account Listing"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$

4250 GOTO 10

4900 p = p + 1
     PRINT #10, TAB(23); "Tax Customer Account Listing : Merchant Capital"
     PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "####"; p
     PRINT #10, "Real Estate Customers"

4910 PRINT #10, "Acct. #"; TAB(10); "Name"; TAB(40); "Address"; TAB(126); "Value"
     PRINT #10, STRING$(130, "="): lc = 5: RETURN

5000 REM
     CustCount = 0: TotalMT# = 0
5010 scren$ = "2.02.05": CLS : GOSUB 63500
    COLOR 11: LOCATE 1, 32: PRINT "Customer Listing : Mobile Homes"
5020 LOCATE 4, 1
5030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount # or (L)ast Name (A,L  or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
5040 IF cf = 1 THEN 10
5050 IF NOT (in$ = "A" OR in$ = "L" OR in$ = "M") THEN 5020
5060 SortSpec$ = in$: LOCATE 4, 79: PRINT SortSpec$
5070 TaxType$ = "P"
     ReportFile$ = "TBMH.PRN"
     RptHandle% = 10
     OPEN ReportFile$ FOR OUTPUT AS #RptHandle%
     GOSUB 5900
5140 GOSUB 60000
5150 IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10
5160 IF SortSpec$ = "A" THEN
       EndofFile = LOF(TBFile) / LEN(TBPPCust(1))
     END IF

     IF SortSpec$ = "L" THEN
        OPEN TBpath$ + "TBPPname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS MREC$: EndofFile = LOF(4) / 24
     END IF
5170 FOR x = 1 TO EndofFile
     IF SortSpec$ = "L" THEN
      GET 4, x
      GET 1, CVS(MREC$), TBPPCust(1)
      Account! = CVS(MREC$)
        ELSE
      GET 1, x, TBPPCust(1)
     END IF
      Account! = x


5191 IF TBPPCust(1).MobileHome > 0 THEN
     IF TBPPCust(1).Deleted = "Y" THEN 5230
     Address$ = RTRIM$(TBPPCust(1).Address1) + "  " + RTRIM$(TBPPCust(1).Address2)

     Address2$ = RTRIM$(TBPPCust(1).City) + " " + RTRIM$(TBPPCust(1).State) + "  " + RTRIM$(TBPPCust(1).Zip)
     PRINT #10, Account!;
     PRINT #10, TAB(10); TBPPCust(1).CustName;
     PRINT #10, TAB(40); LEFT$(LTRIM$(Address$), 43); TAB(95); LEFT$(Address2$, 28);
     PRINT #10, TAB(125); USING "######"; TBPPCust(1).MobileHome
     lc = lc + 1
     CustCount = CustCount + 1
     TotalMT# = TotalMT# + TBPPCust(1).MobileHome
     END IF

5220 IF lc >= 58 THEN PRINT #10, CHR$(12); : GOSUB 5900
5230 NEXT x
     PRINT #10,
     PRINT #10, "Total Customers .. "; USING "######,#"; CustCount
     PRINT #10, "Total Value : Mobile Homes .. "; USING "$$######,#"; TotalMT#

5240 CLOSE
      CLS
      LOCATE 1, 1
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      Header$ = "Tax Account Listing"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$

5250 GOTO 10

5900 p = p + 1
     PRINT #10, TAB(23); "Tax Customer Account Listing : Mobile Homes"
     PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "####"; p
     PRINT #10, "Real Estate Customers"

5910 PRINT #10, "Acct. #"; TAB(10); "Name"; TAB(40); "Address"; TAB(126); "Value"
     PRINT #10, STRING$(130, "="): lc = 5: RETURN
6000 scren$ = "2.01.06": CLS : GOSUB 63500: COLOR 11
      LOCATE 1, 32: PRINT "Customer New Bldg Listing"
      TotCust! = 0
      Tot1# = 0

6020 LOCATE 4, 1: PRINT STRING$(79, 32): LOCATE 4, 1
6030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount or (L)ast Name (A,L or ESC-ABORT) ..... "; : fl = 1: GOSUB 62210
6040 IF cf = 1 THEN 10
6050 IF NOT (in$ = "A" OR in$ = "L") THEN 2020
6060 SortSpec$ = in$: LOCATE 4, 79: PRINT SortSpec$
     TaxType$ = "R"
     GOSUB 60000
     
6160 IF SortSpec$ = "A" THEN
      EndofFile = LOF(TBFile) / LEN(TBRECust(1))
     END IF

     IF SortSpec$ = "L" THEN
      OPEN TBpath$ + "TBREname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
      FIELD 4, 20 AS lastnam$, 4 AS MREC$: EndofFile = LOF(4) / 24
     END IF

     LOCATE 11, 1: COLOR 11: PRINT STRING$(79, "="): LOCATE 13, 23: PRINT "Now Processing : Please be Patient!": COLOR 14
6120 ReportFile$ = "TBNWBLDG.PRN"
     RptHandle% = 10
     OPEN ReportFile$ FOR OUTPUT AS #RptHandle%
     GOSUB 6900
6170 FOR Cnt! = 1 TO EndofFile
     IF SortSpec$ = "L" THEN
      GET 4, Cnt!
      GET TBFile, CVS(MREC$), TBRECust(1)
      Account! = CVS(MREC$)
      GOTO 6190
     END IF

      
      GET TBFile, Cnt!, TBRECust(1)
      
      Account! = Cnt!
6190 IF TBRECust(1).Deleted = "Y" THEN 6230
     
     IF TBRECust(1).NewBldgValue > 0 THEN
      PRINT #10, Account!; TAB(10); TBRECust(1).CustName;
       PRINT #10, TAB(50); USING "#######,#"; TBRECust(1).NewBldgValue
      PRINT #10, TBRECust(1).NewCom1;
       PRINT #10, TAB(30); TBRECust(1).NewCom2
      lc = lc + 2
      PRINT #10, STRING$(79, "-")
      lc = lc + 1
      Tot1# = Tot1# + TBRECust(1).NewBldgValue
      TotCust! = TotCust! + 1

     END IF
     
6220 IF lc >= 55 THEN PRINT #10, CHR$(12); : GOSUB 6900
6230 NEXT Cnt!
        PRINT #10, "Totals   Customers:"; USING "#####,#"; TotCust!;
         PRINT #10, TAB(50); USING "#######,#"; Tot1#
6235 PRINT CHR$(12);
6240 CLOSE
      CLS
      LOCATE 1, 1
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      Header$ = "New Bldg Tax Report"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$

6250 GOTO 10

6900 p = p + 1
     PRINT #10, TAB(27); "New Bldg Customer Listing"
     PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "####"; p
     PRINT #10, "":
     PRINT #10, "Real Estate Customers w/New Bldg Values"
     
6910 PRINT #10, "Acct #"; TAB(10); "Account Name"; TAB(50); "New Bldg Value"
     PRINT #10, "New Comment 1"; TAB(30); "New Comment 2"
6940 PRINT #10, STRING$(79, "="): lc = 4: RETURN

7000 scren$ = "2.01.01": CLS : GOSUB 63500
     TotCust! = 0
     TotTax# = 0
     Tot1# = 0

     COLOR 11: LOCATE 1, 27: PRINT "New Bldg Pre-Bill Proof"
     LOCATE 3, 1
     PRINT "(N)umerical Order or (A)lphabetical Order (N,A or <ESC>-Cancel) .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN RUN "tbccustr"
     IF NOT (in$ = "N" OR in$ = "A") THEN 7000
     SortType$ = in$

7101 LOCATE 5, 1
     COLOR 11
     PRINT "Enter the Tax Year (YYYY) (ESC-Cancel) .... "; : fl = -4: GOSUB 62000
     IF cf = 1 THEN 10
     IF LEN(in$) <> 4 THEN 7101
     TaxYear$ = in$
     COLOR 11
7102 PRINT "Enter the Tax Rate Per $100 (ESC-Cancel) .. "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
     TaxRate! = VAL(in$)
     COLOR 11
7103 PRINT "Enter the Beginning Bill Number ........... "; : fl = -7: GOSUB 62000
     BillNumber! = VAL(in$)
7104 COLOR 11
     LOCATE 10, 1
     PRINT "Is the Above Information Correct (Y/N) (ESC-Cancel) .. "; : fl = 1
     GOSUB 62210
     IF cf = 1 THEN 10
     IF NOT (in$ = "Y" OR in$ = "N") THEN 7104
     IF in$ = "N" THEN 7000
     TaxType$ = "R"
     GOSUB 60000
     LOCATE 12, 1: COLOR 11: PRINT "NOW SPOOLING OUT REPORT : PLEASE WAIT!!!"
7150 IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10
     ReportFile$ = "TBPREBLD.REP"
     OPEN "O", 10, "TBPREBLD.REP": WIDTH #10, 255
     GOSUB 7900
     IF SortType$ = "A" THEN
        OPEN TBpath$ + "TBREname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS lastnam$, 4 AS MREC$: EndofFile = LOF(4) / 24
      ELSE
       OPEN TBpath$ + "TBMAP.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
       FIELD 4, 20 AS lastnam$, 4 AS MREC$
       EndofFile = LOF(4) / 24
    END IF


7170 FOR x = 1 TO EndofFile
     GET 4, x
     Account! = CVS(MREC$)
     GET TBFile, CVS(MREC$), TBRECust(1)


     IF TBRECust(1).NewBldgValue > 0 THEN

     'Figure Tax
     TaxAmt# = (TBRECust(1).NewBldgValue / 100) * TaxRate!
     TaxAmt# = INT((TaxAmt# * 100) + .5) / 100

     PRINT #10, Account!; TAB(10); TBRECust(1).CustName;
       PRINT #10, TAB(50); USING "#######,#"; TBRECust(1).NewBldgValue;
        PRINT #10, TAB(62); USING "######.##"; TaxAmt#;
         PRINT #10, TAB(72); USING "#######"; BillNumber!
        
     TBRECust(1).BBCurYrBillNumber = LTRIM$(STR$(BillNumber!))
     TBRECust(1).BBCurYrTaxAmount = TaxAmt#
     PUT TBFile, CVS(MREC$), TBRECust(1)

     BillNumber! = BillNumber! + 1
     lc = lc + 1
     Tot1# = Tot1# + TBRECust(1).NewBldgValue
     TotTax# = TotTax# + TaxAmt#
     TotCust! = TotCust! + 1

     END IF


7220 IF lc >= 58 THEN PRINT #10, CHR$(12); : GOSUB 1900
7230 NEXT x
     PRINT #10, STRING$(79, "=")
     PRINT #10, " Total Bills to Generate: "; USING "#####,#"; TotCust!
     PRINT #10, "Total Tax Amount to Bill: "; USING "#####,#.##"; TotTax#
     PRINT #10, " Total Value of New Bldg: "; USING "#######,#"; Tot1#

     PRINT #10, CHR$(12);
7240 CLOSE
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      
      Header$ = "New Building Pre-Bill Report"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
7250 RUN "tbccustr"



7900 p = p + 1
     PRINT #10, TAB(27); "New Bldg Customer Pre-Billing Report"
     PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "####"; p
     PRINT #10, "":
     PRINT #10, "Real Estate Customers w/New Bldg Values"
7910 PRINT #10, "Acct #"; TAB(10); "Account Name"; TAB(45); "New Bldg Value";
       PRINT #10, TAB(63); "Tax Amt"; TAB(72); "Bill #"
       PRINT #10, STRING$(79, "="): lc = 4: RETURN

8000 RUN "tbmiles"
9000 scren$ = "2.09.00": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 28: PRINT "Real Estate Edited Accts"
     LOCATE 3, 1
     TaxType$ = "R"
9011 LOCATE 4, 1: COLOR 11
     PRINT "The next question will prompt you for the date.  This date"
     PRINT "should be the earliest date to begin including edited real"
     PRINT "estate accounts.  i.e. All accounts from this date will be"
     PRINT "included on the report!!!"
9012 LOCATE 9, 1: COLOR 11
     PRINT "Enter Date (mmddyyyy)  to Pull <ESC>-Cancel .. "; : fl = -8
     GOSUB 62000
     IF cf = 1 THEN 10
     IF LEN(in$) <> 8 THEN 9012
     DatetoPull$ = LEFT$(in$, 2) + "/" + MID$(in$, 3, 2) + "/" + MID$(in$, 5, 4)
     DatetoPull = Date2Num(DatetoPull$)
     LOCATE 15, 1: COLOR 11: PRINT "Now Printing the Report : Please Wait!"
9100 GOSUB 60000
9150 IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10
     OPEN "O", 10, "TBREEDIT.REP": WIDTH #10, 255
     GOSUB 9900
     OPEN TBpath$ + "TBREname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
     FIELD 4, 20 AS lastnam$, 4 AS MREC$: EndofFile = LOF(4) / 24
     
9170 FOR x = 1 TO EndofFile
     GET 4, x
     GET 1, CVS(MREC$), TBRECust(1)
     Account! = CVS(MREC$)

9191 IF TBRECust(1).Deleted = "Y" THEN 9230

   ' Do the tests here

     IF LEN(QPTrim$(TBRECust(1).MapNumber)) = 0 THEN 9230 ' do not allow no map numbers to pass thru
     IF TBRECust(1).LastDateChange <= 0 THEN 9230        ' do not allow no map numbers to pass thru
     IF TBRECust(1).LastDateChange < DatetoPull THEN 9230
   

   PRINT #10, USING "######"; Account!;
     PRINT #10, TAB(10); TBRECust(1).CustName; TAB(45); TBRECust(1).MapNumber;
     PRINT #10, TAB(70); num2date(TBRECust(1).LastDateChange)
    lc = lc + 1
    CustCount = CustCount + 1
    
      IF lc >= 58 THEN PRINT #10, CHR$(12); : GOSUB 1900
9230 NEXT x
     PRINT #10, "Total Customers Edited .. "; USING "######,#"; CustCount
     PRINT CHR$(12);
9240 CLOSE
      'Set Report Defaults
      LPTPort% = 1
      EntryPoint% = 1
      ReportFile$ = "TBREEDIT.REP"
      Header$ = "Real Estate Edited Accounts"
      PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint%
      KILL ReportFile$
      GOTO 10
     
9900 p = p + 1
     PRINT #10, TAB(27); "Real Estate Edited Accounts"
     PRINT #10,
     PRINT #10, "Accounts Edited Since: "; num2date$(DatetoPull)
     PRINT #10, "Date: "; DATE$; TAB(68); "Page #"; USING "#####"; p
     PRINT #10, "Acct. #"; TAB(10); "Name & Address"; TAB(45); "Map Reference"; TAB(69); "Last Edited"
     PRINT #10, STRING$(80, "=")
     lc = 5
     RETURN

10000 RUN "tbrep25"
11000 RUN "tbrep26"
12000 RUN "tbrep27"
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
  FIELD TBIDXFile, 20 AS lastnam$, 4 AS MREC$
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
GET TBFile, CVS(MREC$), TBRECust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBRECust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(MREC$)
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
  ggo = 1: Record! = CVS(MREC$)
  FOR LLL = 8 TO 13:  LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
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
  FOR ll = 1 TO LOF(TBIDXFile) / 24
  GET TBIDXFile, ll
  IF LEFT$(lastnam$, nn) = nn$ THEN GOTO PPn1
PPn3:
  NEXT ll
  LOCATE 8, 1: PRINT STRING$(79, 32)
  ggo = 0
  RETURN

PPn1:
GET TBFile, CVS(MREC$), TBPPCust(1)
LOCATE 11, 1: COLOR 11
nme$ = TBPPCust(1).CustName
COLOR 10
FOR LLL = 10 TO 12: LOCATE LLL, 1: PRINT STRING$(79, 32): NEXT LLL
LOCATE 11, 1: COLOR 11
PRINT "Full Account Name: "; : COLOR 15: PRINT nme$
COLOR 11
PRINT "   Account Number: "; : COLOR 15: PRINT CVS(MREC$)
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
  ggo = 1: Record! = CVS(MREC$)
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

