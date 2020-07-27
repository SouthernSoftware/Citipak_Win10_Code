10 REM vendor.b
12 CLS
   CLEAR
   ON ERROR GOTO fixerror
   DIM SHARED perc(15), ta#(15), ta$(15)
60 COLOR 10: PRINT "Sys. Date: "; DATE$; : COLOR 11:
   PRINT TAB(31); "Vendor Maintenance"; : COLOR 10: PRINT TAB(70); "1.00.00": PRINT STRING$(79, "=")
70 PRINT : PRINT : PRINT : COLOR 11
80 PRINT TAB(23); "1- Add New Vendor "
85 PRINT
90 PRINT TAB(23); "2- Change/Delete Vendor "
95 PRINT
100 PRINT TAB(23); "3- Print Numerical Listing"
105 PRINT
110 PRINT TAB(23); "4- Print Alphabetical Listing"
    PRINT
    PRINT TAB(23); "5- Recalculate YTD Payments "
    PRINT
    PRINT TAB(23); "6- Setup Default Invoice Distribution"
125 LOCATE 19, 26: COLOR 11: PRINT "Selection # (ESC - MENU) .. "; : fl = -1: GOSUB 62000
130 IF cf = 1 THEN CLOSE : CLS : RUN "apmenu"
140 IF VAL(in$) < 1 OR VAL(in$) > 6 THEN 125
150 ON VAL(in$) GOTO 1000, 2000, 3000, 4000, 5000, 6000
1000 CLS : COLOR 10: PRINT "Sys. Date: "; DATE$; : COLOR 11: PRINT TAB(33); "Add New Vendor"; : COLOR 11: PRINT TAB(70); "1.01.00": PRINT STRING$(79, "=")
1010 CLOSE 1: GOSUB 10000: record = LOF(1) / 384 + 1
1090 LOCATE 4, 1: COLOR 11: PRINT "    Assigned Vendor Number .. "; : COLOR 15: PRINT USING "#####"; record

1100 COLOR 11: LOCATE 6, 1: PRINT " 1. Vendor Name ............... "; : fl = 30
     GOSUB 62000
     IF cf = 1 THEN
       IF ADDFLAG = 1 THEN
        CLOSE
        GOTO 10
         ELSE
        CLOSE : GOTO 10
       END IF
       ELSE
       BA$ = in$
     END IF
1110 COLOR 11: PRINT " 2. Address Line 1 ............ "; : fl = 30: GOSUB 62000: IF cf = 1 THEN 10
     BB$ = in$
     COLOR 11: PRINT " 3. Address Line 2 ............ "; : fl = 30: GOSUB 62000: IF cf = 1 THEN 10
     ADD2$ = in$
1120 COLOR 11: PRINT " 4. City ...................... "; : fl = 14: GOSUB 62000: IF cf = 1 THEN 10
     BC$ = in$
1121 COLOR 11: LOCATE 10, 1: PRINT " 5. State ..................... "; : fl = 2: GOSUB 62000: IF cf = 1 THEN 10
     BD$ = in$
1122 COLOR 11: LOCATE 11, 1: PRINT " 6. Zip Code .................. "; : fl = 10: GOSUB 62000: IF cf = 1 THEN 10
     B2$ = in$

1131 COLOR 11: LOCATE 18, 1: PRINT "Selection # to Change (ESC-OK) .. "; : fl = -1: GOSUB 62000: IF cf = 1 THEN 1140

1133 IF VAL(in$) < 1 OR VAL(in$) > 6 THEN 1131

1137 ON VAL(in$) GOSUB 1350, 1360, 1370, 1380, 1390, 1395
1138 GOTO 1131
1140 LOCATE 18, 1: PRINT STRING$(79, 32)
     COLOR 11: LOCATE 12, 1: PRINT "Do You Pay to the Same Name & Address  (Y/N) ... "; : fl = 1: GOSUB 62210: IF NOT (in$ = "Y" OR in$ = "N") THEN 1140 ELSE Q$ = in$
1145 IF Q$ = "Y" THEN AZ = 1: GOTO 1210
1150 AZ = 0: COLOR 11: LOCATE 12, 1: PRINT STRING$(79, 32): LOCATE 12, 1
1160 COLOR 11: PRINT " 7. Pay to Name ............... "; : fl = 30: GOSUB 62000: CA$ = in$
1170 COLOR 11: PRINT " 8. Address  .................. "; : fl = 30: GOSUB 62000: CU$ = in$
     COLOR 11: PRINT " 9. Address Line 2 ............ "; : fl = 30: GOSUB 62000: PADD2$ = in$
1180 COLOR 11: PRINT "10. City ...................... "; : fl = 14: GOSUB 62000: CC$ = in$
1181 COLOR 11: PRINT "11. State ..................... "; : fl = 2: GOSUB 62000: CD$ = in$
1182 COLOR 11: PRINT "12. Zip Code .................. "; : fl = 10: GOSUB 62000: C1$ = in$
1190 COLOR 11: LOCATE 18, 1: PRINT "Selection # to Change  (ESC-OK) .. "; : fl = -2
      GOSUB 62000: IF cf = 1 THEN 1250

1200 IF VAL(in$) < 6 OR VAL(in$) > 12 THEN 1190
1205 ON VAL(in$) GOSUB 1350, 1360, 1370, 1380, 1390, 1395, 1400, 1410, 1420, 1430, 1440, 1445
     GOTO 1190
1210 LOCATE 18, 1: PRINT STRING$(79, 32)
     LOCATE 18, 1
1250 IF AZ = 1 THEN CA$ = BA$: CU$ = BB$: CC$ = BC$: CD$ = BD$: C1$ = B2$
     PADD2$ = ADD2$
1260 LSET VENNUMBER$ = MKI$(record)
     LSET venname$ = BA$
     LSET addRESS$ = BB$
     LSET address2$ = ADD2$
1270 LSET CitY$ = BC$: LSET St$ = BD$: LSET ZiP$ = B2$
1280 LSET payto$ = CA$: LSET payadd$ = CU$: LSET paddress2$ = PADD2$: LSET paycity$ = CC$
1290 LSET payst$ = CD$: LSET payzip$ = C1$
1300 LSET curbal$ = MKD$(0): LSET first$ = MKS$(0): LSET last$ = MKS$(0)
1320 PUT 1, record
     FOR LL = 12 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
     LOCATE 12, 1: COLOR 11
     PRINT " 7. Telephone Number .......... "; : fl = 12: GOSUB 62000: LSET TEL$ = in$
     COLOR 11
     PRINT " 8. YTD Payments .............. "; : fl = -10: GOSUB 62000: LSET YTDPAY$ = MKS$(VAL(in$))
     COLOR 11
     PRINT " 9. Federal ID Number ......... "; : fl = 10: GOSUB 62000: LSET FEDID$ = in$
     COLOR 11
     PRINT "10. State Code ................ "; : fl = 3: GOSUB 62000: LSET stcode$ = in$
     COLOR 11
     PRINT "11. County Code ............... "; : fl = 3: GOSUB 62000: LSET ctycode$ = in$
     COLOR 11
     PRINT "12. Get 1099 (Y/N) ............ "; : fl = 1: GOSUB 62210: LSET ten99$ = in$
1325 COLOR 11: LOCATE 18, 1: PRINT "Selection # to Change  (ESC-OK) .. "; : fl = -2
      GOSUB 62000: IF cf = 1 THEN 1330

 IF VAL(in$) < 6 OR VAL(in$) > 12 THEN 1190
 ON VAL(in$) GOSUB 1500, 1500, 1500, 1500, 1500, 1500, 1500, 1510, 1520, 1530, 1540, 1550
 GOTO 1325
    
1330 ADDFLAG = 1
     LSET YTDPUR$ = MKS$(0)
     LSET POFIRST$ = MKS$(0)
     LSET POLAST$ = MKS$(0)
     PUT 7, record
     CLOSE
1340 GOTO 1000

1350 LOCATE 6, 33: PRINT ""; : fl = 30: GOSUB 62000: BA$ = in$: RETURN
1360 LOCATE 7, 33: PRINT ""; : fl = 30: GOSUB 62000: BB$ = in$: RETURN
1370 LOCATE 8, 33: PRINT ""; : fl = 30: GOSUB 62000: ADD2$ = in$: RETURN
1380 LOCATE 9, 33: PRINT ""; : fl = 22: GOSUB 62000: BC$ = in$: RETURN
1390 LOCATE 10, 33: PRINT ""; : fl = 2: GOSUB 62000: BD$ = in$:  RETURN
1395 LOCATE 11, 33: PRINT ""; : fl = 10: GOSUB 62000: B2$ = in$: RETURN
1400 LOCATE 12, 33: PRINT ""; : fl = 30: GOSUB 62000: CA$ = in$: RETURN
1410 LOCATE 13, 33: PRINT ""; : fl = 30: GOSUB 62000: CU$ = in$: RETURN
1420 LOCATE 14, 33: PRINT ""; : fl = 30: GOSUB 62000: PADD2$ = in$: RETURN
1430 LOCATE 15, 33: PRINT ""; : fl = 22: GOSUB 62000: CC$ = in$: RETURN
1440 LOCATE 16, 33: PRINT ""; : fl = 2: GOSUB 62000: CD$ = in$: RETURN
1445 LOCATE 17, 33: PRINT ""; : fl = 10: GOSUB 62000: C1$ = in$: RETURN


1500 LOCATE 12, 33: PRINT ""; : fl = 12: GOSUB 62000: LSET TEL$ = in$: RETURN
1510 LOCATE 13, 33: PRINT ""; : fl = 10: GOSUB 62000: LSET YTDPAY$ = MKS$(VAL(in$)): RETURN
1520 LOCATE 14, 33: PRINT ""; : fl = 10: GOSUB 62000: LSET FEDID$ = in$: RETURN
1530 LOCATE 15, 33: PRINT ""; : fl = 3: GOSUB 62000: LSET stcode$ = in$: RETURN
1540 LOCATE 16, 33: PRINT ""; : fl = 3: GOSUB 62000: LSET ctycode$ = in$: RETURN
1550 LOCATE 17, 33: PRINT ""; : fl = 1: GOSUB 62210: LSET ten99$ = in$: RETURN


2000 CLS : COLOR 10: PRINT "Sys. Date: "; DATE$; : COLOR 11: PRINT TAB(30); "Change/Delete Vendor"; : COLOR 10: PRINT TAB(70); "1.02.00": PRINT STRING$(79, "=")
2010 COLOR 11: LOCATE 3, 1: PRINT "Vendor Account Number (ESC-Menu:TAB-Name Search) .. ";
     fl = -5: GOSUB 62000
     IF cf = 1 THEN CLOSE : GOTO 10
     IF cf = 2 THEN
      LOCATE 4, 1: COLOR 11
      PRINT "Enter Name (or part) .. "; : fl = 30: GOSUB 62000
      searchnam$ = in$
      GOSUB 10000
      FOR x = 1 TO LOF(1) / 384: GET 1, x
      IF LEFT$(venname$, 5) = "delet" THEN 2019
          IF LEFT$(venname$, LEN(searchnam$)) = searchnam$ THEN
           COLOR 10
           LOCATE 6, 1
           PRINT "Account of "; venname$
           LOCATE 7, 1: COLOR 11
           PRINT "Correct Account (Y/N/Q) .. "; : fl = 1
           GOSUB 62210
           IF in$ = "Y" THEN f1 = x: account = x: GOTO 2030 ELSE IF in$ = "Q" THEN CLOSE : GOTO 2000
          END IF
2019     NEXT x: CLOSE : GOTO 2000
          
     END IF
     account = VAL(in$)
     GOSUB 10000
     f1 = account: record = account
2030 IF f1 < 1 OR f1 > LOF(1) / 384 THEN 2070
2040 x = f1: GET 1, f1: GET 7, f1
     IF LEFT$(venname$, 6) = "delete" THEN 2070 ELSE GOTO 2100
2070 CLOSE
2080 COLOR 12
     PRINT "Account Not on File, Press any key to Continue .. ";
     fl = 1: GOSUB 62000
     CLOSE : GOTO 2000

2100 LOCATE 3, 1: FOR xx = 3 TO 15: LOCATE xx, 1: PRINT STRING$(79, 32): NEXT xx: LOCATE 3, 1
2110 COLOR 11: PRINT "Vendor Account Number: "; : COLOR 15
     PRINT USING "#####"; account
     COLOR 11
2120 PRINT " 1- Vendor Name    ..... "; : COLOR 15: PRINT venname$
     COLOR 11
2130 PRINT " 2- Address Line 1 ..... "; : COLOR 15: PRINT addRESS$
     COLOR 11
     PRINT " 3- Address Line 2 ..... "; : COLOR 15: PRINT address2$
     COLOR 11
2140 PRINT " 4- City ............... "; : COLOR 15: PRINT CitY$
     COLOR 11
2141 PRINT " 5- State .............. "; : COLOR 15: PRINT St$
     COLOR 11
2142 PRINT " 6- Zip ................ "; : COLOR 15: PRINT ZiP$
     COLOR 11
2150 PRINT STRING$(79, "-")
2155 PRINT " 7- Pay to   ........... "; : COLOR 15: PRINT payto$
     COLOR 11
2160 PRINT " 8- Address  ........... "; : COLOR 15: PRINT payadd$
     COLOR 11
     PRINT " 9- Address Line 2 ..... "; : COLOR 15: PRINT paddress2$
     COLOR 11
2170 PRINT "10- City ............... "; : COLOR 15: PRINT paycity$
     COLOR 11
2171 PRINT "11- State .............. "; : COLOR 15: PRINT payst$
     COLOR 11
2172 PRINT "12- Zip ................ "; : COLOR 15: PRINT payzip$
     COLOR 11
     PRINT "13- Telephone Number ... "; : COLOR 15: PRINT TEL$
     COLOR 11
     PRINT "14- YTD Payments ....... "; : COLOR 15: PRINT USING "$#####,#.##"; CVS(YTDPAY$)
     COLOR 11
     PRINT "15- Federal ID Number .. "; : COLOR 15: PRINT FEDID$
     COLOR 11
     PRINT "16- State Code ......... "; : COLOR 15: PRINT stcode$
     COLOR 11
     PRINT "17- County Code ........ "; : COLOR 15: PRINT ctycode$
     COLOR 11
     PRINT "18- Get 1099's (Y/N) ... "; : COLOR 15: PRINT ten99$
     PRINT
2190 COLOR 11: LOCATE 23, 1
     PRINT "Selection # (ESC-OK)([TAB] to DELETE) .. "; : fl = -2: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 2000
2192 IF cf = 2 THEN
      LOCATE 23, 1: PRINT STRING$(50, 32): PRINT CHR$(7); : COLOR 11: LOCATE 23, 1
       PRINT "OK to Delete (Y/N) .. "; : fl = 1: GOSUB 62210
      IF in$ = "Y" THEN
       LSET venname$ = "delete": PUT 1, f1: CLOSE 1: GOTO 2000
      ELSE
      GOTO 2190
      END IF
     END IF
     f5 = VAL(in$)
2200 IF f5 < 1 OR f5 > 18 THEN 2190
2210 ON VAL(in$) GOSUB 2320, 2330, 2335, 2340, 2350, 2360, 2370, 2380, 2385, 2390, 2400, 2410, 2420, 2430, 2440, 2450, 2460, 2470
2220 GOTO 2190
2320 LOCATE 4, 26: PRINT ""; : fl = 30: GOSUB 62000: J$ = in$: LSET venname$ = J$: PUT 1, x: sortflag = 1: RETURN
2330 LOCATE 5, 26: PRINT ""; : fl = 30: GOSUB 62000: J$ = in$: LSET addRESS$ = J$: PUT 1, x: RETURN
2335 LOCATE 6, 26: PRINT ""; : fl = 30: GOSUB 62000: J$ = in$: LSET address2$ = J$: PUT 1, x: RETURN
2340 LOCATE 7, 26: PRINT ""; : fl = 22: GOSUB 62000: J$ = in$: LSET CitY$ = J$: PUT 1, x: RETURN
2350 LOCATE 8, 26: PRINT ""; : fl = 2: GOSUB 62000: J$ = in$: LSET St$ = J$: PUT 1, x: RETURN
2360 LOCATE 9, 26: PRINT ""; : fl = 10: GOSUB 62000: J$ = in$: LSET ZiP$ = J$: PUT 1, x: RETURN
2370 LOCATE 11, 26: PRINT ""; : fl = 30: GOSUB 62000: J$ = in$: LSET payto$ = J$: PUT 1, x: RETURN
2380 LOCATE 12, 26: PRINT ""; : fl = 30: GOSUB 62000: J$ = in$: LSET payadd$ = J$: PUT 1, x: RETURN
2385 LOCATE 13, 26: PRINT ""; : fl = 30: GOSUB 62000: J$ = in$: LSET paddress2$ = J$: PUT 1, x: RETURN
2390 LOCATE 14, 26: PRINT ""; : fl = 22: GOSUB 62000: J$ = in$: LSET paycity$ = J$: PUT 1, x: RETURN
2400 LOCATE 15, 26: PRINT ""; : fl = 2: GOSUB 62000: J$ = in$: LSET payst$ = J$: PUT 1, x: RETURN
2410 LOCATE 16, 26: PRINT ""; : fl = 10: GOSUB 62000: J$ = in$: LSET payzip$ = J$: PUT 1, x: RETURN
2420 LOCATE 17, 26: PRINT ""; : fl = 12
     GOSUB 62000: J$ = in$: LSET TEL$ = J$: PUT 7, x: RETURN
2430 LOCATE 18, 26: PRINT STRING$(15, 32): LOCATE 18, 26
     PRINT ""; : fl = 9
     GOSUB 62000: LSET YTDPAY$ = MKS$(VAL(in$)): PUT 7, x: RETURN
2440 LOCATE 19, 26: PRINT STRING$(15, 32): LOCATE 19, 26
     PRINT ""; : fl = 12
     GOSUB 62000: LSET FEDID$ = in$: PUT 7, x: RETURN
2450 LOCATE 20, 26: PRINT STRING$(15, 32): LOCATE 20, 26
     PRINT ""; : fl = 3
     GOSUB 62000: LSET stcode$ = in$: PUT 7, x: RETURN
2460 LOCATE 21, 26: PRINT STRING$(15, 32): LOCATE 21, 26
     PRINT ""; : fl = 3
     GOSUB 62000: LSET ctycode$ = in$: PUT 7, x: RETURN
2470 LOCATE 22, 26: PRINT STRING$(15, 32): LOCATE 22, 26
     PRINT ""; : fl = 1
     GOSUB 62000: LSET ten99$ = in$: PUT 7, x: RETURN

3000 REM vendor numeric
3001 CLS
3010 COLOR 11: PRINT "Output to (S)creen or (P)rinter (S,P or ESC-Abort) .. "; : fl = 1: GOSUB 62210: IF cf = 1 THEN 10
 IF NOT (in$ = "S" OR in$ = "P") THEN 3000
 devspec$ = in$: LOCATE 4, 55: PRINT devspec$
 IF devspec$ = "S" THEN DEV$ = "scrn:": LOCATE 4, 1: PRINT STRING$(79, 32): LOCATE 3, 1:  ELSE DEV$ = "lpt1:"
 OPEN "O", 10, DEV$: WIDTH #10, 255
3060 GG$ = "Numerical":  GOSUB 3900
     IF DEV$ = "scrn:" THEN 3085
3070 COLOR 14: PRINT " Now Reading Accounts"
3085 GOSUB 10000
3090 FOR x = 1 TO LOF(1) / 384: GET 1, x: GET 7, x
     IF LEFT$(venname$, 6) = "delete" THEN 3130
3100 IF DEV$ = "scrn:" THEN 3105 ELSE 3110
3105 IF LC >= 18 THEN LOCATE 21, 1: PRINT "Press any key for next page! (ESC-Quit)"; : fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10 ELSE CLS : GOSUB 3900: GOTO 3120
3110 IF LC > 55 THEN PRINT #10, CHR$(12); : GOSUB 3900
3120 PRINT #10, USING "#####"; x;
     PRINT #10, TAB(20); venname$;
     PRINT #10, TAB(60); ctycode$; TAB(70); stcode$;
     IF CVS(FIRSTAPDIST$) = 0 THEN PRINT #10, TAB(77); "NO" ELSE PRINT #10, TAB(77); "YES"
      LC = LC + 1

3130 NEXT x: IF devspec$ = "S" THEN CLOSE : LOCATE 21, 21: PRINT "End of Listing - Press Any Key for Menu "; : fl = 1: GOSUB 62000: GOTO 10
     PRINT #10, CHR$(12);
3140 CLOSE 10: CLOSE 1
3160 GOTO 10
3900 CLS : PRINT #10, TAB(10); "ACCOUNTS PAYABLE VENDOR LISTING:   "; GG$; " Order"
     page = page + 1
3910 PRINT #10, "Date - "; DATE$; TAB(65); "Page - "; page: PRINT #10, " "
3920 PRINT #10, "Vendor Number"; TAB(20); "Vendor Name"; TAB(57); "Cty Code"; TAB(67); "St Code"; TAB(76); "Dist"
     PRINT #10, STRING$(79, "=")
3930 LC = 6: RETURN
4000 REM - alphabetical listing
     CLS
4001 COLOR 11: PRINT "Output to (S)creen or (P)rinter (S,P or ESC-Abort) .. "; : fl = 1: GOSUB 62210: IF cf = 1 THEN 10
     IF NOT (in$ = "S" OR in$ = "P") THEN 4000
     devspec$ = in$: LOCATE 4, 55: PRINT devspec$
     IF devspec$ = "S" THEN DEV$ = "scrn:": LOCATE 4, 1: PRINT STRING$(79, 32): LOCATE 3, 1:  ELSE DEV$ = "lpt1:"
4005 GOSUB 4500
     CLOSE
     OPEN "O", 10, DEV$: WIDTH #10, 255
4008 GG$ = "Alphabetical": GOSUB 4090
4012 GOSUB 10000
4014 OPEN "I", 2, "vendoral.ind"
4017 INPUT #2, A$
4020 IF RIGHT$(A$, 5) = "99999" THEN 4033
4023 GET 1, VAL(RIGHT$(A$, 5)): GET 7, VAL(RIGHT$(A$, 5))
4024 IF DEV$ = "scrn:" THEN 4025 ELSE 4026
4025 IF LC >= 18 THEN LOCATE 21, 1: PRINT "Press any key for next page! (ESC-Quit)"; : fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10 ELSE CLS : GOSUB 4090: GOTO 4029
4026 IF LC > 55 THEN PRINT #10, CHR$(12); : GOSUB 4090
4029 PRINT #10, USING "#####"; VAL(RIGHT$(A$, 5));
    PRINT #10, TAB(12); venname$;
    PRINT #10, TAB(50); TEL$: LC = LC + 1
4032 GOTO 4017
4033 IF devspec$ = "S" THEN CLOSE : LOCATE 21, 21: PRINT "End of Listing - Press Any Key for Menu "; : fl = 1: GOSUB 62000: GOTO 10
4035 PRINT #10, CHR$(12);
4036 CLOSE 10: CLOSE 1
4041 GOTO 10
4090 CLS : PRINT #10, TAB(15); "ACCOUNTS PAYABLE VENDOR LISTING:   "; GG$; " Order"
4100 page = page + 1
     PRINT #10, ""
     PRINT #10, "Date - "; DATE$; TAB(65); "Page - "; page
     PRINT #10, " "
4110 PRINT #10, "Vendor Number & Name"; TAB(50); "Telephone #": PRINT #10, STRING$(80, "=")
4120 LC = 6: RETURN
4500 CLS : COLOR 14: PRINT "Now Sorting : please wait!!!"
4510 GOSUB 10000: OPEN "O", 2, "vendor.alp": WIDTH #2, 255
4520 FOR x = 1 TO (LOF(1) / 384)
4530 GET 1, x
4540 IF LEFT$(venname$, 6) = "delete" THEN 4560
4542 zz = INSTR(venname$, ","): IF zz = 0 THEN zz = LEN(venname$) + 1
4550 PRINT #2, LEFT$(venname$, zz - 1); " : "; USING "#####"; x
4560 NEXT x
4570 PRINT #2, "ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ : 99999"
4580 CLOSE 2
4590 SHELL ("sort  <vendor.alp > vendoral.ind")
4600 RETURN

5000 CLS : COLOR 10: PRINT "Sys. Date: "; DATE$;
     COLOR 11: PRINT TAB(30); "Calculate YTD Payments";
     COLOR 10: PRINT TAB(70); "1.05.00": PRINT STRING$(79, "=")
5010 COLOR 11: LOCATE 3, 1: PRINT "Enter Year (XXXX) or (ESC-Menu) .. "; : fl = -4: GOSUB 62000
     IF cf = 1 THEN 10
     paydate$ = in$
     GOSUB 10000: GOSUB 11000
     FOR x = 1 TO LOF(1) / 384: GET 1, x: GET 7, x
     IF LEFT$(venname$, 6) = "delete" THEN 5100
     GOSUB processclient
5100 NEXT x
     CLOSE
     GOTO 10

processclient:
payments# = 0
LOCATE 15, 1: PRINT STRING$(79, 32): LOCATE 15, 1
COLOR 11
PRINT "Vendor Name ... "; : COLOR 15: PRINT venname$
COLOR 11
LSET YTDPAY$ = MKS$(0): PUT 7, x
tr = CVS(first$)

nexone:
IF tr = 0 THEN
 
 LSET YTDPAY$ = MKS$(payments#): PUT 7, x: RETURN
END IF
GET 2, tr
IF LEFT$(invnum$, 1) = "C" THEN
 IF VAL(RIGHT$(invdate$, 2)) = VAL(RIGHT$(paydate$, 2)) THEN
  payments# = payments# + (CVD(amount$) * -1)
  payments# = INT((payments# * 100) + .5) / 100
 END IF
END IF
tr = CVS(ndata$): GOTO nexone

6000 CLS : COLOR 10: PRINT "Sys. Date: "; DATE$;
     COLOR 11: PRINT TAB(30); "Set Default Distribution";
     COLOR 10: PRINT TAB(70); "1.06.00": PRINT STRING$(79, "=")
6010 COLOR 11: LOCATE 3, 1
     PRINT "Vendor Account Number (ESC-Menu:TAB-Name Search) .. ";
     fl = -5: GOSUB 62000
     IF cf = 1 THEN CLOSE : GOTO 10
     IF cf = 2 THEN
      LOCATE 4, 1: COLOR 11
      PRINT "Enter Name (or part) .. "; : fl = 30: GOSUB 62000
      searchnam$ = in$
      GOSUB 10000
      FOR x = 1 TO LOF(1) / 384: GET 1, x
          IF LEFT$(venname$, LEN(searchnam$)) = searchnam$ THEN
           COLOR 10: LOCATE 6, 1: PRINT "Account of "; venname$
           LOCATE 7, 1: COLOR 11
           PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210: IF in$ = "Y" THEN f1 = x: GOTO 6030 ELSE IF in$ = "Q" THEN CLOSE : GOTO 6000
          END IF
      NEXT x: CLOSE : GOTO 6000
          
     END IF
     account = VAL(in$)
     GOSUB 10000: f1 = account

6030 IF f1 < 1 OR f1 > LOF(1) / 384 THEN 6070
6040 x = f1: GET 1, f1: GET 7, f1
     IF LEFT$(venname$, 6) = "delete" THEN 6070 ELSE GOTO 6100
6070 CLOSE
6080 COLOR 12: PRINT "Account Not on File, Press any key to Continue .. "; : fl = 1: GOSUB 62000: CLOSE : GOTO 6000
6100 LOCATE 3, 1: FOR xx = 3 TO 15: LOCATE xx, 1: PRINT STRING$(79, 32): NEXT xx: LOCATE 3, 1
6110 COLOR 11: PRINT "Vendor Account Number: "; : COLOR 15: PRINT CVI(VENNUMBER$); : COLOR 11: PRINT TAB(35); "Vendor Name: "; : COLOR 15: PRINT venname$
COLOR 15
FOR xx = 8 TO 21: LOCATE xx, 1
PRINT USING "##"; xx
NEXT xx
COLOR 11
LOCATE 7, 1: COLOR 11
PRINT TAB(5); "Account #  /  Name"; TAB(48); "Percentage"; TAB(62); "Running % = "
k = 8
begentry:
   IF k > 21 THEN GOTO endentry
   LOCATE k, 5: PRINT STRING$(74, 32): LOCATE k, 5: PRINT ""; : fl = -9: GOSUB 62000: IF cf = 1 THEN perc(k - 7) = 0: LOCATE k, 5: PRINT STRING$(12, 32): GOTO endentry
   LOCATE k, 5: IF LEN(in$) = 7 THEN in$ = in$ + "  "
   PRINT LEFT$(in$, 2); "-"; MID$(in$, 3, 3); "-"; MID$(in$, 6, 2); "-"; RIGHT$(in$, 2);
   ta$(k - 7) = in$: ACCT# = VAL(in$): GOSUB 17500
   IF NG = 0 THEN COLOR 12: LOCATE k, 20: PRINT "ACCT NOT LOCATED, PRESS ANY KEY TO CONTINUE "; : fl = 1: GOSUB 62000: GOTO begentry
   COLOR 11: GOSUB 12000: GET 14, record: LOCATE k, 20: PRINT LEFT$(ACCTNAME$, 22);
   ta#(k - 7) = ACCT#: tn$(k - 7) = ACCTNAME$
   LOCATE k, 50: PRINT ""; : fl = -7: GOSUB 62000: perc(k - 7) = VAL(in$)
   LOCATE k, 50: PRINT USING "###.##"; perc(k - 7)

runperc = 0
FOR LL = 1 TO 14: runperc = runperc + perc(LL): NEXT LL
LOCATE 7, 74: COLOR 15: PRINT USING "###.##"; runperc
IF runperc = 100 THEN GOTO endentry ELSE k = k + 1: GOTO begentry

endentry:
runperc = 0
FOR LL = 1 TO 14: runperc = runperc + perc(LL): NEXT LL
LOCATE 7, 72: COLOR 15: PRINT USING "###.##"; runperc

LOCATE 23, 1: PRINT STRING$(79, 32): LOCATE 23, 1: COLOR 11
PRINT "Enter Selection # (ESC-OK : F1-CANCEL) ... "; : fl = -2
GOSUB 62000
IF cf = 11 THEN CLOSE : GOTO 10
IF cf = 1 THEN GOSUB storeit: CLOSE : GOTO 10
k = VAL(in$)
IF k < 8 OR k > 21 THEN GOTO endentry
editentry:
   LOCATE k, 5: PRINT STRING$(74, 32): LOCATE k, 5: PRINT ""; : fl = -9: GOSUB 62000: IF cf = 1 THEN perc(k - 7) = 0: LOCATE k, 5: PRINT STRING$(12, 32): GOTO endentry
   LOCATE k, 5: IF LEN(in$) = 7 THEN in$ = in$ + "  "
   PRINT LEFT$(in$, 2); "-"; MID$(in$, 3, 3); "-"; MID$(in$, 6, 2); "-"; RIGHT$(in$, 2);
   ta$(k - 7) = in$: ACCT# = VAL(in$): GOSUB 17500
   IF NG = 0 THEN COLOR 12: LOCATE k, 20: PRINT "ACCT NOT LOCATED, PRESS ANY KEY TO CONTINUE "; : fl = 1: GOSUB 62000: GOTO editentry
   COLOR 11: GOSUB 12000: GET 14, record: LOCATE k, 20: PRINT LEFT$(ACCTNAME$, 22);
   ta#(k - 7) = ACCT#: tn$(k - 7) = ACCTNAME$
   LOCATE k, 50: PRINT ""; : fl = -7: GOSUB 62000: perc(k - 7) = VAL(in$)
   LOCATE k, 50: PRINT USING "###.##"; perc(k - 7)
   GOTO endentry



storeit:
IF runperc <> 100 THEN
 PRINT CHR$(7);
 LOCATE 23, 1: COLOR 11: PRINT "Percentage Does Not Equal 100 : Please Correct : Press <ENTER> to Continue "; : fl = 1
 GOSUB 62000: GOTO endentry
END IF
OPEN "R", 9, "apinvdis.dat", 13
FIELD 9, 9 AS apacct$, 4 AS apperc$
FOR xx = 1 TO 14
LSET apacct$ = ta$(xx): LSET apperc$ = MKS$(perc(xx))
PUT 9, LOF(9) / 13 + 1
IF xx = 1 THEN firstrecord = LOF(9) / 13
NEXT xx
LSET FIRSTAPDIST$ = MKS$(firstrecord): PUT 7, f1
CLOSE
RETURN




10000 CLOSE 1: OPEN "apvendor.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 384
10010 FIELD 1, 2 AS VENNUMBER$, 8 AS notused$, 30 AS venname$, 30 AS addRESS$, 22 AS CitY$, 2 AS St$, 10 AS ZiP$, 30 AS payto$, 30 AS payadd$, 22 AS paycity$, 2 AS payst$, 10 AS payzip$, 8 AS curbal$, 4 AS first$, 4 AS last$, 30 AS address2$, 30 AS  _
paddress2$
      CLOSE 7: OPEN "apvendr1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #7 LEN = 256
      FIELD 7, 12 AS TEL$, 4 AS YTDPUR$, 4 AS YTDPAY$, 12 AS FEDID$, 3 AS ctycode$, 2 AS stcode$, 4 AS POFIRST$, 4 AS POLAST$, 4 AS FIRSTAPDIST$, 1 AS ten99$, 206 AS bk$
      RETURN
11000 CLOSE 2: OPEN "apledger.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 96
11010 FIELD 2, 2 AS vennum$, 8 AS invdate$, 25 AS invnum$, 10 AS ponum$, 8 AS amount$, 2 AS PAYCODE$, 4 AS FDIST$, 4 AS LDIST$, 4 AS ndata$
11020 RETURN

12000 CLOSE 14
      OPEN "baacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = 128
      FIELD 14, 12 AS MASTACCT$, 35 AS ACCTNAME$, 1 AS ACCTTYPE$, 8 AS acctbal$, 8 AS mtdact$, 8 AS ytdact$, 8 AS BUDAMT$, 4 AS first$, 4 AS last$, 8 AS begbal$, 8 AS recalcbal$, 8 AS encumbamt$, 16 AS blank$
      RETURN

62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
      IF LEN(z$) <> 2 THEN 62065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN cf = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN cf = 12: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 61 THEN cf = 13: RETURN

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

fixerror:
     REM ERROR CODES
     Number = ERR
     SELECT CASE ERR
          CASE IS = 24, 25
               LOCATE 23, 1: PRINT STRING$(79, 32): LOCATE 23, 1: COLOR 28
               PRINT "Error Message:"; : COLOR 15: PRINT "Device Time-out! Make Sure Printer Is ON_LINE (or Selected): Press ENTER to Resume Printing";
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

17500 OPEN "baacctdx.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #9 LEN = 12
     FIELD 9, 8 AS ACT$, 4 AS REC$
     IF ACCT# < 9999999 THEN ACCT# = ACCT# * 100
17520 IF LOF(9) / 12 = 0 THEN NG = 0: CLOSE 9: RETURN
17540 HI = LOF(9) / 12: LO = 1: ME = HI
17560 FOR xx = 1 TO 16: GET 9, ME
17580 IF ACCT# = CVD(ACT$) THEN NG = 1: record = CVS(REC$): CLOSE 9: RETURN
17600 IF ACCT# > CVD(ACT$) THEN LO = ME ELSE HI = ME
17620 ME = INT((HI - LO) / 2) + LO
17640 NEXT xx: CLOSE 9: NG = 0: RETURN

