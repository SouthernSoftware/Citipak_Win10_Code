0 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
10 REM
20 REM wbbilling (ELLENBORO, NC)
30 REM Water Billing Menu
40 REM Last Revision Date: 9-28-88
50 REM by Nick Picerno
60 CLS : DIM SHARED BD$(12), CR$(12), PR$(12), BT$(12), ocharge#(8)
   ON ERROR GOTO fixerror
   WIDTH LPRINT 132
100 GOSUB 63000: REM get town in town$
110 SCREN$ = "3.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500
130 LOCATE 4, 1: COLOR 11
200 PRINT TAB(29); "Utility Billing System"
210 PRINT TAB(30); "       Billing"
220 PRINT : PRINT
230 PRINT TAB(25); "1. Enter Meter Readings"
240 PRINT
245 PRINT TAB(25); "2. Calculate and Print Pre-Billing Report"
    PRINT
250 PRINT TAB(25); "3. Print Utility Bills For Customers"
260 PRINT
270 PRINT TAB(25); "4. Print Billing Report Summary"
280 PRINT
290 PRINT TAB(25); "5. Age Accounts"
300 PRINT
302 PRINT TAB(25); "6. Apply Late Fee"
303 PRINT
304 PRINT TAB(25); "7. PC/3000 Hand Held Processing"
305 PRINT
310 LOCATE 22, 1: COLOR 11: PRINT TAB(25); "Selection No. (ESC-Master Menu) : "; : fl = -1: GOSUB 62000
320 IF cf = 1 THEN CLOSE : RUN "wbmenu"
325 IF LEN(in$) = 0 THEN 310
330 IF VAL(in$) < 1 OR VAL(in$) > 7 THEN 310
340 ON VAL(in$) GOTO 1000, 7000, 2000, 3000, 4000, 5000, 6000
410 RUN "wbcustmaint"

1000 REM METER READING SUB-ROUTINE
1010 SCREN$ = "3.01.00": CLS : GOSUB 63500
1020 COLOR 11: LOCATE 1, 32: PRINT "Enter Meter Readings"; : LOCATE 3, 1
1030 COLOR 11: PRINT "Enter Book Number to Process (ESC-Quit) .. "; : fl = -2: GOSUB 62000: IF cf = 1 THEN 10
1035 BOOK = VAL(in$)
1040 LOCATE 4, 1: PRINT STRING$(79, "=")
1050 GOSUB 60000: OPEN "R", 4, "wbmeter.dat", 16: FIELD 4, 12 AS MET$, 4 AS mrec$
1060 FOR x = 1 TO LOF(4) / 16: GET 4, x: IF CVS(mrec$) = 0 THEN 1080 ELSE GET 1, CVS(mrec$)
1070 IF VAL(LEFT$(MET$, 2)) = BOOK THEN GOSUB 1100: GOSUB 1300
1080 NEXT x: LOCATE 22, 1: CLOSE : COLOR 11: PRINT "That's all for this book! Do you wish to key another book (Y/N) .. "; : fl = 1: GOSUB 62210: IF NOT (in$ = "Y" OR in$ = "N") THEN 1080
1090 IF in$ = "Y" THEN 1010 ELSE 10
1100 LOCATE 5, 1: B = INSTR(firstname$, "  "): IF B = 1 THEN B = 0: GOTO 1101
     IF B = 0 THEN B = 20
1101 z = INSTR(meter$, "/"): IF z = 0 THEN z = 9
1110 COLOR 11: PRINT LEFT$(firstname$, B); " "; lname$; TAB(55); LEFT$(meter$, z - 1)
1120 LOCATE 7, 1: PRINT "Previous Reading ........ "; USING "##########"; CVD(CURREAD$)
1130 LOCATE 8, 50: PRINT "Press ESC to Skip This Acct"; : LOCATE 8, 1: PRINT "Enter Current Reading ... "; : fl = -10: GOSUB 62000: IF cf = 1 THEN RETURN
1131 CURREAD# = VAL(in$): LOCATE 8, 27: PRINT USING "##########"; CURREAD#
1140 LOCATE 9, 26: COLOR 11: PRINT "-----------"
1145 IF CURREAD# < CVD(CURREAD$) THEN GOSUB 1700 ELSE CALCUSAGE# = CURREAD# - CVD(CURREAD$): REM check for meter triping
1150 COLOR 10: LOCATE 10, 1: PRINT "Calculated Usage Is ..... "; USING "##########"; CALCUSAGE#
1152 LOCATE 11, 1: PRINT "Correct (ESC-NO) (Enter-OK) .. "; : fl = 1: GOSUB 62000: IF cf = 1 THEN GOSUB 1300: GOTO 1100
1160 IF CVI(AVGCT$) = 0 THEN 1170 ELSE GOSUB 1500: REM check for out of bounds
1165 IF BG = 1 THEN BG = 0: GOSUB 1300: GOTO 1100
1170 LSET PASTREAD$ = MKD$(CVD(CURREAD$)): LSET CURREAD$ = MKD$(CURREAD#): PUT 1, CVS(mrec$): RETURN
1300 LOCATE 5, 1: PRINT STRING$(62, 32): LOCATE 7, 26: PRINT STRING$(15, 32); : LOCATE 8, 26: PRINT STRING$(15, 32): LOCATE 10, 26: PRINT STRING$(15, 32): LOCATE 12, 1: PRINT STRING$(45, 32)
1310 RETURN

1500 REM check for avg use
1505 AVG = CVD(AVGUSE$): VAR = AVG + (AVG * .75)
1510 IF CALCUSAGE# > VAR THEN PRINT CHR$(7): LOCATE 14, 1: COLOR 11: PRINT "Please Confirm This : Usage Exceeds Normal Usage by More Than 75%": LOCATE 15, 1: PRINT "Is this entry correct (Y/N) .. "; : fl = 1: GOSUB 62210: IF NOT (in$ = "Y" OR in$ =  _
"N") THEN 1510
1515 IF in$ = "Y" THEN LOCATE 14, 1: PRINT STRING$(70, 32): LOCATE 15, 1: PRINT STRING$(70, 32): BG = 0: RETURN
1520 IF in$ = "N" THEN LOCATE 14, 1: PRINT STRING$(70, 32): LOCATE 15, 1: PRINT STRING$(70, 32): BG = 1: RETURN
1530 RETURN

1700 REM check for meter triping
1710 J = LEN(STR$(CVD(CURREAD$))) - 1
1715 MAX = 10 ^ J
1720 CALCUSAGE# = MAX - CVD(CURREAD$) + CURREAD#
1725 RETURN
1799 RETURN

2000 REM PRINT BILLS SUB-ROUTINE

2001 iflag = 0
2010 SCREN$ = "3.03.00": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 28: PRINT "Calculate & Print Water Bills"
    
2020 COLOR 11: LOCATE 3, 1: PRINT "(I)ndividual Bill or (A)ll Bills (ESC-Quit) .. "; : fl = 1: GOSUB 62210: IF cf = 1 THEN 10
2030 IF NOT (in$ = "I" OR in$ = "A") THEN 2020
 LOCATE 3, 48: PRINT in$
 IF in$ = "I" THEN iflag = 1 ELSE iflag = 0
2035 COLOR 11: LOCATE 10, 1: PRINT "ENTER (R)eprint Bill Only  or  (F)ull Bill Processing (R or F) .. "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN CLOSE : GOTO 10
     IF NOT (in$ = "R" OR in$ = "F") THEN 2035
     billinfo$ = in$

2100 REM get general billing information
2110 LOCATE 9, 1: COLOR 11: PRINT STRING$(79, "-"): PRINT "General Billing Information ... "; TAB(55); "ESC-Abort Process"
2130 LOCATE 13, 1: PRINT "Enter the Service To Date (mm/dd/yy) .... "; : fl = 8: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10
      IF LEN(in$) <> 8 THEN 2130 ELSE SERTO$ = in$
     COLOR 11
2140 LOCATE 14, 1: PRINT "Enter the Billing Date (mm/dd/yy) ....... "; : fl = 8: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10
      IF LEN(in$) <> 8 THEN 2140 ELSE BILLDATE$ = in$
     COLOR 11
2150 LOCATE 15, 1: PRINT "Enter the Due Date (mm/dd/yy) ........... "; : fl = 8: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10
      IF LEN(in$) <> 8 THEN 2150 ELSE DUEDATE$ = in$
     COLOR 11
2155 LOCATE 16, 1: PRINT "Enter Water Test Fee Amount (ESC-MENU) .. "; : fl = -8
      GOSUB 62000: IF cf = 1 THEN 10
      testfee = VAL(in$)
      COLOR 11
2157 LOCATE 17, 1: PRINT "Enter Garbage Fee Amount (ESC-MENU) ..... "; : fl = -8
      GOSUB 62000: IF cf = 1 THEN 10
      trash = VAL(in$)


2160 COLOR 11: LOCATE 18, 1: PRINT "Are the above answers correct (Y/N) ..... "; : fl = 1: GOSUB 62210: IF NOT (in$ = "Y" OR in$ = "N") THEN 2160
2170 IF in$ = "N" THEN FOR J = 12 TO 18: LOCATE J, 1: PRINT STRING$(60, 32): NEXT J: GOTO 2100
2200 REM all bills
2201 LOCATE 19, 1: COLOR 10: PRINT STRING$(79, "-"): COLOR 11: LOCATE 20, 1: PRINT "Do you wish to print an alignment mask (Y/N) .. "; : fl = 1: GOSUB 62210: IF NOT (in$ = "Y" OR in$ = "N") THEN 2201 ELSE LOCATE 20, 49: PRINT in$
2202 IF in$ = "Y" THEN GOSUB 2800: GOTO 2201
2210 IF iflag = 1 THEN 2300
2211 LOCATE 20, 1: PRINT STRING$(75, 32)
     LOCATE 21, 1: PRINT STRING$(79, 32)
     LOCATE 20, 1: PRINT "Enter Meter # to Begin (0 for all) .. "; : fl = -2
     GOSUB 62000: IF VAL(in$) = 0 THEN ME = 1: GOTO 2218
2212 meters$ = in$ + "-"
2213 LOCATE 20, 41
     PRINT "-"; : fl = 6: GOSUB 62000: IF cf = 1 THEN START$ = "": GOTO 2211
2214 meters$ = meters$ + in$: IF LEN(meters$) < 8 THEN meters$ = meters$ + STRING$(8 - LEN(meters$), 47)
     GOSUB 20000: IF NG = 0 THEN LOCATE 21, 1: PRINT "Could Not Locate this Meter Number to Begin : Press any key to Redo "; : fl = 1: GOSUB 62000: meters$ = "": GOTO 2211
2215 IF billinfo$ = "F" THEN CLOSE : OPEN "R", 4, "wbsumary.dat", 128: FIELD 4, 8 AS SMET$, 30 AS SNAME$, 8 AS SUSAGE$, 4 AS SPASTDUE$, 8 AS SCURAMT$, 8 AS STOTAL$, 2 AS SRATE$, 4 AS revenue1$, 4 AS revenue2$, 4 AS revenue3$, 4 AS revenue4$, 4 AS  _
revenue5$, 4 AS revenue6$, 1 AS citylimit$, 4 AS revenue7$, 4 AS revenue8$: GOTO 2220
2218 IF billinfo$ = "F" THEN CLOSE : OPEN "R", 4, "wbsumary.dat", 128: CLOSE 4: KILL "wbsumary.dat": OPEN "R", 4, "wbsumary.dat", 128: FIELD 4, 8 AS SMET$, 30 AS SNAME$, 8 AS SUSAGE$, 4 AS SPASTDUE$, 8 AS SCURAMT$, 8 AS STOTAL$, 2 AS SRATE$, 4 AS  _
revenue1$, 4 AS revenue2$, 4 AS revenue3$, 4 AS revenue4$, 4 AS revenue5$, 4 AS revenue6$, 1 AS citylimit$, 4 AS revenue7$, 4 AS revenue8$
2220 GOSUB 60000: GOSUB 60100: OPEN "R", 3, "wbmeter.dat", 16: FIELD 3, 12 AS MET$, 4 AS mrec$
2230 FOR x = ME TO LOF(3) / 16: GET 3, x
2240 IF CVS(mrec$) = 0 OR LEFT$(MET$, 1) = " " THEN 2270
2245 acctrecord = CVS(mrec$)
2250 GET 1, acctrecord: GET 6, acctrecord
     IF CVI(Rate$) < 1 AND CVD(PASTDUE$) = 0 THEN 2270
2252 bb = INSTR(meter$, "/"): IF bb = 0 THEN bb = 9
2253 m$ = LEFT$(meter$, bb - 1)
2255 GOSUB 2400: REM Check Rate Codes
     IF CVI(Rate$) = 1 THEN GOSUB 41000
     IF CVI(Rate$) = 2 THEN GOSUB 42000
     IF CVI(Rate$) = 3 THEN GOSUB 43000   'New Outside Rate Added 7-12-99
     wateramount# = INT((wateramount# * 100) + .5) / 100
     IF LEFT$(watertest$, 1) = "Y" THEN testamount# = testfee ELSE testamount# = 0
     IF LEFT$(sanitation$, 1) = "Y" THEN trashamount# = trash ELSE trashamount# = 0

2265 GOSUB 2500: REM print bill
     IF billinfo$ = "R" THEN 2270
2267 GOSUB 2700: REM update all accounts
2269 GOSUB 2900: REM update summary report
2270 wateramount# = 0
     AMOUNT# = 0
     testamount# = 0
     miscamount# = 0: trashamount# = 0
     FOR jj = 1 TO 5: ocharge#(jj) = 0: NEXT jj
     NEXT x: LPRINT CHR$(12); : CLOSE : LOCATE 22, 1: PRINT "Water Bills Now Completed, Press Any Key to Continue "; : fl = 1: GOSUB 62000: CLOSE : GOTO 10
2300 REM one account only
             FOR xx = 3 TO 22: LOCATE xx, 1: PRINT STRING$(79, 32): NEXT xx
             LOCATE 4, 1: iflag = 1: COLOR 11
             PRINT "Enter Customer Acct Number (ESC-End) .. "; : fl = -5: GOSUB 62000
                   IF cf = 1 THEN CLOSE : GOTO 10
                   record = VAL(in$): GOSUB 60000
                   IF record < 1 OR record > LOF(1) / 256 THEN CLOSE : GOTO 2300
                   CLOSE : GOTO 2304
                 
2304 GOSUB 60000: GOSUB 60100: GET 6, record: GET 1, record: LOCATE 6, 1: COLOR 10: COLOR 11: PRINT "Account of "; lname$; ", "; firstname$
2305 LOCATE 7, 1: COLOR 11: PRINT "Is this the Correct Account (Y/N) .. "; : fl = 1: GOSUB 62210: IF NOT (in$ = "Y" OR in$ = "N") THEN 2305
             LOCATE 7, 38: PRINT in$
             IF in$ = "N" THEN CLOSE : GOTO 2300

2310 acctrecord = record
2320 bb = INSTR(meter$, "/"): IF bb = 0 THEN bb = 9
2330 m$ = LEFT$(meter$, bb - 1)
2340 GOSUB 2400
 
2350 REM Check rate codes
     IF CVI(Rate$) = 1 THEN GOSUB 41000
     IF CVI(Rate$) = 2 THEN GOSUB 42000
     wateramount# = INT((wateramount# * 100) + .5) / 100
     IF LEFT$(watertest$, 1) = "Y" THEN testamount# = testfee ELSE testamount# = 0
     IF LEFT$(sanitation$, 1) = "Y" THEN trashamount# = trash ELSE trashamount# = 0

2360 GOSUB 2500
     IF billinfo$ = "R" THEN 2370
     GOSUB 2700: OPEN "R", 4, "wbsumary.dat", 128: FIELD 4, 8 AS SMET$, 30 AS SNAME$, 8 AS SUSAGE$, 4 AS SPASTDUE$, 8 AS SCURAMT$, 8 AS STOTAL$, 2 AS SRATE$, 4 AS revenue1$, 4 AS revenue2$, 4 AS revenue3$, 4 AS revenue4$, 4 AS revenue5$: GOSUB 2900
2370 FOR jj = 1 TO 8: ocharge#(jj) = 0: NEXT jj
     wateramount# = 0: AMOUNT# = 0: testamount# = 0: seweramount# = 0
     miscamount# = 0: trashamount# = 0
     CLOSE : GOTO 2300

2400 REM figure usage
2405 CURRENT# = 0: PAST# = 0: CALCUSAGE# = 0: miscamount# = 0: seweramount# = 0
      AMOUNT# = 0: wateramount# = 0: testamount# = 0: trashamount# = 0
2410 CURRENT# = CVD(CURREAD$): PAST# = CVD(PASTREAD$)
2420 IF CURRENT# < PAST# THEN 2450
2430 CALCUSAGE# = CURRENT# - PAST#
     RETURN
2450 J = LEN(STR$(PAST#)) - 1: MAX = 10 ^ J
2460 CALCUSAGE# = MAX - PAST# + CURRENT#
2470 RETURN

2500 REM print bill
     lc = 0
     GOSUB 60100: GET 6, acctrecord
     IF CVS(RAMT1$) = 0 THEN 2501
      oamt# = CVS(RAMT1$): SOURCE = VAL(LEFT$(RACCT1$, 1)): GOSUB getocharge
     IF CVS(RAMT2$) = 0 THEN 2501
      oamt# = CVS(RAMT2$): SOURCE = VAL(LEFT$(RACCT2$, 1)): GOSUB getocharge
     IF CVS(RAMT3$) = 0 THEN 2501
      oamt# = CVS(RAMT3$): SOURCE = VAL(LEFT$(RACCT3$, 1)): GOSUB getocharge
2501 IF CVS(OAMT1$) = 0 THEN 2505
      oamt# = CVS(OAMT1$): SOURCE = VAL(LEFT$(OACCT1$, 1)): GOSUB getocharge
      oamt# = CVS(OAMT2$): SOURCE = VAL(LEFT$(OACCT2$, 1)): GOSUB getocharge
2505 wateramount# = wateramount# + ocharge#(1)
     testamount# = testamount# + ocharge#(2)
     penaltyamount# = penaltyamount# + ocharge#(3)
     trashamount# = trashamount# + ocharge#(4)
     IF CVI(Rate$) = 0 THEN wateramount# = 0: testamount# = 0: trashamount# = 0: penaltyamount# = 0
     lc = 0
2510 FOR LOP = 1 TO 6: LPRINT "": NEXT LOP
2520 LPRINT TAB(13); LEFT$(SERTO$, 2); TAB(17); MID$(SERTO$, 4, 2); TAB(22); LEFT$(DUEDATE$, 2); TAB(27); MID$(DUEDATE$, 4, 2); TAB(32); RIGHT$(DUEDATE$, 2)
     IF LEFT$(firstname$, 2) = "  " THEN NME$ = lname$: GOTO 2570
     B = INSTR(firstname$, "  "): IF B = 0 THEN B = 20
     NME$ = LEFT$(firstname$, B) + lname$
2570 LPRINT TAB(36); LEFT$(NME$, 20)
2580 LPRINT TAB(36); LEFT$(ADD1$, 20)
2590 IF LEFT$(ADD2$, 1) = " " THEN EXTRA = 1 ELSE EXTRA = 0
2600 IF EXTRA = 0 THEN LPRINT TAB(36); LEFT$(ADD2$, 20)
2610 LPRINT TAB(36); LEFT$(CITY$, 14); " "; ST$; " "; ZP$
2630 IF EXTRA = 1 THEN LPRINT ""
2635 LPRINT USING "########"; PAST#;
      LPRINT TAB(9); USING "########"; CURRENT#;
       LPRINT TAB(18); USING "######"; CALCUSAGE#;
        LPRINT TAB(25); "WA"; TAB(28); USING "$###.##"; wateramount#
        
2637 PASTDUE# = INT((CVD(PASTDUE$) * 100) + .5) / 100
     IF PASTDUE# = 0 THEN
      PFLAG = 0
      ELSE
      PFLAG = 1
      LPRINT TAB(25); "AR"; TAB(28); USING "$###.##"; CVD(PASTDUE$)
     END IF
     IF testamount# = 0 THEN
      LPRINT
      ELSE
      LPRINT TAB(25); "WT"; TAB(28); USING "$###.##"; testamount#
     END IF
     IF trashamount# = 0 THEN
      LPRINT
      ELSE
      LPRINT TAB(25); "SA"; TAB(28); USING "$###.##"; trashamount#
     END IF
2638 IF sFLAG = 0 THEN LPRINT
2639 IF PFLAG = 0 THEN LPRINT
2640 FOR LOP = 1 TO 2: LPRINT "": NEXT LOP
2660 AMOUNT# = wateramount# + PASTDUE# + trashamount# + testamount#
2680 LPRINT TAB(14); acctrecord;
      LPRINT TAB(27); USING "$####.##"; AMOUNT#;
       LPRINT TAB(39); USING "$####.##"; AMOUNT#;
        LPRINT TAB(49); acctrecord
2690 LPRINT : LPRINT TAB(14); m$; TAB(48); DUEDATE$: LPRINT : RETURN

getocharge:
IF SOURCE < 1 OR SOURCE > 4 THEN SOURCE = 4
     ocharge#(SOURCE) = ocharge#(SOURCE) + oamt#
     RETURN

2700 REM update account and transaction file
2705 GOSUB 60200: GET 7, acctrecord: GET 9, acctrecord
     FOR l = 0 TO 10: LSET BD$(l) = BD$(l + 1): LSET CR$(l) = MKD$(CVD(CR$(l + 1)))
                      LSET PR$(l) = MKD$(CVD(PR$(l + 1)))
                      LSET BT$(l) = MKS$(CVS(BT$(l + 1))): NEXT l
     LSET BD$(11) = BILLDATE$
     LSET CR$(11) = MKD$(CURRENT#)
     LSET PR$(11) = MKD$(PAST#)
     LSET BT$(11) = MKS$(AMOUNT#): PUT 7, acctrecord
2710 USED# = CVI(AVGCT$) * CVD(AVGUSE$): USED# = USED# + CALCUSAGE#
     COUNTER = CVI(AVGCT$) + 1
     LSET AVGUSE$ = MKD$((USED# / COUNTER)): LSET AVGCT$ = MKI$(COUNTER)
2720 LSET curbal$ = MKD$(wateramount# + seweramount# + miscamount# + testamount# + trashamount#)
     IF CYCLE$ = "F" THEN LSET CYCLE$ = "B": LSET Rate$ = MKI$(0)
2730 PUT 1, acctrecord
     GET 9, acctrecord
     LSET IRev1$ = MKD$(CVD(IRev1$) + wateramount#)
     LSET IRev2$ = MKD$(CVD(IRev2$) + testamount#)
     LSET irev3$ = MKD$(CVD(irev3$) + penaltyamount#)
     LSET irev4$ = MKD$(CVD(irev4$) + trashamount#)
     LSET IRev5$ = MKD$(0)
     LSET IRev6$ = MKD$(0)
     LSET irev7$ = MKD$(0)
     LSET irev8$ = MKD$(0)
     PUT 9, acctrecord
2740 GOSUB 61000: NEXRECORD = LOF(2) / 50 + 1
2750 LSET TDATE$ = BILLDATE$: LSET TAMT$ = MKD$(AMOUNT#): LSET TTYPE$ = MKI$(1): LSET TDESC$ = "Water Bill": LSET tmrec$ = MKS$(acctrecord): LSET tnexttr$ = MKS$(0):
     LSET tcuramt$ = MKS$(wateramount# + seweramount# + testamount# + miscamount# + trashamount#): PUT 2, NEXRECORD
2760 IF CVD(firsttr$) = 0 THEN LSET firsttr$ = MKD$(NEXRECORD): LSET lasttr$ = MKD$(NEXRECORD): PUT 1, acctrecord: CLOSE 2: RETURN
2770 PREV = CVD(lasttr$): IF PREV = 0 THEN PREV = CVD(firsttr$)
     LSET lasttr$ = MKD$(NEXRECORD): PUT 1, acctrecord: GET 2, PREV: LSET tnexttr$ = MKS$(NEXRECORD): PUT 2, PREV: CLOSE 2: RETURN
2800 REM alignment mask
2810 FOR LOP = 1 TO 6: LPRINT "": NEXT LOP
     LPRINT TAB(13); "XX  XX"; TAB(22); "XX   XX   XX"
2830 LPRINT TAB(36); "XXXXXXXXXXXXXXXXXXXXXXXXX"
2840 LPRINT TAB(36); "XXXXXXXXXXXXXXXXXXXXXXXXX"
2850 LPRINT TAB(36); "XXXXXXXXXXXXXXXXXXXXXX XX"
2855 LPRINT TAB(36); "XXXXXXXXXXXXXXXXXXXXXX XX"
2860 LPRINT "XXXXXXX XXXXXXX  XXXXXX XX XXXX.XX"
2870 FOR LOP = 1 TO 6: LPRINT "": NEXT LOP
2880 LPRINT TAB(14); "XXXXXX"; TAB(27); "XXXXX.XX"; TAB(39); "XXXXX.XX"; TAB(49); "XXXXXX"
2890 LPRINT : LPRINT : LPRINT : RETURN

2900 REM update wbsumary.dat
2910 LSET SMET$ = STR$(acctrecord)
     LSET SNAME$ = NME$
     LSET SUSAGE$ = MKD$(CALCUSAGE#)
     LSET SPASTDUE$ = MKS$(PASTDUE#)
     LSET SCURAMT$ = MKD$(wateramount#)
     LSET STOTAL$ = MKD$(AMOUNT#)
     LSET SRATE$ = MKI$(CVI(Rate$))
     LSET revenue1$ = MKS$(wateramount#)
     LSET revenue2$ = MKS$(testamount#)
     LSET revenue3$ = MKS$(penaltyamount#)
     LSET revenue4$ = MKS$(trashamount#)
     LSET revenue5$ = MKS$(0)
     LSET citylimit$ = ""
     PUT 4, LOF(4) / 128 + 1: RETURN

3000 REM SUMMARY REPORT SUB-ROUTINE
3001 TOTGAL1# = 0: TOTGAL2# = 0: t1# = 0: t2# = 0: t3# = 0: t4# = 0: t5# = 0: t6# = 0: t7# = 0: tc# = 0: tp# = 0
     billsin = 0: billsout = 0: totpast# = 0: pastcr# = 0
3010 SCREN$ = "3.04.00": CLS : GOSUB 63500: LOCATE 1, 30: COLOR 11: PRINT "Water Billing Summary Report"
3020 COLOR 11: LOCATE 3, 1: PRINT "Press any key for summary report (ESC-Menu) .. "; : fl = 1: GOSUB 62000: IF cf = 1 THEN 10
3030 OPEN "R", 4, "wbsumary.dat", 128
     FIELD 4, 8 AS SMET$, 30 AS SNAME$, 8 AS SUSAGE$, 4 AS SPASTDUE$, 8 AS SCURAMT$, 8 AS STOTAL$, 2 AS SRATE$, 4 AS revenue1$, 4 AS revenue2$, 4 AS revenue3$, 4 AS revenue4$, 4 AS revenue5$, 4 AS revenue6$, 1 AS citylimit$, 4 AS revenue7$, 4 AS  _
revenue8$
3040 IF LOF(4) / 128 = 0 THEN CLOSE : LOCATE 5, 1: COLOR 12: PRINT "Nothing in File to Print, Press Any Key to Continue "; : fl = 1: GOSUB 62000: CLOSE : GOTO 10
3050 GOSUB 3900
3060 FOR x = 1 TO LOF(4) / 128: GET 4, x
3065 IF lc >= 56 THEN LPRINT CHR$(12); : GOSUB 3900
3070 LPRINT SMET$; TAB(10); SNAME$; TAB(65); USING "**$#####,#.##"; CVD(STOTAL$)
     LPRINT TAB(5); USING "$$#####.##"; CVS(revenue1$);
     LPRINT TAB(15); USING "$$#####.##"; CVS(revenue2$);
     LPRINT TAB(25); USING "$$#####.##"; CVS(revenue3$);
     LPRINT TAB(35); USING "$$#####.##"; CVS(revenue4$);
     LPRINT TAB(55); USING "$$#####.##"; CVS(SPASTDUE$)
     IF CVS(SPASTDUE$) < 0 THEN pastcr# = pastcr# + CVS(SPASTDUE$)
     IF CVS(SPASTDUE$) > 0 THEN totpast# = totpast# + CVS(SPASTDUE$)
     t1# = t1# + CVS(revenue1$): t2# = t2# + CVS(revenue2$): t3# = t3# + CVS(revenue3$)
     t4# = t4# + CVS(revenue4$): t5# = t5# + CVS(revenue5$): t6# = t6# + CVS(revenue6$)
     t7# = t7# + CVS(revenue7$): t8# = t8# + CVS(revenue8$)
3071 tp# = tp# + CVS(SPASTDUE$): tc# = tc# + CVD(STOTAL$): LPRINT STRING$(79, "-")
3075 TOTGAL1# = TOTGAL1# + CVD(SUSAGE$):  lc = lc + 3
                          
3080 NEXT x: LPRINT CHR$(12);
3090 LPRINT "Utility Billing Summary Page     :       Date "; DATE$
3100 LPRINT
3110 REM
3130 LPRINT "Total Gallons Billed ............ "; USING "########,#"; TOTGAL1# + TOTGAL2#
3140 LPRINT
3150 LPRINT "Total Current Billing ........... "; USING "$$#####,.##"; tc#
     LPRINT
     LPRINT "OF WHICH ..."
3160 LPRINT "Total Past Due .................. "; USING "$$#####,#.##"; totpast#
     LPRINT "Total Past Due Credits .......... "; USING "$$#####,#.##"; pastcr#
3166 LPRINT "Total Amount Due:Water .......... "; USING "$$#####,#.##"; t1#
     LPRINT "Total Amount Due:Water Test Fees. "; USING "$$#####,#.##"; t2#
     LPRINT "Total Amount Due:Misc ........... "; USING "$$#####,#.##"; t3#
     LPRINT "Total Amount Due:Trash .......... "; USING "$$#####,#.##"; t4#
     LPRINT
3180 LPRINT CHR$(12); : CLOSE : GOTO 10

3900 LPRINT TAB(26); "Water Billing Summary Report"
3910 LPRINT "Date: "; DATE$: LPRINT
3920 LPRINT "Loc. #"; TAB(10); "Name"
     LPRINT TAB(10); "Water"; TAB(17); "Test Fee"; TAB(29); "Pen'ty"; TAB(38); "Sanitat'n"; TAB(60); "Past Due"; TAB(72); "Total"
     LPRINT STRING$(79, "="): lc = 5: RETURN
4000 REM AGE ACCOUNTS SUB-ROUTINE
4010 SCREN$ = "3.05.00": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 33: PRINT "Age Accounts"
4020 COLOR 11: LOCATE 3, 1
4030 OPEN "R", 2, "wbage.dat", 8: FIELD 2, 8 AS AGEDATE$
4040 IF LOF(2) = 0 THEN AGE$ = "First Time": GOTO 4100
4050 GET 2, 1: AGE$ = AGEDATE$
4100 PRINT "You Last Aged Your Accounts ... "; : COLOR 10: PRINT AGE$
4110 COLOR 11: LOCATE 5, 1: PRINT "Press Enter to Age Accounts (ESC-ABORT) .. "; : fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10
4120 REM ok to age
4130 LSET AGEDATE$ = LEFT$(DATE$, 6) + RIGHT$(DATE$, 2): PUT 2, 1: CLOSE 2
4140 GOSUB 60000
4150 FOR x = 1 TO LOF(1) / 256: GET 1, x
4160 IF LEFT$(meter$, 1) = " " THEN 4200
4170 LOCATE 10, 1: PRINT "Now Processing Account of "; : B = INSTR(meter$, "/"): IF B = 0 THEN B = 9
4171 LOCATE 10, 31: PRINT STRING$(10, 32): LOCATE 10, 31
4175 PRINT LEFT$(meter$, B - 1)
4180 LSET PASTDUE$ = MKD$(CVD(PASTDUE$) + CVD(curbal$)): LSET curbal$ = MKD$(0)
4190 PUT 1, x
4200 NEXT x: CLOSE : LOCATE 20, 1: PRINT "All Done, Press Any Key to Continue "; : fl = 1: GOSUB 62000: GOTO 10
5000 REM late fee
5001 FEEFLAG = 0: PENFLAG = 0
5010 SCREN$ = "3.06.00": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 30: PRINT "Apply Penalty Fees"
5020 LOCATE 3, 1
5030 COLOR 11: PRINT "Enter Standard Penalty (<TAB> if % : ESC-Menu) .. "; : fl = -6: GOSUB 62000: IF cf = 1 THEN 10
     IF cf = 2 THEN LOCATE 5, 1: COLOR 11: PRINT "Enter Penalty Percentage .... "; : fl = -5: GOSUB 62000: PERCENT = VAL(in$): PENFLAG = 1
5040 IF PENFLAG = 0 THEN fee# = VAL(in$): LOCATE 3, 57: PRINT USING "$###.##"; fee# ELSE fee# = PERCENT: LOCATE 3, 57: PRINT USING "###.##%"; fee#
5045 IF cf = 2 THEN LOCATE 7, 1: PRINT "Apply Penalty (A)utomatic or (M)anual ?  (ESC-Menu) .. ";  ELSE LOCATE 5, 1: PRINT "Apply Penalty (A)utomatic or (M)anual ?  (ESC-Menu)";
     fl = 1: GOSUB 62210: IF cf = 1 THEN GOTO 10
     IF NOT (in$ = "A" OR in$ = "M") THEN GOTO 5045
     auto$ = in$
5050 LOCATE 8, 1: COLOR 10: PRINT STRING$(79, "-")
5060 OPEN "R", 2, "wbfee.dat", 46: CLOSE 2: KILL "wbfee.dat": GOSUB 60000: OPEN "R", 4, "wbmeter.dat", 16: FIELD 4, 12 AS MET$, 4 AS mrec$
5070 FOR x = 1 TO LOF(4) / 16: GET 4, x: IF CVS(mrec$) = 0 THEN 5090 ELSE GET 1, CVS(mrec$): GET 9, CVS(mrec$)
5080 IF CVD(PASTDUE$) > .01 THEN GOSUB 5100
5090 NEXT x: GOSUB 5700: CLOSE : GOTO 10
5100 REM display account and query for charge
5101 bb = INSTR(meter$, "/"): IF bb = 0 THEN bb = 9
5102 m$ = LEFT$(meter$, bb - 1)
5110 B = INSTR(firstname$, "  "): IF B = 0 THEN B = 21
5120 IF B = 2 THEN NME$ = lname$: GOTO 5140
5130 NME$ = LEFT$(firstname$, B) + lname$
5140 LOCATE 10, 1: COLOR 11: PRINT STRING$(53, 32): LOCATE 10, 1: PRINT "Account of ... "; NME$: LOCATE 10, 55: PRINT STRING$(24, 32): LOCATE 10, 55: PRINT "Balance of "; USING "$####.##"; CVD(curbal$) + CVD(PASTDUE$)
     IF auto$ = "A" THEN GOTO 5165
5145 LOCATE 12, 1: PRINT "Charge Penalty Fee ? (Y/N) .."; : fl = 1: GOSUB 62210: IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO 5145
     IF in$ = "N" THEN GOTO 5090
5165 IF PENFLAG = 1 THEN fee# = CVD(PASTDUE$) * (PERCENT / 100): fee# = INT((fee# * 100) + .5) / 100
5170 LSET PASTDUE$ = MKD$((CVD(PASTDUE$) + fee#))
     acctbal# = CVD(curbal$) + CVD(PASTDUE$)
    
     REM ***** ALL REVENUE INTO REVENUE SOURCE 1 FOR NOW *****
     REM LSET IRev1$ = MKD$(CVD(IRev1$) + FEE#): PUT 9, CVS(mrec$)
     REM ***** ALL REVENUE INTO REVENUE SOURCE 1 FOR NOW *****

     REM 8/31/94 wants penalty seperate
         LSET irev3$ = MKD$(CVD(irev3$) + fee#): PUT 9, CVS(mrec$)

     REM UPDATE TRANSACTION FILE
5180 GOSUB 61000: NEXRECORD = LOF(2) / 50 + 1
5190 LSET TDATE$ = LEFT$(DATE$, 6) + RIGHT$(DATE$, 2)
     LSET tcuramt$ = MKS$(fee#): LSET TAMT$ = MKD$(acctbal#)
     LSET TTYPE$ = MKI$(2): LSET TDESC$ = "Penalty Fee": LSET tmrec$ = MKD$(CVS(mrec$))
      PUT 2, NEXRECORD
5200 IF CVD(firsttr$) = 0 THEN LSET firsttr$ = MKD$(NEXRECORD): LSET lasttr$ = MKD$(NEXRECORD): PUT 1, CVS(mrec$): CLOSE 2: GOTO 5220
5210 PREV = CVD(lasttr$): LSET lasttr$ = MKD$(NEXRECORD): PUT 1, CVS(mrec$): GET 2, PREV: LSET tnexttr$ = MKS$(NEXRECORD): PUT 2, PREV: CLOSE 2
5220 REM put in temp file for printout
5230 OPEN "R", 2, "wbfee.dat", 46: FIELD 2, 8 AS FEEMETER$, 30 AS FEENAME$, 8 AS FEEAMT$
5240 LSET FEEMETER$ = m$: LSET FEENAME$ = NME$: LSET FEEAMT$ = MKD$(fee#)
5250 PUT 2, LOF(2) / 46 + 1: CLOSE 2: FEEFLAG = 1: RETURN
5700 REM open and print fee charged report
5705 IF FEEFLAG = 0 THEN RETURN
5710 LOCATE 18, 1: COLOR 10: PRINT STRING$(79, "-")
5720 LOCATE 19, 1: COLOR 11: PRINT "Get the Printer Ready for Printing, Press Any Key When Ready "; : fl = 1: GOSUB 62000
5730 OPEN "R", 2, "wbfee.dat", 46: FIELD 2, 8 AS FEEMETER$, 30 AS FEENAME$, 8 AS FEEAMT$
5740 GOSUB 5900
5750 FOR x = 1 TO LOF(2) / 46: GET 2, x
5760 IF lc >= 56 THEN LPRINT CHR$(12); : GOSUB 5900
5770 LPRINT FEEMETER$; TAB(15); FEENAME$; TAB(50); USING "$$###.##"; CVD(FEEAMT$): t# = t# + CVD(FEEAMT$): lc = lc + 1
5780 NEXT x
5790 LPRINT STRING$(65, "-"): LPRINT "Total Fees Charge "; TAB(49); USING "$$####.##"; t#: t# = 0: LPRINT CHR$(12); : RETURN
5900 LPRINT TAB(30); "Penalty Fees Report": LPRINT "Date: "; DATE$: LPRINT
5910 LPRINT " Loc  # "; TAB(15); "Customer Name "; TAB(50); "Fee Charged": LPRINT STRING$(65, "="): lc = 4: RETURN
6000 REM HAND HELD SUB-ROUTINE
     SHELL "CONNECT 1,any,com1,19200"
     SCREN$ = "3.07.00": CLS : GOSUB 63500
6010 COLOR 11: LOCATE 1, 32: PRINT "Process Meter Readings"; : LOCATE 3, 1
6020 LOCATE 4, 25: PRINT "1- Prepare PC/3000 for Monthly Readings"
     LOCATE 6, 25: PRINT "2- Retrieve Data from PC/3000 "
6021 LOCATE 8, 25: PRINT "Selection # (ESC-Menu) .. "; : fl = -1: GOSUB 62000: IF cf = 1 THEN 10
     IF VAL(in$) < 1 OR VAL(in$) > 2 THEN 6021
     ON VAL(in$) GOTO 6100, 6500
6100 CLS : PRINT "Sys. Date: "; DATE$; : COLOR 11: PRINT TAB(30); "PC/3000 Preparation"; : COLOR 10: PRINT TAB(70); "3.06.01"
     PRINT STRING$(79, "=")
     COLOR 11: LOCATE 3, 1: PRINT "STEP #1: Plug the PC/3000 I/F cable in the PC/3000": GOSUB 6490
     COLOR 11: LOCATE 4, 1: PRINT "STEP #2: Put the PC/3000 into communications mode by:"
     COLOR 11: LOCATE 5, 1: PRINT "      A: Turn on the PC/3000": GOSUB 6490
     COLOR 11: LOCATE 6, 1: PRINT "      B: Using Arrow Keys : Highlight Directory Functions and press <ENTER>": GOSUB 6490
     COLOR 11: LOCATE 7, 1: PRINT "      C: Press the L Key on the PC/3000 ": GOSUB 6490
     COLOR 11: LOCATE 10, 1: PRINT "   The PC/3000 should display SERIAL COMMUNICATIONS IN PROGRESS"
6110 LOCATE 11, 1: COLOR 11: PRINT "   -- DOES IT?? (Y/N) .. "; : fl = 1: GOSUB 62210: IF NOT (in$ = "Y" OR in$ = "N") THEN 6110
     IF in$ = "Y" THEN 6120
     LOCATE 12, 1: COLOR 11: PRINT "Turn off the hand held unit (Shift E) ": GOSUB 6490
     GOTO 6100
6120 REM FILE DOWNLOAD
6121 CLS : LOCATE 13, 1: PRINT "           System Now Compiling & Transfering Routing Information"
     GOSUB 60000
     OPEN "R", 2, "wbcusttr.dat", 94
     FIELD 2, 20 AS trlname$, 19 AS trfirstname$, 1 AS readcode$, 20 AS tradd1$, 6 AS trtime$, 4 AS trdate$, 8 AS travguse$, 8 AS trcurread$, 8 AS TRPASTREAD$
     FOR x = 1 TO LOF(1) / 256: GET 1, x
     LSET trlname$ = lname$: LSET trfirstname$ = firstname$
     LSET trtime$ = "": LSET trdate$ = ""
     LSET tradd1$ = ADD2$: LSET travguse$ = MKD$(CVD(AVGUSE$))
     LSET trcurread$ = MKD$(CVD(CURREAD$))
     LSET TRPASTREAD$ = MKD$(CVD(PASTREAD$))
     LSET readcode$ = "N"
     PUT 2, x
6125 NEXT x
     CLOSE
6130 SHELL "CONNECT 1,any,com1,19200"
     SHELL "copy wbcusttr.dat d:"
6140 SHELL "copy wbmeter.dat d:"
6150 CLS : PRINT "Sys. Date: "; DATE$; : COLOR 11: PRINT TAB(30); "PC/3000 Processing": COLOR 10: PRINT STRING$(79, "=")
     COLOR 11: LOCATE 3, 1: PRINT "All Done, Please Disconnect Cable": GOSUB 6490
     COLOR 11: LOCATE 5, 1: PRINT "The PC/3000 can now be used in the field using the directions"
              COLOR 11:    PRINT "Found in the Manual": GOSUB 6490
                  GOTO 10
6490 LOCATE 20, 1: COLOR 11: PRINT "Press any key when this step is complete"; : fl = 1: GOSUB 62000: LOCATE 20, 1: PRINT STRING$(79, 32): IF cf = 1 THEN CLOSE : GOTO 10 ELSE RETURN

6500 REM FILE UPLOAD
6510 CLS : PRINT "Sys. Date: "; DATE$; : COLOR 11: PRINT TAB(30); "PC/3000 Retrieval"; : COLOR 10: PRINT TAB(70); "3.06.02"
     PRINT STRING$(79, "=")
     COLOR 11: LOCATE 3, 1: PRINT "STEP #1: Plug the PC/3000 I/F cable in the PC/3000": GOSUB 6490
     COLOR 11: LOCATE 4, 1: PRINT "STEP #2: Put the PC/3000 into communications mode by:"
     COLOR 11: LOCATE 5, 1: PRINT "      A: Turn on the PC/3000": GOSUB 6490
     COLOR 11: LOCATE 6, 1: PRINT "      B: Using Arrow Keys : Highlight Directory Functions and press <ENTER>": GOSUB 6490
     COLOR 11: LOCATE 7, 1: PRINT "      C: Press the L Key on the PC/3000 ": GOSUB 6490
     COLOR 11: LOCATE 10, 1: PRINT "   The PC/3000 should display SERIAL COMMUNICATIONS IN PROGRESS"
6511 COLOR 11: LOCATE 11, 1: PRINT "   -- DOES IT?? (Y/N) .. "; : fl = 1
     GOSUB 62210: IF NOT (in$ = "Y" OR in$ = "N") THEN 6510
     IF in$ = "Y" THEN 6520
     COLOR 11: LOCATE 12, 1: PRINT "Turn off the hand held unit (Shift E) ": GOSUB 6490
     GOTO 6500
6520 REM
     SHELL "CONNECT 1,any,com1,19200"
     SHELL "copy d:wbcusttr.dat"
     OPEN "R", 2, "wbcusttr.dat", 94: GOSUB 60000: GOSUB 60100
     FIELD 2, 20 AS trlname$, 19 AS trfirstname$, 1 AS readcode$, 20 AS tradd1$, 6 AS trtime$, 4 AS trdate$, 8 AS travguse$, 8 AS trcurread$, 8 AS TRPASTREAD$
     FOR x = 1 TO LOF(2) / 94: GET 2, x: GET 1, x: GET 6, x
     IF readcode$ = "N" THEN 6530
     LSET PASTREAD$ = MKD$(CVD(TRPASTREAD$))
     LSET CURREAD$ = MKD$(CVD(trcurread$))
     LSET mtrtime$ = trtime$: LSET mtrdate$ = trdate$
     PUT 1, x: PUT 6, x
6530 NEXT x: CLOSE
     GOTO 10

7000 REM PRE-BILLING REPORT SUB-ROUTINE
     grandtot# = 0
     SCREN$ = "3.02.00": CLS : GOSUB 63500: LOCATE 1, 28: COLOR 11: PRINT "Print Pre-Billing Report"
7020 LOCATE 3, 1:       COLOR 11: PRINT "This Report Calculates and Prints a Report Showing What Each Customer Will "
     PRINT "Be Billed Based on the Current Billing Information Entered"
     LOCATE 5, 1: COLOR 11
     PRINT "Enter Water Test Fee Amount (ESC-MENU) .. "; : fl = -8
      GOSUB 62000: IF cf = 1 THEN 10
      testfee = VAL(in$)
7030 LOCATE 7, 1: COLOR 11
     PRINT "Enter Garbage Fee Amount (ESC-MENU) ..... "; : fl = -8
      GOSUB 62000: IF cf = 1 THEN 10
      trash = VAL(in$)
7050 LOCATE 9, 1: COLOR 11
     PRINT "Ready to Print the Report (Y/N) ... "; : fl = 1
     GOSUB 62210: IF cf = 1 THEN 10
     IF NOT (in$ = "Y" OR in$ = "N") THEN 7050
     IF in$ = "N" THEN 10
     
7100 GOSUB 60000: GOSUB 60100
     OPEN "R", 3, "wbmeter.dat", 16: FIELD 3, 12 AS MET$, 4 AS mrec$
     GOSUB 7900
     FOR x = 1 TO LOF(3) / 16: GET 3, x
     IF CVS(mrec$) < 1 OR LEFT$(MET$, 1) = " " THEN 7110
     GET 1, CVS(mrec$): GET 6, CVS(mrec$)
     IF CYCLE$ = "I" THEN 7110
     IF CVI(Rate$) < 1 AND CVD(PASTDUE$) = 0 THEN 7110
     GOSUB 7200
7110 NEXT x: LPRINT : LPRINT TAB(50); "Total Billing:  "; : LPRINT USING "$$######,.##"; grandtot#
     LPRINT CHR$(12);
     CLOSE
     GOTO 10

7200 REM calc and print
     LSET PASTDUE$ = MKD$(CVD(PASTDUE$) + CVD(curbal$)): LSET curbal$ = MKD$(0)
     PUT 1, CVS(mrec$)
     bb = INSTR(meter$, "/"): IF bb = 0 THEN bb = 9
     m$ = LEFT$(meter$, bb - 1)
     GOSUB 2400
7201 REM check rate code
     IF CVI(Rate$) = 1 THEN GOSUB 41000
     IF CVI(Rate$) = 2 THEN GOSUB 42000
     IF CVI(Rate$) = 3 THEN GOSUB 43000   'New Outside Rate Added 7-12-99
     wateramount# = INT((wateramount# * 100) + .5) / 100
     IF LEFT$(watertest$, 1) = "Y" THEN testamount# = testfee ELSE testamount# = 0
     IF LEFT$(sanitation$, 1) = "Y" THEN trashamount# = trash ELSE trashamount# = 0
7210 REM get other billings
     GOSUB 7500
     IF CVI(Rate$) < 1 THEN
        wateramount# = 0
        penaltyamount# = 0
        testamount# = 0
        trashamount# = 0
        GOTO 7250
     END IF
7220 wateramount# = wateramount# + ocharge#(1)
     testamount# = testamount# + ocharge#(2)
     penaltyamount# = penaltyamount# + ocharge#(3)
     trashamount# = trashamount# + ocharge#(4)
     wateramount# = INT((wateramount# * 100) + .5) / 100
     testamount# = INT((testamount# * 100) + .5) / 100
     penaltyamount# = INT((penaltyamount# * 100) + .5) / 100
     trashamount# = INT((trashamount# * 100) + .5) / 100
7250 calcbill# = wateramount# + testamount# + CVD(PASTDUE$) + penaltyamount# + trashamount#
     grandtot# = grandtot# + calcbill#
7300 REM print line and total
     IF lc >= 56 THEN LPRINT CHR$(12); : GOSUB 7900
     B = INSTR(firstname$, "  "): IF B = 0 THEN B = 20
     NME$ = LEFT$(firstname$, B) + lname$
     LPRINT USING "#####"; CVS(mrec$);
      LPRINT TAB(12); LEFT$(NME$, 30);
      LPRINT TAB(45); USING "########"; CALCUSAGE#;
       LPRINT TAB(55); USING "#####.##"; CVD(PASTDUE$);
       LPRINT TAB(65); USING "$#####,.##"; calcbill#;
     IF calcbill# < 0 OR CALCUSAGE# > 50000 THEN LPRINT TAB(77); "XX" ELSE LPRINT
     LPRINT USING "####.##"; wateramount#;
     LPRINT TAB(15); USING "####.##"; testamount#;
     LPRINT TAB(30); USING "####.##"; penaltyamount#;
     LPRINT TAB(45); USING "####.##"; trashamount#

     lc = lc + 3: LPRINT STRING$(79, "-")
     calcbill# = 0:  wateramount# = 0
     penaltyamount# = 0
     trashamount# = 0: testamount# = 0
     FOR jj = 1 TO 5: ocharge#(jj) = 0: NEXT jj
     RETURN
7500 REM other billings
     GET 6, CVS(mrec$)
     SOURCE = VAL(LEFT$(RACCT1$, 1)): GOSUB chksource: ocharge#(SOURCE) = ocharge#(SOURCE) + CVS(RAMT1$)
     SOURCE = VAL(LEFT$(RACCT2$, 1)): GOSUB chksource: ocharge#(SOURCE) = ocharge#(SOURCE) + CVS(RAMT2$)
     SOURCE = VAL(LEFT$(RACCT3$, 1)): GOSUB chksource: ocharge#(SOURCE) = ocharge#(SOURCE) + CVS(RAMT3$)
     SOURCE = VAL(LEFT$(OACCT1$, 1)): GOSUB chksource: ocharge#(SOURCE) = ocharge#(SOURCE) + CVS(OAMT1$)
     SOURCE = VAL(LEFT$(OACCT2$, 1)): GOSUB chksource: ocharge#(SOURCE) = ocharge#(SOURCE) + CVS(OAMT2$)
     RETURN
chksource:
     IF SOURCE < 1 OR SOURCE > 5 THEN SOURCE = 5
     RETURN
7900 LPRINT TAB(20); "Pre-Billing Report    ///    Date: "; DATE$
     LPRINT : LPRINT "Acct #"; TAB(12); "Acct Name"; TAB(45); "  Usage"; TAB(55); "Pastdue"; TAB(65); "Calc Bill"; TAB(77); "CHK"
     LPRINT "  Water"; TAB(15); "Water Test"; TAB(30); "Penalty"; TAB(45); " Trash "
     LPRINT STRING$(79, "="): lc = 4: RETURN

20000 CLOSE 4: OPEN "R", 4, "wbmeter.dat", 16: FIELD 4, 12 AS MET$, 4 AS ARREC$
 IF LOF(4) / 16 = 0 THEN m$ = smeters$:  GOTO 20010
 HI = LOF(4) / 16: LO = 1: ME = HI: GOSUB 20009: leng = LEN(meters$)
 FOR xx = 1 TO 16: GET 4, ME
 IF meters$ = LEFT$(MET$, leng) THEN NG = 1: record = CVS(ARREC$): CLOSE 4: RETURN
 IF meters$ > LEFT$(MET$, leng) THEN LO = ME ELSE HI = ME
 ME = INT((HI - LO) / 2) + LO
 NEXT xx: CLOSE 4
 CLOSE 4: NG = 0: RETURN
20009 sp = 4: SPECFLAG = 0: REM specflag indicated more than one hyphen in meter#
 B = INSTR(sp, meters$, "-")
 IF B = 0 THEN B = 8 ELSE SPECFLAG = 1
 m$ = LEFT$(meters$, 3)
IF VAL(MID$(meters$, 4, B)) < 10 THEN m$ = m$ + "    " + MID$(meters$, 4, 6): GOTO 20011
IF VAL(MID$(meters$, 4, B)) < 100 THEN m$ = m$ + "   " + MID$(meters$, 4, 6): GOTO 20011
IF VAL(MID$(meters$, 4, B)) < 1000 THEN m$ = m$ + "  " + MID$(meters$, 4, 6): GOTO 20011
IF VAL(MID$(meters$, 4, B)) < 10000 THEN m$ = m$ + " " + MID$(meters$, 4, 6): GOTO 20011
20010 m$ = m$ + MID$(meters$, 4, 6)
20011 meters$ = m$
      RETURN

41000 REM intown rate rate code=1
41010 IF CALCUSAGE# <= 1000 THEN wateramount# = 10.14: SEWER# = 0: RETURN
'41010 IF CALCUSAGE# <= 1000 THEN wateramount# = 9.66: SEWER# = 0: RETURN
41020 USAGE# = CALCUSAGE#
'41030 wateramount# = 9.66
41030 wateramount# = 10.14
41040 USAGE# = USAGE# - 1000
      'wateramount# = wateramount# + ((USAGE# / 1000) * 3.57)
      wateramount# = wateramount# + ((USAGE# / 1000) * 3.75)
      SEWER# = 0
      RETURN



42000 REM POOL rate rate code=2
42010 wateramount# = ((USAGE# / 1000) * 8.05): SEWER# = 0
42020 RETURN


43000 REM Out-Town rate rate code=3
'43010 IF CALCUSAGE# <= 1000 THEN wateramount# = 13.89: SEWER# = 0: RETURN
43010 IF CALCUSAGE# <= 1000 THEN wateramount# = 14.58: SEWER# = 0: RETURN

43020 USAGE# = CALCUSAGE#
'43030 wateramount# = 13.89
43030 wateramount# = 14.58
43040 USAGE# = USAGE# - 1000
      'wateramount# = wateramount# + ((USAGE# / 1000) * 4.83)
      wateramount# = wateramount# + ((USAGE# / 1000) * 5.07)
      SEWER# = 0
      RETURN

43100 IF USAGE# > 20000 THEN 43200
43110 wateramount# = 54.63
43120 USAGE# = USAGE# - 10000: wateramount# = wateramount# + ((USAGE# / 1000) * 3.63)
      SEWER# = 0
      RETURN

43200 wateramount# = 90.93
43210 USAGE# = USAGE# - 20000: wateramount# = wateramount# + ((USAGE# / 1000) * 3.5)
      SEWER# = 0
      RETURN
43250 RETURN


60000 CLOSE 1: OPEN "R", 1, "wbcust.dat", 256
60010 FIELD 1, 25 AS lname$, 25 AS firstname$, 30 AS ADD1$, 30 AS ADD2$, 20 AS CITY$, 2 AS ST$, 9 AS ZP$, 12 AS PHONE$, 2 AS BOOK$, 12 AS meter$, 2 AS Rate$, 8 AS watertest$, 8 AS sanitation$, 8 AS AVGUSE$, 2 AS AVGCT$, 8 AS curbal$, 8 AS PASTDUE$,  _
1 AS CYCLE$, 5 AS NOTUSED$, 8 AS CURREAD$, 8 AS PASTREAD$, 8 AS firsttr$, 8 AS lasttr$, 4 AS waterdep$
60015 CLOSE 9: OPEN "R", 9, "wbcust2.dat", 64
 FIELD 9, 8 AS IRev1$, 8 AS IRev2$, 8 AS irev3$, 8 AS irev4$, 8 AS IRev5$, 8 AS IRev6$, 8 AS irev7$, 8 AS irev8$
60030 RETURN
60100 CLOSE 6: OPEN "R", 6, "wbcust1.dat", 256
60110 FIELD 6, 4 AS RAMT1$, 4 AS RAMT2$, 4 AS RAMT3$, 25 AS RDESC1$, 25 AS RDESC2$, 25 AS RDESC3$, 9 AS RACCT1$, 9 AS RACCT2$, 9 AS RACCT3$, 4 AS OAMT1$, 4 AS OAMT2$, 25 AS ODESC1$, 25 AS ODESC2$, 9 AS OACCT1$, 9 AS OACCT2$, 2 AS sewerrate$, 2 AS  _
CUSTTYPE$, 14 AS METERNUM$, 30 AS add3$, 2 AS sprinklerrate$, 2 AS minimumS$, 10 AS NOTUSED$
60120 RETURN
60200 CLOSE 7: OPEN "R", 7, "wbconsum.dat", 336
60210 FOR xx = 0 TO 11: FIELD 7, (28 * xx) AS DUM$, 8 AS BD$(xx), 8 AS CR$(xx), 8 AS PR$(xx), 4 AS BT$(xx): NEXT xx: RETURN
61000 CLOSE 2: OPEN "R", 2, "wbartran.dat", 50
61010 FIELD 2, 8 AS TDATE$, 8 AS TAMT$, 2 AS TTYPE$, 15 AS TDESC$, 4 AS tmrec$, 4 AS tnexttr$, 4 AS tcuramt$
61020 RETURN
62000 cf = 0: COLOR 14: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
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

62061 IF z$ = CHR$(9) THEN cf = 2: RETURN
62062 IF z$ = ";" THEN cf = 2: RETURN
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
63000 OPEN "R", 3, "wbtown.dat": FIELD 3, 40 AS TOWNNAME$: GET 3, 1: TOWN$ = TOWNNAME$: CLOSE 3: RETURN
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN
fixerror:
     REM ERROR CODES
     Number = ERR
     SELECT CASE ERR
          CASE IS = 24, 25, 71
               LOCATE 23, 1: PRINT STRING$(79, 32): LOCATE 23, 1: COLOR 28
               PRINT "Error Message:"; : COLOR 15: PRINT "Device Time-out! Make Sure Printer Is ON_LINE : Press ENTER to Resume Printing";
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


