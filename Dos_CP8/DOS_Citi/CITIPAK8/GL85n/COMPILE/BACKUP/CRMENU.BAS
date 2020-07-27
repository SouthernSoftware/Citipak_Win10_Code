10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM wbmenu
30 REM main water billing system menu
40 REM Last Revision Date: 7-14-87
50 REM by Nick Picerno
60 CLS : DIM chknum(1000)
100 GOSUB 63000: REM get town in town$
110 SCREN$ = "Main": SYSDATE$ = DATE$
120 CLS : GOSUB 63500
130 LOCATE 5, 1: COLOR 11
200 PRINT TAB(30); " Check Reconciliation"
210 PRINT TAB(30); "     Master Menu      "
220 PRINT : PRINT
230 PRINT TAB(28); "1. Enter Cleared Checks"
240 PRINT
250 PRINT TAB(28); "2. Print Outstanding Checks"
260 PRINT
    PRINT TAB(28); "3. Add Old Checks to System"
310 LOCATE 17, 1
     PRINT TAB(28); "Selection No. (ESC-End) : "; : fl = -1: GOSUB 62000
320 IF cf = 1 THEN CLOSE : CLS : RUN "CPMENU"
325 IF LEN(in$) = 0 THEN 310
330 IF VAL(in$) < 1 OR VAL(in$) > 3 THEN 310
340 ON VAL(in$) GOTO 1000, 2000, 3000
1000 REM enter cleared checks
1010 SCREN$ = "07.01.00": SYSDATE$ = DATE$
1020 CLS : GOSUB 63500: COLOR 11: LOCATE 1, 30: PRINT "Check Clearing Entry"
1030 LOCATE 3, 1: COLOR 14: PRINT "*** Enter All Checks to Clear "
     l = 1: chknum = 0
1040 LOCATE 4, 1: COLOR 11: PRINT "Last Check Number Entered .... "; chknum; TAB(60); "Entry # "; l
     LOCATE 5, 1: PRINT STRING$(79, "-")
1050 LOCATE 6, 1: COLOR 11: PRINT "Enter Check Number to Clear (ESC-Done) .. "; : fl = -8:
     GOSUB 62000
     IF cf = 1 THEN 1100
     chknum = VAL(in$): chknum(l) = chknum: l = l + 1
     GOTO 1040
1100 IF l = 1 THEN CLOSE : GOTO 10
     GOSUB 1500
     GOSUB 10000
     CLS
1110 LOCATE 10, 1: PRINT "Now Clearing the Checks : Please Wait !!!"
     IF l = 1 THEN CLOSE : GOTO 10
1120 FOR xx = 1 TO LOF(2) / 64: GET 2, xx
1130 IF CVS(chknum$) = 0 THEN 1150
1140 FOR YY = 1 TO l - 1
      chknum = CVS(chknum$): chknum = FIX(chknum)
      LSET chknum$ = MKS$(chknum): PUT 2, xx
      IF CVS(chknum$) = chknum(YY) THEN
       LSET chknum$ = MKS$(0)
       chknum(YY) = 0
       : PUT 2, xx: GOTO 1150
      END IF
     NEXT YY
1150 NEXT xx
1160 CLOSE : CLS
 FOR ZZ = 1 TO l - 1
 IF chknum(ZZ) > 0 THEN PRINT "CHECK "; chknum(ZZ); " DID NOT CLEAR : COULD NOT LOCATE"
 NEXT ZZ
 PRINT "ALL DONE : PRESS <ENTER> TO CONTINUE "; : fl = 1: GOSUB 62000
 GOTO 10
1500 REM review now
     CLS : PRINT "For Your Review"
     LPRINT "Sel #  Chk #"; TAB(20); "Sel #  Chk #"; TAB(40); "Sel #  Chk #"; TAB(60); "Sel #  Chk #"
1510 FOR jj = 1 TO l - 1
     PRINT jj; chknum(jj),
   IF jj / 4 = INT(jj / 4) OR jj = 1 THEN
     IF jj = 1 THEN
      LPRINT USING "####"; jj;
       LPRINT TAB(6); USING "#####"; chknum(jj);
        LPRINT " ____ ";
         LPRINT TAB(20); USING "####"; jj + 1;
          LPRINT TAB(26); USING "#####"; chknum(jj + 1);
           LPRINT " ____";
            LPRINT TAB(40); USING "####"; jj + 2;
             LPRINT TAB(46); USING "#####"; chknum(jj + 2);
              LPRINT " ____";
         LPRINT TAB(60); USING "####"; jj + 3;
          LPRINT TAB(66); USING "#####"; chknum(jj + 3);
           LPRINT " ____"
       ELSE

      LPRINT USING "####"; jj + 1;
       LPRINT TAB(6); USING "#####"; chknum(jj + 1);
        LPRINT " ____ ";
         LPRINT TAB(20); USING "####"; jj + 2;
          LPRINT TAB(26); USING "#####"; chknum(jj + 2);
           LPRINT " ____";
            LPRINT TAB(40); USING "####"; jj + 3;
             LPRINT TAB(46); USING "#####"; chknum(jj + 3);
              LPRINT " ____";
         LPRINT TAB(60); USING "####"; jj + 4;
          LPRINT TAB(66); USING "#####"; chknum(jj + 4);
           LPRINT " ____"
      END IF
   END IF

     IF jj / 80 = INT(jj / 80) THEN GOSUB 1520: CLS : PRINT "For Your Review"
     NEXT jj
1520 LOCATE 22, 1: PRINT "Selection to Change (ESC-ALL OK) .. "; : fl = -3: GOSUB 62000
     IF cf = 1 THEN RETURN
     choic = VAL(in$)
     LOCATE 23, 1: PRINT "Enter New Check Number for Selection # "; choic; " .. "; : fl = -8: GOSUB 62000
     chknum(choic) = VAL(in$): GOTO 1500

2000 REM print outstanding checks
TOTCHKS# = 0
2010 SCREN$ = "07.02.00": SYSDATE$ = DATE$
2020 CLS : GOSUB 63500: LOCATE 1, 30: PRINT "List Outstanding Checks"
     LOCATE 3, 1: COLOR 11
     PRINT "ENTER (1)Payroll Checks or (2)A/P & Manual Checks .. "; : fl = -1: GOSUB 62000
     IF cf = 1 THEN 10
     typ = VAL(in$)
2025 LOCATE 4, 1: COLOR 11
     PRINT "Enter Month/Year to Report (mm/yy  or ESC-Menu) ... "; : fl = 5
     GOSUB 62000
     IF cf = 1 THEN 10
     IF LEN(in$) <> 5 THEN 2025
     mth$ = in$
     mth = VAL(RIGHT$(in$, 2) + LEFT$(in$, 2))
2030 GOSUB 10000
     IF LOF(2) / 64 = 0 THEN CLOSE : GOTO 10
     GOSUB 2900
2040 FOR x = 1 TO LOF(2) / 64: GET 2, x
     chknum = CVS(chknum$): chknum = FIX(chknum):
     IF CVS(chknum$) = 0 THEN 2050
     IF typ = 2 AND CVI(chksource$) <> 1 THEN 2045
     IF typ = 1 AND CVI(chksource$) = 1 THEN 2045
     GOTO 2050
2045 cdate = VAL(RIGHT$(chkdate$, 2))
     IF cdate > 0 THEN cdate = VAL(RIGHT$(chkdate$, 2) + LEFT$(chkdate$, 2)): GOTO 2046
     cdate = VAL(RIGHT$(chkdate$, 4) + LEFT$(chkdate$, 2))
2046   IF cdate <= mth THEN 2047 ELSE 2050
2047 IF lc >= 58 THEN LPRINT CHR$(12); : GOSUB 2900
     LPRINT USING "########"; CVS(chknum$);
      LPRINT TAB(15); chkdate$; TAB(30); chkdesc$; TAB(64); USING "$$######,.##"; CVS(chkamt$): lc = lc + 1: TOTCHKS# = TOTCHKS# + CVS(chkamt$)
2050 NEXT x
     LPRINT
     LPRINT "TOTAL OUTSTANDING CHECKS .... "; USING "$$######,.##"; TOTCHKS#
     LPRINT CHR$(12); : CLOSE
     GOTO 10
2900 LPRINT TAB(30); "Outstanding Check Report"
2910 LPRINT "Report Date: "; DATE$; TAB(55); "Through: "; mth$
2920 LPRINT
2930 LPRINT "Check Numb"; TAB(15); "Check Date"; TAB(30); "Paid to"; TAB(64); "Check Amount"
2940 LPRINT STRING$(79, "="): lc = 4: RETURN


3000 TOTCHKS# = 0
3010 SCREN$ = "07.03.00": SYSDATE$ = DATE$
3020 CLS : GOSUB 63500: LOCATE 1, 30: PRINT "Enter Old Checks Manually"
     LOCATE 3, 1: COLOR 11
     GOSUB scren1
3030 GOSUB 3100: GOSUB 3110: GOSUB 3120: GOSUB 3130
3040 LOCATE 10, 1: COLOR 11
     PRINT "Enter Selection to Change (ESC-Menu : <F1>-Process Entry) .. "; : fl = -1
     GOSUB 62000
     IF cf = 1 THEN CLOSE : GOTO 10
     IF cf = 11 THEN
      GOSUB 10000
      LSET chknum$ = MKS$(chknum)
      LSET chkdate$ = cdate$: LSET chkdesc$ = cdesc$: LSET chkamt$ = MKS$(amount#)
      LSET chksource$ = MKI$(2): LSET nul$ = ""
      PUT 2, LOF(2) / 64 + 1
      CLOSE 2
      GOTO 3000
     END IF
     IF VAL(in$) < 1 OR VAL(in$) > 4 THEN 3040
     ON VAL(in$) GOSUB 3100, 3110, 3120, 3130
     GOTO 3040
   
3100 LOCATE 5, 30: fl = -5: GOSUB 62000
     IF cf = 1 THEN CLOSE : GOTO 10
     chknum = VAL(in$): RETURN
3110 LOCATE 6, 30: fl = 8: GOSUB 62000
     IF cf = 1 THEN CLOSE : GOTO 10
     IF LEN(in$) <> 8 THEN 3110
     cdate$ = in$: RETURN
3120 LOCATE 7, 30: fl = 30: GOSUB 62000
     IF cf = 1 THEN CLOSE : GOTO 10
     cdesc$ = in$: RETURN
3130 LOCATE 8, 30: fl = -10: GOSUB 62000
     IF cf = 1 THEN CLOSE : GOTO 10
     amount# = VAL(in$): RETURN
    
scren1:
FOR ll = 5 TO 23: LOCATE ll, 1: PRINT STRING$(79, 32): NEXT ll
COLOR 11
LOCATE 5, 1
PRINT " 1- Check Number ........... "; TAB(60); "ESC-MENU"
PRINT " 2- Check Date (mm/dd/yy) .. "
PRINT " 3- Check Description ...... "
PRINT " 4- Check Amount ........... "
RETURN

10000 CLOSE 2: OPEN "crchk.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 64
10010 FIELD 2, 4 AS chknum$, 8 AS chkdate$, 30 AS chkdesc$, 4 AS chkamt$, 2 AS chksource$, 16 AS nul$
10020 RETURN

62000 cf = 0: in$ = "": COLOR 14: RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
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
63000 OPEN "R", 3, "wbtown.dat": FIELD 3, 40 AS TOWNNAME$: GET 3, 1: TOWN$ = TOWNNAME$: CLOSE 3: RETURN
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN

