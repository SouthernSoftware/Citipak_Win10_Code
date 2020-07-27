REM v--   *** Steve G added this : 2013-09-12  --v
CONST BAACCTFILELEN = 100 'was 128
CONST BATRANSFILELEN = 85 'same
CONST ACCOUNTLENGTH = 9  'was 12
CONST ACCOUNTNAMELEN = 25    'was 35
CONST ACCTBLANKLEN = 1   'was 16
'  Used in fle open and field statements (10000 - 17520)
'
'  Also changed DAYSBETWEEN (at end) to account for years past 2000.
'   Before, it was adding 1900 to year (ALWAYS) which puts
'   years 00 thru 13 = 1900 thru 1913 instead of 2000 thru 2013.
'  So I added an IF statement to check for years >50 for 1900s
'   Else it is assumed to be 2000s. (Should work until 2050.)
'
'  There still may be some data problems because the totals
'   ended with very long numbers indicating numbers larger
'   than the USING format could hold. (eg: %9132470841.36 or some such)
'
REM ^--   ***                                  --^

70 CLEAR
   DDATE$ = "930101": GOSUB DAYSBETWEEN
   BASEDAYS = DAYS
   ON ERROR GOTO networkerror
   DEF FNdn! (y, M, d) = y * 365 + INT((y - 1) / 4) + (M - 1) * 28 + VAL(MID$("000303060811131619212426", (M - 1) * 2 + 1, 2)) - ((M > 2) AND ((y AND NOT -4) = 0)) + d
79 GOSUB 80: GOTO 100
80 DIM ACCT%(12250), abud%(12250)
85 RETURN
100 REM main menu

4000 REM Account History Report
     totdebits# = 0: totcredits# = 0: sysdate$ = DATE$
4010 SCREN$ = "8.04.00": CLS : GOSUB 63500: COLOR 11
    GOSUB 10500
    IF LOF(14) = 0 THEN
     COLOR 12
     LOCATE 5, 1: PRINT "... Your Reporting Period Must Be Set Before This Report Can Run"
     LOCATE 6, 1: PRINT "... Proceed to the Account Period Setup (#7 on Main Menu) and"
     LOCATE 7, 1: PRINT "... Set Up the Period Dates as Instructed There"
     CLOSE
     COLOR 11
     LOCATE 9, 1: PRINT "Press any Key to Continue "; : fl = 1: GOSUB 62000
     RUN "bareport"
    END IF
    CLOSE

      LOCATE 1, 32: PRINT "ACCOUNT ACTIVITY REPORT "

    GOSUB 10500: GET 14, 1
    IF VAL(PERIOD$) = 1 THEN
     curbeg = VAL(RIGHT$(curbeg$, 2) + LEFT$(curbeg$, 2) + MID$(curbeg$, 3, 2))
     curend = VAL(RIGHT$(curend$, 2) + LEFT$(curend$, 2) + MID$(curend$, 3, 2))
     ELSE
     curbeg = VAL(RIGHT$(nexbeg$, 2) + LEFT$(nexbeg$, 2) + MID$(nexbeg$, 3, 2))
     curend = VAL(RIGHT$(nexend$, 2) + LEFT$(nexend$, 2) + MID$(nexend$, 3, 2))
    END IF

4020 COLOR 11: LOCATE 3, 1
     PRINT "Enter BEGINNING Account Number (ESC-Menu:F1=Beg) .. ";
     fl = -9: GOSUB 62000: IF cf = 1 THEN RUN "bareport"
     IF cf = 11 THEN begacct# = 0: LOCATE 3, 53: COLOR 15: PRINT "First Account": GOTO 4035
4030 IF LEN(IN$) = 7 OR LEN(IN$) = 9 THEN
     begacct# = VAL(IN$): begacct$ = IN$
     IF begacct# < 9999999 THEN begacct# = begacct# * 100
     IF LEN(IN$) = 7 THEN begacct$ = begacct$ + "  "
     LOCATE 3, 53: PRINT LEFT$(begacct$, 2); "-"; MID$(begacct$, 3, 3); "-"; MID$(begacct$, 6, 2); "-"; RIGHT$(begacct$, 2)
     ELSE GOTO 4020
     END IF
4035 COLOR 11: LOCATE 4, 1: PRINT "Enter ENDING    Account Number (ESC-Menu:F1=End) .. ";
     fl = -9: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN "bareport"
     IF cf = 11 THEN endacct# = 999999999: LOCATE 4, 53: COLOR 15: PRINT "Last Account": GOTO 4040
     IF LEN(IN$) = 7 OR LEN(IN$) = 9 THEN
      endacct# = VAL(IN$): endacct$ = IN$
      IF endacct# < 9999999 THEN endacct# = endacct# * 100
      IF LEN(IN$) = 7 THEN endacct$ = endacct$ + "  "
      LOCATE 4, 53: PRINT LEFT$(endacct$, 2); "-"; MID$(endacct$, 3, 3); "-"; MID$(endacct$, 6, 2); "-"; RIGHT$(endacct$, 2)
      ELSE GOTO 4035
      END IF
4040 COLOR 11: LOCATE 6, 1: PRINT "Enter the Beginning Date ((mmddyy) : <ESC>=Menu : F1=Beg) ... ";
     fl = -6: GOSUB 62000:
     IF cf = 1 THEN CLOSE : RUN "bareport"
     IF cf = 11 THEN bdate$ = "000000": begdate = 0: COLOR 15: begdate$ = "Beginning": LOCATE 6, 63: PRINT "Beginning": GOTO 4045
     IF LEN(IN$) <> 6 THEN 4040
     bdate$ = RIGHT$(IN$, 2) + LEFT$(IN$, 2) + MID$(IN$, 3, 2)
     begdate$ = LEFT$(IN$, 2) + "/" + MID$(IN$, 3, 2) + "/" + RIGHT$(IN$, 2)
     begdate = VAL(bdate$)
4045 COLOR 11: LOCATE 7, 1: PRINT "Enter the Ending Date ((mmddyy) : <ESC>=Menu : F1=End) ...... ";
     fl = -6: GOSUB 62000
     IF cf = 1 THEN CLOSE : RUN "bareport"
     IF cf = 11 THEN edate$ = "000000": COLOR 15: enddate$ = "End": enddate = 999999: LOCATE 7, 63: PRINT "End": GOTO 4050
     IF LEN(IN$) <> 6 THEN 4045
     edate$ = RIGHT$(IN$, 2) + LEFT$(IN$, 2) + MID$(IN$, 3, 2): enddate$ = LEFT$(IN$, 2) + "/" + MID$(IN$, 3, 2) + "/" + RIGHT$(IN$, 2)
     enddate = VAL(edate$)
IF enddate < begdate THEN
   LOCATE 10, 1: COLOR 11
   PRINT "Invalid Date Range : Press Any Key to Continue ";
   fl = 1: GOSUB 62000
   FOR ll = 6 TO 11: LOCATE ll, 1: PRINT STRING$(79, 32): NEXT ll
   GOTO 4040
END IF
4050 CLOSE : GOSUB 11000: GOSUB 12000
     OPEN "baacctdx.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 12: FIELD 4, 8 AS ACT$, 4 AS rec$
4080 COLOR 11: LOCATE 10, 1: PRINT "Output Specification : Output to (S)creen or (P)rinter (S,P or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
4090 IF cf = 1 THEN 4000
4100 IF NOT (IN$ = "S" OR IN$ = "P") THEN 4080
4110 devspec$ = IN$: LOCATE 10, 78: PRINT devspec$
4120 COLOR 11: LOCATE 11, 1: PRINT STRING$(79, "="): LOCATE 13, 23: PRINT "Now Processing : Please be Patient!": IF devspec$ = "S" THEN DEV$ = "scrn:" ELSE DEV$ = "lpt1:"
4130 OPEN "O", 10, DEV$: WIDTH #10, 255
     IF devspec$ = "S" THEN CLS : GOSUB 4900 ELSE GOSUB 4900
4140 FOR jk = 1 TO LOF(4) / 12: GET 4, jk: IF CVS(rec$) = 0 THEN 4500 ELSE GET 2, CVS(rec$)
     MACCT$ = LEFT$(MASTACCT$, 2) + MID$(MASTACCT$, 4, 3) + MID$(MASTACCT$, 8, 2) + MID$(MASTACCT$, 11, 2)
     MACCT# = VAL(MACCT$)
     IF MACCT# < 9999999 THEN MACCT# = MACCT# * 100
     IF MACCT# >= begacct# AND MACCT# <= endacct# THEN 4150 ELSE 4500
4150 COLOR 11
     GOSUB 4950
     BEGBALANCE# = CVD(begbal$)
     firstline = 0
     runbalance# = CVD(begbal$)
     
4170 PRINT #10, "Account Number : "; MASTACCT$; TAB(34); ACCTNAME$: lc = lc + 1
     FOR tr = 1 TO count
      trrecord = abud%(tr)
4180 IF count = 0 THEN
       PRINT #10, "******** Balance Forward ...................  "; TAB(60); USING "$$##########,.##"; BEGBALANCE#
       GOTO 4500
     END IF
4190 IF trrecord <= 0 THEN
        trrecord = trrecord - 32767: trrecord = trrecord * -1
     END IF
     GET 3, trrecord
     TRDATE$ = RIGHT$(trandate$, 2) + LEFT$(trandate$, 2) + MID$(trandate$, 4, 2)
     trandate = VAL(TRDATE$)
     IF VAL(TRDATE$) < begdate THEN
       IF accttype$ = "A" OR accttype$ = "E" THEN BEGBALANCE# = BEGBALANCE# + CVD(dramt$) - CVD(cramt$): runbalance# = runbalance# + (CVD(dramt$) - CVD(cramt$)):    runbalance# = INT((runbalance# * 100) + .5) / 100: GOTO 4230
       IF accttype$ = "L" OR accttype$ = "R" THEN BEGBALANCE# = BEGBALANCE# + CVD(cramt$) - CVD(dramt$): runbalance# = runbalance# + (CVD(cramt$) - CVD(dramt$)):      runbalance# = INT((runbalance# * 100) + .5) / 100: GOTO 4230
     END IF

     IF VAL(TRDATE$) > enddate THEN 4231
     IF accttype$ = "A" OR accttype$ = "E" THEN runbalance# = runbalance# + (CVD(dramt$) - CVD(cramt$)):     runbalance# = INT((runbalance# * 100) + .5) / 100
     IF accttype$ = "L" OR accttype$ = "R" THEN runbalance# = runbalance# + (CVD(cramt$) - CVD(dramt$)):      runbalance# = INT((runbalance# * 100) + .5) / 100


4200 IF devspec$ = "P" AND lc >= 48 THEN
      PRINT #10, CHR$(12); : GOSUB 4900
      PRINT #10, "Account Number : "; MASTACCT$; TAB(34); ACCTNAME$: lc = lc + 1
     END IF
4210 IF devspec$ = "S" AND lc >= 14 THEN
      COLOR 13: LOCATE 23, 1
      PRINT "Press Any Key for Next Page (ESC-Exit)";
      fl = 1: GOSUB 62000: CLS : LOCATE 1, 1
      IF cf = 1 THEN CLOSE : RUN "bareport"
      GOSUB 4900: PRINT #10, "Account Number : "; MASTACCT$; TAB(34); ACCTNAME$: lc = lc + 1
     END IF
4220 COLOR 11:
     IF firstline = 0 THEN
      
      COLOR 10: PRINT #10, "******** Balance Forward ...................  "; TAB(60); USING "$$##########,.##"; BEGBALANCE#
      firstline = 1: lc = lc + 1
     END IF
4225 COLOR 11: PRINT #10, trandate$; TAB(12); LEFT$(TRANDESC$, 16); TAB(30); DETAIL$; TAB(40); USING "$##########.##"; CVD(dramt$); : PRINT #10, TAB(54); USING "$##########.##"; CVD(cramt$); : PRINT #10, TAB(70); SOURCE$: lc = lc + 1
     totdebits# = totdebits# + CVD(dramt$): totcredits# = totcredits# + CVD(cramt$)
4230 NEXT tr
4231     IF firstline = 0 THEN
       COLOR 10
      PRINT #10, "******** Balance Forward ...................  "; TAB(60); USING "$$##########,.##"; BEGBALANCE#
      firstline = 1: lc = lc + 1
     END IF
     IF runbalance# < 0 THEN COLOR 12 ELSE COLOR 10
      PRINT #10, "********* Ending Balance ....................  "; TAB(60); USING "$$##########,.##"; runbalance#: lc = lc + 1
REM     LSET ACCTBAL$ = MKD$(RUNBALANCE#): PUT 2, CVS(REC$)
     IF accttype$ = "A" OR accttype$ = "L" THEN 4500
     IF accttype$ = "E" THEN PRINT #10, "Encumbered Amt: "; USING "$$##########,.##"; CVD(encumb$): lc = lc + 1
                  
    GOSUB 10500: GET 14, 1
    IF VAL(PERIOD$) = 1 THEN
    BUDGET# = CVD(BUDAMT$)
     ELSE
    GOSUB 13000
    IF LOF(7) = 0 THEN BUDGET#(count) = 0: CLOSE 7: GOTO 4235
    GET 7, CVS(rec$)
     IF CVD(ESTAMT$) > 99999999 THEN LSET ESTAMT$ = MKD$(0): PUT 7, CVS(rec$)
     IF CVD(REQAMT$) > 99999999 THEN LSET REQAMT$ = MKD$(0): PUT 7, CVS(rec$)
     IF CVD(RECAMT$) > 99999999 THEN LSET RECAMT$ = MKD$(0): PUT 7, CVS(rec$)
     IF CVD(APPAMT$) > 99999999 THEN LSET APPAMT$ = MKD$(0): PUT 7, CVS(rec$)
    GET 7, CVS(rec$)
    BUDGET# = CVD(APPAMT$)
    CLOSE 7
    END IF
  
  
4235 COLOR 15: PRINT #10, "Budgeted Amt: "; USING "$$##########,.##"; BUDGET#;


     IF accttype$ = "R" THEN PRINT #10, TAB(35); "Expected Revenue to Receive "; USING "$$#########,.##"; BUDGET# - runbalance#
     IF accttype$ = "E" THEN PRINT #10, TAB(35); "   Amount Remaining to Spend "; USING "$$#########,.##"; BUDGET# - runbalance# - CVD(encumb$)
     COLOR 11: PRINT STRING$(79, "-"): lc = lc + 1
4500 IF devspec$ = "P" AND lc >= 48 THEN
      PRINT #10, CHR$(12); : GOSUB 4900
      PRINT #10, "Account Number : "; MASTACCT$; TAB(34); ACCTNAME$: lc = lc + 1
     END IF
     IF devspec$ = "S" AND lc >= 16 THEN
      COLOR 13: LOCATE 23, 1
      PRINT "Press Any Key for Next Page (ESC-END) .. ";
      fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : RUN "BAREPORT"
      CLS : LOCATE 1, 1
      GOSUB 4900: PRINT #10, "Account Number : "; MASTACCT$; TAB(34); ACCTNAME$: lc = lc + 1
     END IF

      NEXT jk
4505 IF devspec$ = "P" THEN PRINT #10, STRING$(79, "=") ELSE 4510
     PRINT #10, "Total Debits ...... "; USING "$$##########,.##"; totdebits#
     PRINT #10, "Total Credits ..... "; USING "$$##########,.##"; totcredits#
     PRINT #10, CHR$(12); : CLOSE : RUN "bareport"
4510 IF devspec$ = "S" THEN COLOR 13:
     LOCATE 21, 1: PRINT "Total Debits ...... "; USING "$$##########.##"; totdebits#
     LOCATE 22, 1: PRINT "Total Credits ..... "; USING "$$##########.##"; totcredits#
     LOCATE 23, 1: PRINT STRING$(79, 32)
     LOCATE 23, 1: COLOR 11
     PRINT "Press Any Key to Continue "; : fl = 1: GOSUB 62000: CLOSE : RUN "bareport"
4900 COLOR 11: PRINT #10, TAB(24); "Detailed Account History Report"
4910 PRINT #10, "   Report Date: "; DATE$
     PRINT #10, "Beginning Date: "; begdate$; "    thru    Ending Date: "; enddate$
     PRINT #10, ""
4920 PRINT #10, "Tr. Date"; TAB(12); "Description"; TAB(30); "Detail"; TAB(40); "Dr.-Amount"; TAB(55); "Cr.-Amount"; TAB(70); "Source": PRINT #10, STRING$(79, "-"): lc = 5: RETURN
     REM get transactions and store sort and return
4950 tr = CVS(FIRST$)
     IF tr = 0 THEN count = 0: RETURN
     count = 0
4951 IF tr = 0 THEN LLL = CSRLIN: GOSUB 26000: LOCATE LLL, 1: RETURN
     GET 3, tr
       REM catch bad dates here  and fix them
       IF MID$(trandate$, 7, 1) = " " THEN
         LSET trandate$ = LEFT$(trandate$, 2) + "/" + MID$(trandate$, 3, 2) + "/" + MID$(trandate$, 5, 2)
         PUT 3, tr
       END IF

   credit# = CVD(cramt$): debit# = CVD(dramt$)
   credit# = INT((credit# * 100) + .5) / 100
   debit# = INT((debit# * 100) + .5) / 100
   IF credit# = 0 AND debit# = 0 THEN tr = CVS(nexttr$): GOTO 4951

   DDATE$ = RIGHT$(trandate$, 2) + LEFT$(trandate$, 2) + MID$(trandate$, 4, 2)

   date = VAL(DDATE$): TRDATE = date
   REM *** NEW PROGRAMMING
   
   IF accttype$ = "R" OR accttype$ = "E" THEN
    IF TRDATE < curbeg OR TRDATE > curend THEN tr = CVS(nexttr$): GOTO 4951
   END IF
   REM *** TO HERE

4951.1 count = count + 1

   GOSUB DAYSBETWEEN
   ACCT%(count) = DAYS - BASEDAYS
   IF tr >= 32767 THEN tr = 32767 - tr
    abud%(count) = tr
   tr = CVS(nexttr$)
   GOTO 4951

10000 CLOSE 1
     OPEN "bafund.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 64
10010 FIELD 1, 30 AS fundname$, 2 AS fundacct$
10020 RETURN
10500 CLOSE 14: OPEN "R", 14, "bayear.dat", 64
      FIELD 14, 6 AS curbeg$, 6 AS curend$, 6 AS nexbeg$, 6 AS nexend$, 1 AS PERIOD$, 39 AS blk$
      RETURN
11000 CLOSE 2
      OPEN "baacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = BAACCTFILELEN
11010 FIELD 2, ACCOUNTLENGTH AS MASTACCT$, ACCOUNTNAMELEN AS ACCTNAME$, 1 AS accttype$, 8 AS ACCTBAL$, 8 AS mtdact$, 8 AS ytdact$, 8 AS BUDAMT$, 4 AS FIRST$, 4 AS LAST$, 8 AS begbal$, 8 AS recalcbal$, 8 AS encumb$, ACCTBLANKLEN AS blank$
11020 RETURN
12000 CLOSE 3
      OPEN "batrans.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #3 LEN = BATRANSFILELEN
12010 FIELD 3, ACCOUNTLENGTH AS TRANACCT$, 8 AS trandate$, 20 AS TRANDESC$, 8 AS cramt$, 8 AS dramt$, 8 AS DETAIL$, 8 AS SOURCE$, 4 AS nexttr$
12020 RETURN
13000 CLOSE 7
   OPEN "babudprp.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #7 LEN = 64
   FIELD 7, 8 AS ESTAMT$, 8 AS REQAMT$, 8 AS RECAMT$, 8 AS APPAMT$, 8 AS lastyractual$, 8 AS curyrytd$, 16 AS notused$
   RETURN

17500 OPEN "baacctdx.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 12: FIELD 4, 8 AS ACT$, 4 AS rec$
     IF ACCT# < 9999999 THEN ACCT# = ACCT# * 100
17520 IF LOF(4) / 12 = 0 THEN NG = 0: CLOSE 4: RETURN
17540 HI = LOF(4) / 12: LO = 1: ME = HI
17560 FOR XX = 1 TO 16: GET 4, ME
17580 IF ACCT# = CVD(ACT$) THEN NG = 1: RECORD = CVS(rec$): CLOSE 4: RETURN
17600 IF ACCT# > CVD(ACT$) THEN LO = ME ELSE HI = ME
1762 ME = INT((HI - LO) / 2) + LO
1764 NEXT XX: CLOSE 4: NG = 0: RETURN
26000 REM sort
      LOCATE 23, 1: COLOR 18
      PRINT "** NOW SORTING BY TRANSACTION DATE : PLEASE WAIT **"
      COLOR 11
26010 REM count=counter
26020 M = count
26030 M = INT(M / 2)
26040 IF M = 0 THEN 26190
26050 FOR ST = 1 TO M
26060 I = ST
26070 j = ST + M
26080 SW = 0
26090 IF ACCT%(I) <= ACCT%(j) THEN 26120
26100 SW = 1
26110 SWAP ACCT%(I), ACCT%(j): SWAP abud%(I), abud%(j)
26120 I = j
26130 j = j + M
26140 IF j <= count THEN 26090
26150 IF SW = 0 THEN 26170
26160 GOTO 26060
26170 NEXT ST
26180 GOTO 26030
26190 LOCATE 23, 1: PRINT STRING$(79, 32): RETURN
62000 COLOR 14: cf = 0: IN$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
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
62110 IN$ = LEFT$(IN$, LEN(IN$) - 1)
62120 ZL = ZL - 1: GOTO 62030
62130 IF ABS(fl) = ZL THEN 62030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 62180
62140 IF z$ = "." AND ZD = 0 THEN ZD = 1: GOTO 62180
62150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF z$ < "0" OR z$ > "9" THEN 62030
62180 PRINT z$; : IN$ = IN$ + z$: ZL = ZL + 1
62190 IF ABS(fl) = 1 THEN 62050 ELSE 62030
62200 LOCATE RO, PO: PRINT IN$: RETURN
62210 GOSUB 62000: IF LEN(IN$) THEN IN$ = CHR$(ASC(IN$) AND 223)
62220 RETURN
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; sysdate$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN
networkerror:
     REM ERROR CODES
     Number = ERR
     SELECT CASE ERR
          CASE IS = 70, 55
              LOCATE 23, 1: PRINT STRING$(79, 32)
              LOCATE 23, 1: COLOR 28: PRINT "Network Message:"; : COLOR 15
              PRINT "This RECORD is Busy : PLEASE Wait ...": RESUME
              

          CASE IS = 24, 25
               LOCATE 23, 1: PRINT STRING$(79, 32): LOCATE 23, 1: COLOR 28
               PRINT "Error Message:"; : COLOR 15
               PRINT "Device Time-out! Make Sure Printer Is ON-LINE: Press ENTER to Resume Printing";
               fl = 1: GOSUB 62000: LOCATE 23, 1: PRINT STRING$(79, 32)
               RESUME
          CASE IS = 27
               LOCATE 23, 1: PRINT STRING$(79, 32): LOCATE 23, 1: COLOR 28
               PRINT "Error Message:"; : COLOR 15:
               PRINT "Printer Is OUT OF PAPER! : Press ENTER to Resume Printing";
               fl = 1: GOSUB 62000: LOCATE 23, 1: PRINT STRING$(79, 32)
               RESUME
          CASE ELSE
               CLS
               LOCATE 15, 1: COLOR 28: PRINT "Error Message: "; : COLOR 15: PRINT "Error "; Number; " Has Occured in Line "; ERL
               PRINT "Write Down Message and Hit <ESC> Key to exit."
               fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 70 ELSE GOSUB 62000
     END SELECT
     END

DAYSBETWEEN:
yr = VAL(LEFT$(DDATE$, 2))
day = VAL(RIGHT$(DDATE$, 2))
mth = VAL(MID$(DDATE$, 3, 2))
IF yr >= 50 THEN
  yr = yr + 1900
ELSE
  yr = yr + 2000
END IF

IF day = 0 OR mth = 0 THEN DAYS = 0: RETURN
DAYS = FNdn!(yr, mth, day)
RETURN

