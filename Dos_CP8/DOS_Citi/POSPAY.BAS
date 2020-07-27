10 REM (C)1989 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
   CLEAR
11 DIM tc#(20), td#(20), tn$(20), ta$(20), trec(20), ta#(20), tot#(100), translink!(1500)
   ON ERROR GOTO networkerror
110 scren$ = "4.00.00": sysdate$ = LEFT$(DATE$, 6) + RIGHT$(DATE$, 2)
    
120 CLS : GOSUB 63500
    REM INITIALIZE OFTEN USED VARIABLE
    changedue# = 0: amtreceived# = 0: amount# = 0
    IF OPSET = 1 THEN 130 ELSE GOSUB SETOPERATOR
    IF setset = 1 THEN 130 ELSE GOSUB setenv
130 LOCATE 5, 1: COLOR 11
200 PRINT TAB(30); "Daily Cash Management"
210 PRINT TAB(30); "      Payments"
220 PRINT
230 PRINT TAB(25); "1. Enter Payments"
240 PRINT
270 PRINT TAB(25); "2. Print Payment Edit Report"
280 PRINT
290 PRINT TAB(25); "3. Post Payments to Accounts"
300 PRINT
    PRINT TAB(25); "4. Set Message Line"
    PRINT
    PRINT TAB(25); "5. Operator Maintenance"
    PRINT
    PRINT TAB(25); "6. Setup/Change Misc Codes"
    PRINT
    PRINT TAB(25); "7. Set Envelope Number"
        COLOR 15
    LOCATE 3, 1: PRINT "Operator: "; operatorname$; TAB(55); "Operator ID# "; OPERATORNUMBER
    LOCATE 4, 57: PRINT "Envelope # "; ENVELOPE
310 LOCATE 23, 1: COLOR 11: PRINT TAB(23); "Selection No. (ESC-Master Menu) : "; : fl = -1: GOSUB 62000
320 IF cf = 1 THEN CLOSE : RUN "cpmenu"
325 IF LEN(in$) = 0 THEN 310
330 IF VAL(in$) < 1 OR VAL(in$) > 7 THEN 310
340 ON VAL(in$) GOTO 1000, 3000, 4000, 5000, 7000, 8000, 9000
1000 REM
     FOR JJ = 1 TO 6: apply(JJ) = 0: NEXT JJ
1010 scren$ = "4.01.00": CLS : GOSUB 63500
1020 LOCATE 1, 34: COLOR 11: PRINT "Payment Entry"; : LOCATE 3, 1
1022 LOCATE 3, 1: COLOR 11
     PRINT "ENTER:  1-Utility : 2-Tax : 3-A/R : 4-Misc.  Payment Please (ESC-Menu) .. "; : fl = -1
     GOSUB 62000: IF cf = 1 THEN 120
     IF VAL(in$) < 1 OR VAL(in$) > 4 THEN 1022
     source = VAL(in$)
     IF source = 2 THEN GOTO TAXENTRY
     IF source = 3 THEN GOTO arentry
     IF source = 4 THEN 1800
     
1025 DEPFLAG = 0
1030 FOR ii = 3 TO 20: LOCATE ii, 1: PRINT STRING$(79, 32): NEXT ii
     COLOR 11
     LOCATE 1, 30: PRINT STRING$(20, 32)
     LOCATE 1, 30: PRINT "Utility Payment Entry"
     LOCATE 3, 1

     COLOR 11: PRINT "Enter Account Number Please (ESC-Menu : TAB - Name Search) .. "; : fl = -5
     GOSUB 62000
     IF cf = 1 THEN 1000
     IF cf = 2 THEN GOSUB namesearch: IF ggo = 1 THEN GOTO 1080 ELSE CLOSE : GOTO 1000
1040 record = VAL(in$)
1080 CLOSE 1: GOSUB 60000
     IF record < 1 OR record > LOF(1) / 256 THEN
       LOCATE 5, 1: PRINT "Invalid Account Number "
       PRINT CHR$(7);
       FOR XX = 1 TO 250: NEXT XX
       CLOSE : GOTO 1000
     END IF
     GET 1, record: acctrecord = record: GET 8, record
     IF VAL(LEFT$(meter$, 2)) = 0 THEN
      LOCATE 23, 1: COLOR 11: PRINT "Account Deleted : Press <ENTER> to Continue "; : fl = 1
      GOSUB 62000: CLOSE : GOTO 1000
     END IF
1090 LOCATE 3, 1: PRINT STRING$(79, 32)
     B = INSTR(meter$, "/"): IF B = 0 THEN m$ = meter$: GOTO 1095
     m$ = LEFT$(meter$, B - 1)
1095 COLOR 15: LOCATE 3, 1: PRINT "  Account Number: "; USING "#####"; record
     LSET IREVENUE1$ = MKD$((INT((CVD(IREVENUE1$) * 100) + .5) / 100))
     LSET IREVENUE2$ = MKD$((INT((CVD(IREVENUE2$) * 100) + .5) / 100))
     LSET IREVENUE3$ = MKD$((INT((CVD(IREVENUE3$) * 100) + .5) / 100))
     LSET irevenue4$ = MKD$((INT((CVD(irevenue4$) * 100) + .5) / 100))
     LSET irevenue5$ = MKD$((INT((CVD(irevenue5$) * 100) + .5) / 100))
     LSET IREVENUE6$ = MKD$((INT((CVD(IREVENUE6$) * 100) + .5) / 100))
     LSET IREVENUE7$ = MKD$((INT((CVD(IREVENUE7$) * 100) + .5) / 100))
     LSET IREVENUE8$ = MKD$((INT((CVD(IREVENUE8$) * 100) + .5) / 100))
     PUT 8, record
     LOCATE 4, 1: PRINT "Account Location: "; m$
     LOCATE 5, 1
1100 COLOR 11: PRINT "Customer Account of ..."; TAB(45); "Balance = "; USING "####.##"; CVD(curbal$) + CVD(pastdue$)
     owed# = CVD(curbal$) + CVD(pastdue$): owed# = INT((owed# * 100) + .5) / 100
     acctbalance# = owed#
1110 COLOR 14: IF LEFT$(firstname$, 1) = " " THEN nme$ = lname$: GOTO 1135
1120 B = INSTR(firstname$, "  "): IF B = 0 THEN B = 20
     nme$ = LEFT$(firstname$, B) + lname$
1135 PRINT TAB(5); LEFT$(nme$, 39); TAB(45); "1- Water ......"; USING "$####,#.##"; CVD(IREVENUE1$)
1140 PRINT TAB(5); add1$; TAB(45); "2- Sewer ......"; USING "$####,#.##"; CVD(IREVENUE2$)
1150 PRINT TAB(5); add2$; TAB(45); "3- Penalty....."; USING "$####,#.##"; CVD(IREVENUE3$)
1160 PRINT TAB(5); city$; " "; st$; " "; zp$; TAB(45); "4- Trash ......"; USING "$####,#.##"; CVD(irevenue4$)
     PRINT TAB(45); "5- Misc ......."; USING "$####,#.##"; CVD(irevenue5$)
     TOTREV# = CVD(IREVENUE1$) + CVD(IREVENUE2$) + CVD(IREVENUE3$) + CVD(irevenue4$) + CVD(irevenue5$) + CVD(IREVENUE6$) + CVD(IREVENUE7$) + CVD(IREVENUE8$)
     TOTREV# = INT((TOTREV# * 100) + .5) / 100
1170 LOCATE 12, 1: COLOR 11: PRINT STRING$(44, "-")
1180 LOCATE 13, 1: COLOR 11: PRINT "Correct Account (Y/N) or ";
     LOCATE 14, 1: COLOR 10: PRINT "(PRESS <TAB> to edit Amounts : <F1>-History) .. "; : fl = 1: GOSUB 62210:
     IF cf = 2 THEN GOSUB fixamts: GOTO 1090
     IF cf = 11 THEN
        GOSUB history
        FOR LL = 3 TO 23
        LOCATE LL, 1: PRINT STRING$(79, 32)
        NEXT LL
        scren$ = "4.01.00": CLS : GOSUB 63500
        LOCATE 1, 34: COLOR 11: PRINT "Payment Entry"; : LOCATE 3, 1
        GOTO 1090
     END IF
     IF NOT (in$ = "Y" OR in$ = "N") THEN 1180
1190 IF in$ = "Y" THEN 1200 ELSE CLOSE : GOTO 1000
1200 IF owed# = TOTREV# THEN 1205
     LOCATE 16, 1: PRINT "INDIVIDUAL REVENUES DO NOT EQUAL THE BALANCE"
     LOCATE 17, 1: COLOR 10
     PRINT "THE DIFFERENCE IS "; USING "$$#####,.##"; owed# - TOTREV#
     LOCATE 18, 1: COLOR 10
     PRINT "YOU SHOULD CORRECT THE REVENUES TO EQUAL THE BALANCE TO INSURE PROPER"
     PRINT "POSTING TO THE GENERAL LEDGER AND FOR A REVENUE REPORT THAT IS IN BALANCE"
1201 LOCATE 21, 1: COLOR 11
     PRINT "PRESS ANY KEY TO CONTINUE "; : fl = 1
     GOSUB 62000
     FOR JJ = 12 TO 23: LOCATE JJ, 1: PRINT STRING$(79, 32): NEXT JJ
     GOTO 1170
1205 LOCATE 13, 1
      PRINT STRING$(44, 32)
     LOCATE 14, 1
      PRINT STRING$(50, 32)
1210 IF CVS(waterdep$) > 0 THEN COLOR 10: LOCATE 23, 1: PRINT "Deposit on File in the Amount of :"; USING "###.##"; CVS(waterdep$);
1220 GOSUB 1900: LOCATE 13, 1: COLOR 11: PRINT "Enter Payment Code (See Below 1-6) .. "; : fl = -1
     GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 1000 ELSE PCODE = VAL(in$)
     IF PCODE < 1 OR PCODE > 6 THEN 1220
1222 GOSUB 1950: COLOR 11: LOCATE 13, 55: PRINT dd$
1250 IF PCODE = 3 THEN 1260
     LOCATE 14, 1: COLOR 11: PRINT "Enter Amount ... "; : PRINT "("; USING "$####.##"; owed#; : PRINT ") ......... "; : fl = -8: GOSUB 62000
     IF LEN(in$) = 0 THEN amount# = owed#: LOCATE 14, 39: PRINT USING "$####.##"; amount#:  ELSE amount# = VAL(in$): LOCATE 14, 39: PRINT USING "$####.##"; amount#
     GOTO 1261
1260 amount# = CVS(waterdep$)
1261 LOCATE 15, 1: COLOR 11: PRINT "Enter Description (Ck #) ............ "; : fl = 15: GOSUB 62000: des$ = in$
     IF PCODE = 1 OR PCODE = 2 OR PCODE = 5 OR PCODE = 4 THEN 1265 ELSE 1270
1265 LOCATE 16, 1: COLOR 11: PRINT "Enter Actual Amount of Cash Rec'd ... "; : fl = -9: GOSUB 62000
     LOCATE 16, 39: PRINT USING "$####.##"; VAL(in$): amtreceived# = VAL(in$)
     LOCATE 16, 60: COLOR 10: PRINT "Change= "; : COLOR 15: PRINT USING "$$#####.##"; VAL(in$) - amount#
         changedue# = amtreceived# - amount#
1270 LOCATE 17, 1: COLOR 11: PRINT "Is the Above Correct (Y/N) .......... "; : fl = 1: GOSUB 62210: IF NOT (in$ = "Y" OR in$ = "N") THEN 1270 ELSE LOCATE 17, 39: PRINT in$
1280 IF in$ = "N" THEN FOR J = 13 TO 18: LOCATE J, 1: PRINT STRING$(79, 32): NEXT J: GOTO 1200
1285 IF amount# = owed# THEN apply(1) = CVD(IREVENUE1$): apply(2) = CVD(IREVENUE2$): apply(3) = CVD(IREVENUE3$): apply(4) = CVD(irevenue4$): apply(5) = CVD(irevenue5$): apply(6) = CVD(IREVENUE6$): GOTO 1295
     left# = amount#
     IF PCODE = 4 OR PCODE = 5 THEN 1295
1286 FOR J = 12 TO 23: LOCATE J, 1: PRINT STRING$(79, 32): NEXT J
     LOCATE 12, 25: COLOR 11: PRINT "Please Break Down the Payment"
     COLOR 10: LOCATE 13, 1: PRINT "Amount Paid by Customer: "; USING "$####.##"; amount#
     COLOR 11: LOCATE 14, 1: PRINT "1- Amount to Apply : Water ........ "
     LOCATE 15, 1: PRINT "2- Amount to Apply : Sewer ........ "
     LOCATE 16, 1: PRINT "3- Amount to Apply : Penalty ...... "
     LOCATE 17, 1: PRINT "4- Amount to Apply : Trash ........"
     LOCATE 18, 1: PRINT "5- Amount to Apply : Misc. ........ "
1288 GOSUB A1
     IF amount# - left# = 0 THEN LOCATE 21, 1: COLOR 13: PRINT "Amount has been fully distributed " ELSE LOCATE 21, 1: PRINT STRING$(65, 32)
     LOCATE 20, 1: COLOR 11: PRINT "Selection # (ESC-When Done) ... "; : fl = -1: GOSUB 62000
     IF cf = 1 THEN 1291
     LOCATE 15, 50: PRINT STRING$(28, 32)
     IF VAL(in$) < 1 OR VAL(in$) > 5 THEN 1288
     ON VAL(in$) GOSUB A2, a3, a4, a5, a6, a7: GOTO 1288
A1:
left# = 0: FOR JJ = 1 TO 6: left# = left# + apply(JJ): NEXT JJ: left# = INT((left# * 100) + .5) / 100
LOCATE 14, 50: COLOR 10: PRINT "Amount Left to Dist. "; USING "$####.##"; amount# - left#: RETURN

A2: LOCATE 14, 36: fl = -9: GOSUB 62000: apply(1) = VAL(in$): LSET trev1$ = MKS$(VAL(in$)): LOCATE 14, 36: PRINT USING "$####.##"; VAL(in$): RETURN

a3: LOCATE 15, 36: fl = -9: GOSUB 62000: apply(2) = VAL(in$): LSET trev2$ = MKS$(VAL(in$)): LOCATE 15, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a4: LOCATE 16, 36: fl = -9: GOSUB 62000: apply(3) = VAL(in$): LSET TREV3$ = MKS$(VAL(in$)): LOCATE 16, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a5: LOCATE 17, 36: fl = -9: GOSUB 62000: apply(4) = VAL(in$): LSET trev4$ = MKS$(VAL(in$)): LOCATE 17, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a6: LOCATE 18, 36: fl = -9: GOSUB 62000: apply(5) = VAL(in$): LSET trev5$ = MKS$(VAL(in$)): LOCATE 18, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a7: LOCATE 19, 36: fl = -9: GOSUB 62000: apply(6) = VAL(in$): LSET trev6$ = MKS$(VAL(in$)): LOCATE 19, 36: PRINT USING "$####.##"; VAL(in$): RETURN
1291 REM
1292 IF left# = amount# THEN 1295 ELSE LOCATE 15, 50: PRINT CHR$(7); : COLOR 13: PRINT "Not Fully Distributed Yet!": GOTO 1288
1295 CLOSE 2: OPEN OPERATORFILE$ FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 128
     FIELD 2, 2 AS psource$, 8 AS pmeter$, 4 AS prec$, 8 AS pamt$, 15 AS pdesc$, 2 AS ptype$, 30 AS pname$, 4 AS TFDIST$, 4 AS TLDIST$, 4 AS trev1$, 4 AS trev2$, 4 AS TREV3$, 4 AS trev4$, 4 AS trev5$, 4 AS trev6$, 8 AS PDATE$
1300 LSET psource$ = MKI$(1)
     LSET PDATE$ = sysdate$
     LSET pmeter$ = STR$(acctrecord)
     LSET prec$ = MKS$(record)
     LSET pamt$ = MKD$(amount#)
     LSET pdesc$ = "ENV #" + STR$(ENVELOPE)
     LSET ptype$ = MKI$(PCODE)
1320 LSET pname$ = nme$
     LSET trev1$ = MKS$(apply(1)): LSET trev2$ = MKS$(apply(2)): LSET TREV3$ = MKS$(apply(3))
     LSET trev4$ = MKS$(apply(4)): LSET trev5$ = MKS$(apply(5)): LSET trev6$ = MKS$(apply(6))
1330 nexrecord = LOF(2) / 128 + 1:
      CLOSE 8: OPEN "MESSAGE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #8 LEN = 40
       FIELD 8, 40 AS MESSAGE$: IF LOF(8) / 40 = 0 THEN GOTO RECPT
       GET 8, 1
FOR LL = 12 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
COLOR 11
LOCATE 15, 1: PRINT "Amount Paid on Account ....... "; USING "#####,#.##"; amount#
              PRINT "Cash Received from Customer .. "; USING "#####,#.##"; amtreceived#
              PRINT "                               -------------"
              COLOR 15
              PRINT "CHANGE DUE CUSTOMER .......... "; USING "#####,#.##"; changedue#
RECPT:
COLOR 11
     LPRINT "  T O W N   O F   T R O Y"
     LPRINT
     LPRINT "  Water/Sewer Department    "
     LPRINT "         "; DATE$
     LPRINT
     LPRINT "    Env. #:"; ENVELOPE
     LPRINT
     LPRINT "CUSTOMER: "; LEFT$(nme$, 18)
     LPRINT "   ACCT#: "; acctrecord
     LPRINT "    W/S Amt.: "; USING "####,#.##"; apply(1) + apply(2) + apply(3) + apply(5)
     LPRINT " B'kdoor P/U: "; USING "####,#.##"; apply(4)
     LPRINT " Bill Amount: "; USING "####,#.##"; amount#
     LPRINT "Amt Received: "; USING "####,#.##"; amtreceived#
     LPRINT "  Change Due: "; USING "####,#.##"; amtreceived# - amount#
     LPRINT
     LPRINT "Acct Balance: "; USING "####,#.##"; acctbalance# - amount#
     LPRINT
     LPRINT LEFT$(MESSAGE$, 29)
     LPRINT "    T H A N K   Y O U !"
     FOR JJ = 1 TO 5: LPRINT " ": NEXT JJ
     LPRINT CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)
1331 PUT 2, nexrecord
     PRINT CHR$(7); CHR$(7);
     CLOSE
        LOCATE 20, 1: COLOR 11
        PRINT "PRESS <ENTER> TO CONTINUE "; : fl = 1: GOSUB 62000
        GOTO 1000


history:
3100 CLS : GOSUB 3900
3120 tr! = CVD(firsttr$): GOSUB 61000
     GOSUB translink
3130 IF counter = 0 THEN 3210
     FOR kkk = counter TO 1 STEP -1
3140 GET 2, translink!(kkk)
3146 IF lc >= 18 THEN
      LOCATE 22, 1: COLOR 11
      PRINT "Press any key for next page! (ESC-Quit) .. "; : fl = 1: GOSUB 62000
      IF cf = 1 THEN CLOSE 2: RETURN
      CLS : GOSUB 3900: GOTO 3150
      ELSE
      GOTO 3150
     END IF
3150 PRINT TDATE$; TAB(10); TDESC$; TAB(30);
3160 IF CVI(TTYPE$) = 1 THEN PRINT "Charge"; : GOTO 3170
3162 IF CVI(TTYPE$) = 2 THEN PRINT "Late Chrge"; : GOTO 3170
3164 IF CVI(TTYPE$) = 3 THEN PRINT "Payment"; : GOTO 3170
3166 IF CVI(TTYPE$) = 4 THEN PRINT "Pay by Dep"; : GOTO 3170
3167 IF CVI(TTYPE$) = 5 THEN PRINT "Adjustment"; : GOTO 3170
     IF CVI(TTYPE$) = 6 THEN PRINT "Dep. Pymt"; : GOTO 3170
     IF CVI(TTYPE$) = 7 THEN PRINT "Dir. Dep."; : GOTO 3170
3168 PRINT "";
3170 PRINT TAB(45); USING "#####.##"; CVS(tcuramt$);
3171 PRINT TAB(60); USING "$$#####,.##"; CVD(tamt$): lc = lc + 1
3172 NEXT kkk
3210 LOCATE 22, 1: COLOR 11
     PRINT "That's All, Press any key to Continue ";
     fl = 1
     GOSUB 62000
     RETURN

translink:
counter = 0
trlink1:
IF tr! = 0 THEN RETURN
 counter = counter + 1
 translink!(counter) = tr!: GET 2, tr!
 tr! = CVS(tnexttr$): GOTO trlink1

3900 PRINT TAB(20); "Town of Troy : Water Customer Account History"
3910 PRINT "Date: "; DATE$:
     PRINT " Loc # "; m$; TAB(60); "Acct #"; record: PRINT ""
3920 PRINT "Account of ... "; : B = INSTR(firstname$, "  "): IF B = 0 THEN B = 20
3930 PRINT LEFT$(firstname$, B); lname$; TAB(55); "Balance: "; USING "$####.##"; (CVD(curbal$) + CVD(pastdue$)): PRINT ""
3940 PRINT "  Date"; TAB(10); "Description"; TAB(30); "Type"; TAB(44); "  Amount"; TAB(63); "Balance": PRINT STRING$(70, "="): lc = 6: RETURN



1700 OPEN "tbacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 16: FIELD 4, 12 AS met$, 4 AS ARREC$
     IF LEN(account$) < 8 THEN account$ = account$ + STRING$(8 - LEN(account$), 32)
 IF LOF(4) / 16 = 0 THEN RETURN
 HI = LOF(4) / 16: LO = 1: ME = HI: leng = LEN(account$)
 account# = VAL(account$)
 FOR XX = 1 TO 16: GET 4, ME
 IF account# = VAL(LEFT$(met$, leng)) THEN ng = 1: record = CVS(ARREC$): CLOSE 4: RETURN
 IF account# > VAL(LEFT$(met$, leng)) THEN LO = ME ELSE HI = ME
 ME = INT((HI - LO) / 2) + LO
 NEXT XX: CLOSE 4
 CLOSE 4: ng = 0: RETURN


1800 FOR ii = 3 TO 20: LOCATE ii, 1: PRINT STRING$(79, 32): NEXT ii: LOCATE 3, 1
     COLOR 11: LOCATE 3, 1: PRINT "Enter Payee Name (ESC-MENU).. "; : fl = 30: GOSUB 62000
     IF cf = 1 THEN 1000 ELSE nme$ = in$
     COLOR 11: LOCATE 4, 1: PRINT "Enter Address ............... "; : fl = 27: GOSUB 62000: DESC$ = in$
     COLOR 11: LOCATE 5, 1: PRINT "Enter Amount Owed ........... "; : fl = -10: GOSUB 62000: owed = VAL(in$)
     COLOR 11: LOCATE 6, 1: PRINT "Enter Amount Paid ........... "; : fl = -10: GOSUB 62000: amount# = VAL(in$)
     COLOR 11: LOCATE 6, 50: PRINT "Change: "; : COLOR 15: PRINT USING "####.##"; amount# - owed
1805 COLOR 11: LOCATE 7, 1: PRINT "Enter (1)Cash or (2)Check ... "; : fl = -1: GOSUB 62000
     IF VAL(in$) < 1 OR VAL(in$) > 2 THEN 1805
     curency = VAL(in$)
1808 COLOR 11: LOCATE 8, 1: PRINT "Enter Misc Source Code (1-12). "; : fl = -2: GOSUB 62000
     IF VAL(in$) < 1 OR VAL(in$) > 12 THEN 1808
     mpcode = VAL(in$)
     GOSUB misccodefile
     GET 15, mpcode
     LOCATE 8, 45
      COLOR 15
      PRINT CODEDESC$: CATAGORY$ = CODEDESC$
1810 COLOR 11: LOCATE 11, 1: PRINT "Is the Above Correct (Y/N) .. "; : fl = 1: GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN 1810 ELSE LOCATE 11, 31: PRINT in$
     IF in$ = "N" THEN FOR J = 3 TO 11: LOCATE J, 1: PRINT STRING$(79, 32): NEXT J: GOTO 1800
    
     OPEN OPERATORFILE$ FOR RANDOM ACCESS READ WRITE AS #2 LEN = 128
     FIELD 2, 2 AS psource$, 8 AS pmeter$, 4 AS prec$, 8 AS pamt$, 15 AS pdesc$, 2 AS ptype$, 30 AS pname$, 4 AS TFDIST$, 4 AS TLDIST$, 4 AS trev1$, 4 AS trev2$, 4 AS TREV3$, 4 AS trev4$, 4 AS trev5$, 4 AS trev6$, 8 AS PDATE$
     nexrecord = LOF(2) / 128 + 1

      CLOSE 8: OPEN "MESSAGE.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #8 LEN = 40
       FIELD 8, 40 AS MESSAGE$: IF LOF(8) / 40 = 0 THEN GOTO RECPT2
       GET 8, 1
FOR LL = 12 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
COLOR 11
LOCATE 15, 1: PRINT "Amount Paid on Account ....... "; USING "#####,#.##"; owed
              PRINT "Cash Received from Customer .. "; USING "#####,#.##"; amount#
              PRINT "                               -------------"
              COLOR 15
              PRINT "CHANGE DUE CUSTOMER .......... "; USING "#####,#.##"; amount# - owed

RECPT2:

     LPRINT "  T O W N   O F   T R O Y"
     LPRINT
     LPRINT "    RECEIPT FOR PAYMENT"
     LPRINT "    "; CATAGORY$
     LPRINT "        "; DATE$
     LPRINT "    Env. #:"; ENVELOPE
     LPRINT
     LPRINT "Customer: "; LEFT$(nme$, 18)
     LPRINT ".."; DESC$
     LPRINT
     LPRINT "Bill Amount: "; USING "$####,#.##"; owed
     LPRINT "Amount Paid: "; USING "$####,#.##"; amount#
     LPRINT " Change Due: "; USING "$####,#.##"; amount# - owed
     LPRINT
     LPRINT LEFT$(MESSAGE$, 29)
     LPRINT "    T H A N K   Y O U  !"
     FOR JJ = 1 TO 5
     LPRINT " "
     NEXT JJ
     LPRINT CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)

     LSET psource$ = MKI$(4): LSET PDATE$ = sysdate$
     LSET pmeter$ = "": LSET prec$ = MKS$(mpcode)
     LSET pamt$ = MKD$(amount#): LSET pdesc$ = "ENV #" + STR$(ENVELOPE)
     LSET pname$ = nme$: LSET TFDIST$ = MKS$(0): LSET TLDIST$ = MKS$(0)
     LSET ptype$ = MKI$(curency + 8): LSET tcuryr$ = MKS$(0)
     LSET tpastyr$ = MKS$(0)
     LSET trev1$ = MKS$(apply(1)): LSET trev2$ = MKS$(apply(2)): LSET TREV3$ = MKS$(apply(3))
     LSET trev4$ = MKS$(apply(4)): LSET trev5$ = MKS$(apply(5)): LSET trev6$ = MKS$(apply(6))
     LSET TFDIST$ = MKS$(0): LSET TLDIST$ = MKS$(0)
     PUT 2, nexrecord: CLOSE
     LOCATE 20, 1: COLOR 11
     PRINT "PRESS <ENTER> TO CONTINUE "; : fl = 1: GOSUB 62000
     GOTO 1000
     



1900 COLOR 11: LOCATE 19, 1: PRINT "Payment Codes Are As Follows: "
1910 COLOR 14: LOCATE 20, 1: PRINT "1= Reg Payment - CASH      2= Reg Payment - CHECK      3= Apply Deposit"
1920 LOCATE 21, 1: PRINT "4= Deposit Payment - CASH  5= Deposit Payment - CHECK  6= ADJUSTMENT"
1930 RETURN
1950 IF PCODE = 1 THEN dd$ = "Reg Payment - CASH"
     IF PCODE = 2 THEN dd$ = "Reg Payment - CHECK"
     IF PCODE = 3 THEN dd$ = "Apply Deposit"
     IF PCODE = 4 THEN dd$ = "Deposit Payment - CASH"
     IF PCODE = 5 THEN dd$ = "Deposit Payment - CHECK"
     IF PCODE = 6 THEN dd$ = "ADJUSTMENT/CREDIT"
     IF PCODE = 7 THEN dd$ = "Direct Deposit"
1960 LOCATE 19, 1: PRINT STRING$(79, 32)
      PRINT STRING$(79, 32): PRINT STRING$(79, 32)
      PRINT STRING$(79, 32);
      RETURN
namesearch:
LOCATE 3, 1:
COLOR 11: PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : fl = 25
 GOSUB 62000: IF cf = 1 THEN ggo = 0: RETURN
 nn$ = in$
 nn = LEN(nn$)
 CLOSE 4: OPEN "R", 4, "wblname.dat", 24: FIELD 4, 20 AS lastnam$, 4 AS mrec$
 CLOSE 1: GOSUB 60000
FOR LL = 1 TO LOF(4) / 24: GET 4, LL
IF LEFT$(lastnam$, nn) = nn$ THEN GOTO nn1
n3: NEXT LL
ggo = 0: RETURN
nn1:
GET 1, CVS(mrec$)
LOCATE 6, 1: COLOR 11
kk = INSTR(firstname$, "  ")
IF kk = 0 THEN kk = 25
IF kk = 1 THEN nme$ = lname$ ELSE nme$ = LEFT$(firstname$, kk) + lname$
bb = INSTR(meter$, "/")
IF bb = 0 THEN bb = 9
m$ = LEFT$(meter$, bb - 1)
COLOR 10
PRINT "Full Account Name: "; nme$
PRINT " Account Location: "; m$
PRINT "   Account Number: "; CVS(mrec$)
ACCTNUMBER$ = m$
PRINT
n2:
LOCATE 9, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO n2
IF in$ = "Q" THEN ggo = 0: RETURN
IF in$ = "Y" THEN ggo = 1: record = CVS(mrec$): RETURN
IF in$ = "N" THEN GOTO n3
GOTO n2

TAXENTRY:
     CLOSE 1
     OPEN "R", 1, "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
     CURRENTYEAR$ = year$
     CLOSE 1
     scren$ = "5.01.00": CLS : GOSUB 63500
     LOCATE 1, 30: PRINT STRING$(20, 32)
     COLOR 11
     LOCATE 1, 33: PRINT "Tax Payment Entry"
     LOCATE 3, 1
     CLOSE 1: OPEN "R", 1, "tbmpaid.dat", 256
     FIELD 1, 1 AS chargecurrent$, 6 AS paydate$, 4 AS ENVELOPE$, 2 AS TAXYEAR$
     IF LOF(1) = 0 THEN
      CLOSE : LOCATE 3, 1: PRINT "Tax Payment Entry Defaults Not Set ": PRINT CHR$(7); : FOR XX = 1 TO 250: NEXT XX:
      GOTO 6000
     END IF
     GET 1, 1
     chargecurrentinterest$ = chargecurrent$: paymentdate$ = paydate$: ENV = CVS(ENVELOPE$)
     curyr = VAL(CURRENTYEAR$)
     EYEAR = CVI(TAXYEAR$) + 1900: REM good for 20th century
     IF EYEAR > curyr THEN
      LOCATE 3, 1: PRINT "Year Specified as Entry Year Exceeds Current Year"
      LOCATE 4, 1: PRINT "Please Enter New Tax Entry Year (yy) .. "; : fl = -2
      GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 1000
      LSET TAXYEAR$ = MKI$(VAL(in$))
      PUT 1, 1: CLOSE
      GOTO TAXENTRY
     END IF
     IF EYEAR = 1900 THEN
      LOCATE 3, 1: PRINT "Year Specified as Entry Year is Invalid"
      LOCATE 4, 1: PRINT "Please Enter New Tax Entry Year (yy) .. "; : fl = -2
      GOSUB 62000: LSET TAXYEAR$ = MKI$(VAL(in$))
      PUT 1, 1: CLOSE
      GOTO TAXENTRY
      END IF

 AMOUNTPAID = 0: amounttendered = 0: discount = 0
 tx = 0: in = 0: ad = 0: col = 0: lat = 0
 totcredit = 0: stillremaining = 0: remaining = 0
 disttx = 0: distin = 0: distad = 0: distco = 0: distdog = 0
   FOR kk = 0 TO 9
    runtot(kk) = 0: runtot1(kk) = 0: runtot2(kk) = 0: runtot3(kk) = 0: runtot4(kk) = 0
   NEXT kk
 DEPFLAG = 0
     LOCATE 3, 1: COLOR 11
     PRINT "Enter Customer Account Number (TAB-Last Name : ESC-Menu) .. "; : fl = 7
     GOSUB 62000:
     IF cf = 1 THEN
      CLOSE : GOTO 1000
     END IF
     IF cf = 2 THEN
      GOSUB Tnamesearch
      IF ggo = 1 THEN
       FOR LL = 3 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
       LOCATE 3, 1: COLOR 11
       GOTO TE1
      END IF
      IF ggo = 0 THEN LOCATE 10, 1: PRINT "Account Not Found : Press Any Key to Continue "; : fl = 1: GOSUB 62000: GOTO TAXENTRY
     END IF

 account$ = in$
GOSUB findacct
IF ng = 0 THEN
 LOCATE 8, 1: COLOR 11
 PRINT "Account Not Located : Press <ENTER> to Continue "; : fl = 1: GOSUB 62000: CLOSE : GOTO TAXENTRY
END IF
TE1: GOSUB 60000.1
GET 14, record
FOR LL = 3 TO 23: LOCATE 3, LL: PRINT STRING$(79, 32): NEXT LL
COLOR 11
kk = INSTR(firstnam$, "  "): IF kk = 0 THEN kk = 25
IF kk = 1 THEN nme$ = lnam$ ELSE nme$ = LEFT$(firstnam$, kk) + lnam$
LOCATE 3, 1: COLOR 11
PRINT "Acct # "; : COLOR 15: PRINT acct1$;
 COLOR 11: PRINT "Name: "; : COLOR 15: PRINT nme$

IF CVS(firstparcel$) = 0 THEN
 LOCATE 8, 1: COLOR 11
 PRINT "No Parcels On File For This Account : Press <ENTER> to Continue "; : fl = 1
 GOSUB 62000
 CLOSE

 GOTO TAXENTRY
END IF

1001 LOCATE 5, 1: COLOR 11
     PRINT "Enter Parcel Number (F1-Scroll or ESC-End) .. "; : fl = 5: GOSUB 62000
     IF cf = 1 THEN
      CLOSE : GOTO TAXENTRY
     END IF

     IF cf = 11 THEN
        PARCELrecord = CVS(firstparcel$)
1002 IF PARCELrecord = 0 THEN
         FOR LL = 3 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
         LOCATE 5, 1: COLOR 11: PRINT "No More Parcels For This Customer "
         PRINT "Press <ENTER> to Continue "; : fl = 1: GOSUB 62000
         CLOSE : GOTO TAXENTRY
     END IF
      GET 1, PARCELrecord
          GET 3, PARCELrecord
      LOCATE 5, 1: PRINT STRING$(79, 32)
      COLOR 11: LOCATE 5, 1
      PRINT "Parcel # "; pin$
1003 COLOR 11: LOCATE 6, 1
      PRINT "Is This The Right Parcel (Y/N) .. "; : fl = 1
      GOSUB 62210
      IF NOT (in$ = "Y" OR in$ = "N") THEN 1003
      IF in$ = "N" THEN
       PARCELrecord = CVS(nexparcel$)
       FOR LL = 5 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
       GOTO 1002
       ELSE
       GOTO 1007
     END IF
   END IF

        parcel$ = in$: IF LEN(parcel$) < 5 THEN parcel$ = parcel$ + STRING$(5 - LEN(parcel$), 32)
        PARCELrecord = CVS(firstparcel$)
1005 IF PARCELrecord = 0 THEN
         FOR LL = 3 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
         LOCATE 5, 1: COLOR 11: PRINT "No More Parcels For This Customer "
         PRINT "Press <ENTER> to Continue "; : fl = 1: GOSUB 62000
         CLOSE : GOTO TAXENTRY
     END IF
     GET 1, PARCELrecord: GET 3, PARCELrecord
     IF pin$ = parcel$ THEN
     GOTO 1007
     END IF
     PARCELrecord = CVS(nexparcel$): GOTO 1005

1007 LOCATE 23, 1: PRINT STRING$(79, 32);
1090.1 FOR zz = 5 TO 23: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz
     LOCATE 5, 1: COLOR 11
1100.1 PRINT "Customer Account of ..."
     COLOR 15
1110.1 IF LEFT$(firstnam$, 1) = "  " THEN nme$ = lnam$: PRINT TAB(20); lnam$: GOTO 1140
1120.1 B = INSTR(firstnam$, "  "): IF B = 0 THEN B = 20
1130.1 PRINT TAB(10); LEFT$(firstnam$, B); lnam$: nme$ = LEFT$(firstnam$, B) + lnam$
1140.1 PRINT TAB(10); A1$
1150.1 PRINT TAB(10); A2$
1160.1 PRINT TAB(10); cy$; " "; st$; " "; zp$
1170.1 COLOR 11:
      PRINT STRING$(79, "=")
1180.1 COLOR 15: LOCATE 10, 60: PRINT "TAX YR SET AS "; EYEAR
       COLOR 11
       LOCATE 13, 1: PRINT "Correct Account (Y/N)(F1-CHANGE TAX YEAR) .. "; : fl = 1
       GOSUB 62210:
       IF cf = 11 THEN
     CLOSE 13: OPEN "R", 13, "tbmpaid.dat", 256
     FIELD 13, 1 AS chargecurrent$, 6 AS paydate$, 4 AS ENVELOPE$, 2 AS TAXYEAR$
     LOCATE 15, 1: PRINT "ENTER TAX YEAR BEING PAID (YY) .. "; : fl = -2: GOSUB 62000
     LSET TAXYEAR$ = MKI$(VAL(in$))
     PUT 13, 1
     EYEAR = CVI(TAXYEAR$) + 1900
     CLOSE 13
     LOCATE 15, 1: PRINT STRING$(79, 32)
     GOTO 1180.1
     END IF


       IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO 1180.1
1190.1 IF in$ = "N" THEN CLOSE : GOTO TAXENTRY

paymententry:
 cmth = VAL(LEFT$(paymentdate$, 2)): CURRENTYEAR = VAL(RIGHT$(paymentdate$, 2)) + 1900
 FOR LL = 11 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
     curyr = VAL(CURRENTYEAR$)
 IF EYEAR = curyr THEN
 principal = CVS(curprinciple$) + CVS(curlateamt$)
 PAYPRINCIPAL = CVS(curtaxamt$) + CVS(curadvamt$) + CVS(curcolamt$) + CVS(curlateamt$)
  IF chargecurrentinterest$ = "Y" THEN
     interestperc = 2 + (.75 * (cmth - 1))
     INTEREST = (interestperc / 100) * principal
     INTEREST = INTEREST - CVS(curintamt$): REM reduce any interest amount paid
     INTEREST = INT((INTEREST * 100) + .5) / 100
     IF INTEREST < 0 THEN INTEREST = 0
    ELSE
    END IF
   GOTO pe1
 END IF
calcyear = (curyr - EYEAR) - 1
years = CURRENTYEAR - EYEAR - 1

 IF calcyear > 9 THEN calcyear = 9
  principal = CVS(pprinciple$(calcyear)) + CVS(plateamt$(calcyear))
  PAYPRINCIPAL = CVS(ptaxamt$(calcyear)) + CVS(padvamt$(calcyear)) + CVS(pcolamt$(calcyear)) + CVS(plateamt$(calcyear))
  interestperc = 2 + (.75 * (cmth - 1)) + ((12 * .75) * years)
    IF years >= 1 AND cmth = 1 THEN interestperc = interestperc - .75: REM 2% already added for one month
  INTEREST = (interestperc / 100) * principal
  INTEREST = INTEREST - CVS(pintamt$(calcyear)): REM reduce any interest amount paid
  INTEREST = INT((INTEREST * 100) + .5) / 100
  IF INTEREST < 0 THEN INTEREST = 0

pe1:
 LOCATE 11, 1: COLOR 11
 PRINT "* Payment Calculation *"
 PRINT "Tax Amount ........ "; : COLOR 15: PRINT USING "$$######,#.##"; PAYPRINCIPAL
 COLOR 11
 PRINT "Interest Amount ... "; : COLOR 15: PRINT USING "$$######,#.##"; INTEREST;
  PRINT TAB(50); "Total Owed = "; : COLOR 15: PRINT USING "$$######,#.##"; PAYPRINCIPAL + INTEREST: TAXOWED = PAYPRINCIPAL + INTEREST
 LOCATE 14, 1: COLOR 11
 PRINT "Interest Correct (Y/N) .. "; : fl = 1: GOSUB 62210
 IF cf = 1 THEN CLOSE : GOTO 1000
 IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO pe1
 IF in$ = "N" THEN
  PRINT "Enter New Interest Amount Owed .. "; : fl = -9: GOSUB 62000
  LOCATE 15, 1: PRINT STRING$(79, 32)
  INTEREST = VAL(in$)
  GOTO pe1
 END IF
 COLOR 11
 LOCATE 16, 1: PRINT STRING$(79, 32):
 LOCATE 16, 1: COLOR 11
 PRINT "Enter the Amount Paid ......... "; : fl = -9: GOSUB 62000
 IF cf = 1 THEN CLOSE : GOTO 1000
 AMOUNTPAID = VAL(in$)
 LOCATE 16, 33: PRINT USING "$$####,#.##"; AMOUNTPAID
 COLOR 11
 PRINT "Enter the Discount Allowed .... "; : fl = -9: GOSUB 62000
  IF cf = 1 THEN CLOSE : GOTO 1000
  discount = VAL(in$)
  LOCATE 17, 33: PRINT USING "$$####,#.##"; discount
 COLOR 11
 PRINT "Payment Description ........... "; : fl = 20: GOSUB 62000
  IF cf = 1 THEN CLOSE : GOTO 1000
  description$ = in$
pe1a: COLOR 11
 LOCATE 19, 1
 PRINT "Payment Type 1=Pymt : 2=Rel. .. "; : fl = -1: GOSUB 62000
  IF cf = 1 THEN CLOSE : GOTO 1000
  IF VAL(in$) < 1 OR VAL(in$) > 2 THEN GOTO pe1a:
  paytype = VAL(in$)
pe2:
 COLOR 11
 IF paytype = 1 THEN
  relyr = 0
 END IF
 IF paytype = 2 THEN
   relyr = EYEAR
 END IF
pe3a:
 paydate$ = paymentdate$
 envnum = ENV
pe4:
COLOR 11
LOCATE 20, 1: PRINT "Enter the Amount Tendered .. "; : fl = -9: GOSUB 62000
amounttendered = VAL(in$)
LOCATE 20, 50: COLOR 11
PRINT "Change= "; : COLOR 15: PRINT USING "####,#.##"; (amounttendered - AMOUNTPAID)

pe5:
 COLOR 11
 LOCATE 21, 1
  PRINT "Is the above information correct (Y/N) .. "; : fl = 1
  GOSUB 62210
  IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO pe5
  IF in$ = "Y" THEN GOTO processpayment
  GOTO paymententry

processpayment:
 totcredit = AMOUNTPAID + discount
 GOSUB tempfile

autoprocess:
IF EYEAR = curyr THEN
 IF INTEREST > 0 THEN
  IF totcredit > INTEREST THEN
    LSET Pcin$ = MKS$(INTEREST): totcredit = totcredit - INTEREST
     ELSE
    LSET Pcin$ = MKS$(INTEREST): totcredit = 0: GOTO endautoprocess
  END IF
 END IF

IF CVS(curlateamt$) > 0 THEN
  IF totcredit > CVS(curlateamt$) THEN
   LSET pcla$ = MKS$(CVS(curlateamt$)): totcredit = totcredit - CVS(curlateamt$)
   ELSE
   LSET pcla$ = MKS$(totcredit): totcredit = 0: GOTO endautoprocess
  END IF
END IF

IF CVS(curcolamt$) > 0 THEN
  IF totcredit > CVS(curcolamt$) THEN
   LSET Pcco$ = MKS$(CVS(curcolamt$)): totcredit = totcredit - CVS(curcolamt$)
   ELSE
   LSET Pcco$ = MKS$(totcredit): totcredit = 0: GOTO endautoprocess
  END IF
END IF
IF CVS(curadvamt$) > 0 THEN
  IF totcredit > CVS(curadvamt$) THEN
   LSET Pcad$ = MKS$(CVS(curadvamt$)): totcredit = totcredit - CVS(curadvamt$)
   ELSE
   LSET Pcad$ = MKS$(totcredit): totcredit = 0: GOTO endautoprocess
  END IF
END IF
IF totcredit > 0 THEN
   LSET Pctx$ = MKS$(totcredit): totcredit = 0: GOTO endautoprocess
END IF
END IF

REM prior year
calcyear = (curyr - EYEAR) - 1
IF calcyear > 9 THEN calcyear = 9
ap1:
LL = calcyear
IF INTEREST > 0 THEN
  IF totcredit > INTEREST THEN
   LSET pin$(LL) = MKS$(INTEREST): totcredit = totcredit - INTEREST
   ELSE
   LSET pin$(LL) = MKS$(totcredit): totcredit = 0: GOTO endautoprocess
  END IF
END IF

IF CVS(plateamt$(LL)) > 0 THEN
  IF totcredit > CVS(plateamt$(LL)) THEN
   LSET pla$(LL) = MKS$(CVS(plateamt$(LL))): totcredit = totcredit - CVS(plateamt$(LL))
   ELSE
   LSET pla$(LL) = MKS$(totcredit): totcredit = 0: GOTO endautoprocess
  END IF
END IF
IF CVS(pcolamt$(LL)) > 0 THEN
  IF totcredit > CVS(pcolamt$(LL)) THEN
   LSET pco$(LL) = MKS$(CVS(pcolamt$(LL))): totcredit = totcredit - CVS(pcolamt$(LL))
   ELSE
   LSET pco$(LL) = MKS$(totcredit): totcredit = 0: GOTO endautoprocess
  END IF
END IF
IF CVS(padvamt$(LL)) > 0 THEN
  IF totcredit > CVS(padvamt$(LL)) THEN
   LSET pad$(LL) = MKS$(CVS(padvamt$(LL))): totcredit = totcredit - CVS(padvamt$(LL))
   ELSE
   LSET pad$(LL) = MKS$(totcredit): totcredit = 0: GOTO endautoprocess
  END IF
END IF
IF CVS(ptaxamt$(LL)) > 0 THEN
  IF totcredit > CVS(ptaxamt$(LL)) THEN
   LSET ptx$(LL) = MKS$(CVS(ptaxamt$(LL))): totcredit = totcredit - CVS(ptaxamt$(LL))
   ELSE
   LSET ptx$(LL) = MKS$(totcredit): totcredit = 0: GOTO endautoprocess
  END IF
END IF

ap2:


endautoprocess:
acctbalance# = TAXOWED - AMOUNTPAID - discount
FOR LL = 12 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
COLOR 11
LOCATE 15, 1: PRINT "Amount Paid on Account ....... "; USING "#####,#.##"; AMOUNTPAID
              PRINT "Cash Received from Customer .. "; USING "#####,#.##"; amounttendered
              PRINT "                               -------------"
              COLOR 15
              PRINT "CHANGE DUE CUSTOMER .......... "; USING "#####,#.##"; amounttendered - AMOUNTPAID


     LPRINT "  T O W N   O F   T R O Y"
     LPRINT
     LPRINT "        TAX RECEIPT"
     LPRINT "         "; DATE$
     LPRINT "    Env. #:"; ENVELOPE
     LPRINT
     LPRINT "CUSTOMER: "; LEFT$(nme$, 18)
     LPRINT "   ACCT#: "; acct1$
     LPRINT " PARCEL#: "; pin$
     LPRINT "TAX YEAR: "; EYEAR
     LPRINT
     LPRINT "Principle Amt: "; USING "$#####,#.##"; AMOUNTPAID - INTEREST
     LPRINT " Interest Amt: "; USING "$#####,#.##"; INTEREST
     LPRINT " Total Amount: "; USING "$#####,#.##"; AMOUNTPAID
     LPRINT " Amt Received: "; USING "$#####,#.##"; amounttendered
     LPRINT "   Change Due: "; USING "$#####,#.##"; amounttendered - AMOUNTPAID
     LPRINT
     LPRINT
     LPRINT "Acct Balance: "; USING "$####,#.##"; acctbalance#
     LPRINT
     LPRINT LEFT$(MESSAGE$, 29)
     LPRINT "    T H A N K   Y O U !"
     FOR JJ = 1 TO 5: LPRINT " ": NEXT JJ
     LPRINT CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)




LSET PDATE$ = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 2)
LSET pamt$ = MKS$(AMOUNTPAID)
LSET pdisamt$ = MKS$(discount)
LSET PTAXYR$ = MKS$(EYEAR)
LSET PDESC1$ = "ENV #" + STR$(ENVELOPE)
LSET ptype$ = MKI$(paytype)
LSET prec$ = MKS$(PARCELrecord)
LSET penvnum$ = MKI$(ENVELOPE)
LSET ppaydate$ = paydate$
PUT 8, nexrecord
PUT 9, nexrecord
GET 9, nexrecord
FOR LL = 0 TO 9
tx = tx + CVS(ptx$(LL))
in = in + CVS(pin$(LL))
ad = ad + CVS(pad$(LL))
col = col + CVS(pco$(LL))
lat = lat + CVS(pla$(LL))
NEXT LL
LSET Ptax$ = MKS$(tx + CVS(Pctx$))
LSET Pint$ = MKS$(in + CVS(Pcin$))
LSET Padv$ = MKS$(ad + CVS(Pcad$))
LSET Pcol$ = MKS$(col + CVS(Pcco$))
LSET plate$ = MKS$(lat + CVS(pcla$))
PUT 8, nexrecord
CLOSE
LOCATE 20, 1: COLOR 11
PRINT "PRESS <ENTER> TO CONTINUE "; : fl = 1: GOSUB 62000
GOTO 1000

manualprocess:
OPEN "R", 5, "TBYEAR.DAT", 4: FIELD 5, 4 AS year$
GET 5, 1
   TAXYEAR$ = year$: CLOSE 5

 CLS
 COLOR 11
 PRINT "** Manual Distribution of Payment "
 PRINT "** Account of "; nme$
 PRINT "** Current Year First": PRINT
 PRINT "Tax Yr"; TAB(12); "Tax Amt"; TAB(24); "Int. Amt"; TAB(36); "Adv Amt"; TAB(48); "Col Amt"; TAB(59); "Dog Fee"; TAB(70); "REMAINING"
 PRINT STRING$(79, "-")
 TAXYEAR = VAL(TAXYEAR$)
 remaining = totcredit - disttx - distin - distad - distco - distdog
 remaining = INT((remaining * 100) + .5) / 100
 COLOR 15
 PRINT USING "####"; TAXYEAR;
  PRINT TAB(12); USING "$#####.##"; CVS(curtaxamt$);
   PRINT TAB(24); USING "$####.##"; CVS(curintamt$);
    PRINT TAB(36); USING "$####.##"; CVS(curadvamt$);
     PRINT TAB(48); USING "$####.##"; CVS(curcolamt$);
      PRINT TAB(59); USING "$####.##"; CVS(curlateamt$);
      PRINT TAB(70); USING "$####,#.##"; remaining
 COLOR 10
  PRINT " 1-"
mp1:
COLOR 11
 LOCATE 22, 1: PRINT "Enter Item to Change (ESC-Continue) .. "; : fl = -1
 GOSUB 62000: IF cf = 1 THEN GOTO mp2
LOCATE 8, 12
  fl = -9: GOSUB 62000: LSET Pctx$ = MKS$(VAL(in$)): disttx = VAL(in$)
 LOCATE 8, 24
  fl = -9: GOSUB 62000: LSET Pcin$ = MKS$(VAL(in$)): distin = VAL(in$)
 LOCATE 8, 36
  fl = -9: GOSUB 62000: LSET Pcad$ = MKS$(VAL(in$)): distad = VAL(in$)
 LOCATE 8, 48
  fl = -9: GOSUB 62000: LSET Pcco$ = MKS$(VAL(in$)): distco = VAL(in$)
 LOCATE 8, 59
  fl = -9: GOSUB 62000: LSET pcla$ = MKS$(VAL(in$)): distdog = VAL(in$)

 remaining = totcredit - disttx - distin - distad - distco - distdog
 remaining = INT((remaining * 100) + .5) / 100
 LOCATE 8, 70: PRINT USING "$####,#.##"; remaining
GOTO mp1

mp2:
 CLS
 COLOR 11
 stillremaining = remaining
 PRINT "** Prior Years"
 PRINT "Tax Yr"; TAB(14); "Tax Amt"; TAB(24); "Int. Amt"; TAB(36); "Adv Amt"; TAB(48); "Col Amt"; TAB(59); "Dog Fee"; TAB(70); "REMAINING"
 FOR LL = 1 TO 10
 COLOR 15

 PRINT USING "####"; TAXYEAR - LL;
  PRINT TAB(12); USING "$#####.##"; CVS(ptaxamt$(LL - 1));
  PRINT TAB(24); USING "$####.##"; CVS(pintamt$(LL - 1));
  PRINT TAB(36); USING "$####.##"; CVS(padvamt$(LL - 1));
     PRINT TAB(48); USING "$####.##"; CVS(pcolamt$(LL - 1));
      PRINT TAB(59); USING "$####.##"; CVS(plateamt$(LL - 1))

 COLOR 10
 PRINT LL; "-"
 COLOR 11
NEXT LL
LOCATE 4, 70: PRINT USING "$$####,#.##"; stillremaining
mp3:
COLOR 11
LOCATE 23, 1: PRINT "Enter Item to Change (ESC-End:F1 to Cancel) .. "; : fl = -2: GOSUB 62000
IF cf = 11 THEN CLOSE : GOTO 1000
IF cf = 1 THEN
 IF stillremaining = 0 THEN GOTO endautoprocess
LOCATE 23, 1: PRINT CHR$(7); "Amount Paid NOT Fully Distributes : Press <ENTER> to Continue "; : fl = 1: GOSUB 62000
LOCATE 23, 1: PRINT STRING$(78, 32);
GOTO mp3
END IF
c = VAL(in$)
IF c < 1 OR c > 10 THEN GOTO mp3
IF c = 1 THEN cc = 4 ELSE cc = c * 2 + 2

mp01:
summary = 0
LOCATE cc, 12: fl = -9: GOSUB 62000: LSET ptx$(c - 1) = MKS$(VAL(in$)): runtot(c - 1) = VAL(in$)
 LOCATE cc, 24: fl = -9: GOSUB 62000: LSET pin$(c - 1) = MKS$(VAL(in$)): runtot1(c - 1) = VAL(in$)
  LOCATE cc, 36: fl = -9: GOSUB 62000: LSET pad$(c - 1) = MKS$(VAL(in$)): runtot2(c - 1) = VAL(in$)
   LOCATE cc, 48: fl = -9: GOSUB 62000: LSET pco$(c - 1) = MKS$(VAL(in$)): runtot3(c - 1) = VAL(in$)
    LOCATE cc, 59: fl = -9: GOSUB 62000: LSET pla$(c - 1) = MKS$(VAL(in$)): runtot4(c - 1) = VAL(in$)
   FOR kk = 0 TO 9
    summary = summary + (runtot(kk) + runtot1(kk) + runtot2(kk) + runtot3(kk) + runtot4(kk))
   NEXT kk
   stillremaining = remaining - summary
   stillremaining = INT((stillremaining * 100) + .5) / 100
   LOCATE 4, 70: PRINT USING "$$####,#.##"; stillremaining
   runtot = 0
   GOTO mp3


figurebalance:
     curbal# = 0: pastdue# = 0
     curbal# = (CVS(curtaxamt$) + CVS(curadvamt$) + CVS(curcolamt$) + CVS(curlateamt$))
     curbal# = INT((curbal# * 100) + .5) / 100
     FOR LL = 0 TO 9
     pastdue# = pastdue# + (CVS(ptaxamt$(LL)) + CVS(padvamt$(LL)) + CVS(plateamt$(LL)) + CVS(pcolamt$(LL)))
     pastdue# = INT((pastdue# * 100) + .5) / 100
     NEXT LL
     RETURN




arentry:
     FOR JJ = 1 TO 8: apply(JJ) = 0: NEXT JJ
     scren$ = "4.01.00": CLS : GOSUB 63500
     LOCATE 1, 34: COLOR 11: PRINT "A/R Payment Entry"; : LOCATE 3, 1
     COLOR 11: PRINT "Enter Account Number Please (ESC-Menu : TAB - Name Search) .. "; : fl = -5
     GOSUB 62000
     IF cf = 1 THEN 1000
     IF cf = 2 THEN GOSUB arnamesearch: IF ggo = 1 THEN GOTO 1080.1 ELSE CLOSE : GOTO arentry
     record = VAL(in$)
1080.1 CLOSE 1: GOSUB 60000.2
     IF record < 1 OR record > LOF(1) / 512 THEN CLOSE : GOTO 1000
     GET 1, record: acctrecord = record
     IF active$ <> "A" THEN
      LOCATE 23, 1: COLOR 11: PRINT "Account Deleted : Press <ENTER> to Continue "; : fl = 1
      GOSUB 62000: CLOSE : GOTO arentry
     END IF
     LOCATE 3, 1: PRINT STRING$(79, 32)
     LOCATE 4, 1: PRINT STRING$(79, 32)
     LOCATE 5, 1: PRINT STRING$(79, 32)
     LOCATE 6, 1: PRINT STRING$(79, 32)
     LOCATE 7, 1: PRINT STRING$(79, 32)
     COLOR 15: LOCATE 3, 1: PRINT "  Account Number: "; USING "#####"; record
1100.2 COLOR 11: PRINT "Customer Account of ..."; TAB(45); "Balance = "; USING "####.##"; CVD(curbal$) + CVD(pastdue$): owed# = CVD(curbal$) + CVD(pastdue$): owed# = INT((owed# * 100) + .5) / 100
      COLOR 14:
  PRINT TAB(5); LEFT$(nme$, 39)
  PRINT TAB(5); add1$
  PRINT TAB(5); add2$
  PRINT TAB(5); city$; " "; st$; " "; zp$
  LOCATE 10, 1: COLOR 11: PRINT STRING$(44, "-")
  LOCATE 11, 1: COLOR 11: PRINT "Correct Account (Y/N) .. ";
  fl = 1: GOSUB 62210
  IF in$ = "Y" THEN GOTO 1200.2 ELSE CLOSE : GOTO arentry
1200.2 FOR LL = 11 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
 GOSUB 1900.2
 LOCATE 11, 1
 COLOR 11
 PRINT "Enter Payment Code (See Below 1-3) .. ";
 fl = -1: GOSUB 62000
 IF cf = 1 THEN CLOSE : GOTO arentry ELSE PCODE = VAL(in$)
 IF PCODE < 1 OR PCODE > 3 THEN GOTO 1200.2
 GOSUB 1950.2: COLOR 11: LOCATE 11, 55: PRINT dd$
 LOCATE 12, 1: COLOR 11: PRINT "Enter Amount ... "; : PRINT "("; USING "$####.##"; owed#; : PRINT ") ......... "; : fl = -8: GOSUB 62000
 IF LEN(in$) = 0 THEN amount# = owed#: LOCATE 12, 39: PRINT USING "$####.##"; amount#:  ELSE amount# = VAL(in$): LOCATE 12, 39: PRINT USING "$####.##"; amount#
 LOCATE 13, 1: COLOR 11: PRINT "Enter Description (Ck #) ............ "; : fl = 15: GOSUB 62000: des$ = in$
 LOCATE 14, 1: COLOR 11: PRINT "Enter Amount Tendered ............... "; : fl = -9: GOSUB 62000
  LOCATE 14, 50: COLOR 11: PRINT "CHANGE: "; : COLOR 15: PRINT USING "$$####,#.##"; VAL(in$) - amount#
  amtreceived# = VAL(in$): changedue# = amtreceived# - amount#
1200.3 LOCATE 15, 1: COLOR 11: PRINT "Is the Above Correct (Y/N) .......... "; : fl = 1: GOSUB 62210
  IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO 1200.3 ELSE LOCATE 15, 39: PRINT in$
 IF in$ = "N" THEN
    GOTO 1200.2
 END IF

 aroperatorfile$ = "arpay" + filen$ + ".dat"
 CLOSE 2: OPEN aroperatorfile$ FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 256
 FIELD 2, 4 AS prec$, 8 AS pamt$, 15 AS pdesc$, 2 AS ptype$, 30 AS pname$, 4 AS TFDIST$, 4 AS TLDIST$
 LSET prec$ = MKS$(record)
 LSET pamt$ = MKD$(amount#)
 LSET pdesc$ = "ENV #" + STR$(ENVELOPE)
 LSET ptype$ = MKI$(PCODE)
 LSET pname$ = nme$
 nexrecord = LOF(2) / 256 + 1:
 PUT 2, nexrecord
REM determine catagory
OPEN "arcat.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #12 LEN = 64
FIELD 12, 30 AS catname$, 2 AS catnumber$
FOR ii = 1 TO LOF(12) / 64: GET 12, ii
IF CATEGORY$ = catnumber$ THEN CAT$ = catname$: CLOSE 12: GOTO 1200.4
NEXT ii
CLOSE 12
CAT$ = ""

1200.4 FOR LL = 12 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
COLOR 11
LOCATE 15, 1: PRINT "Amount Paid on Account ....... "; USING "#####,#.##"; amount#
              PRINT "Cash Received from Customer .. "; USING "#####,#.##"; amtreceived#
              PRINT "                               -------------"
              COLOR 15
              PRINT "CHANGE DUE CUSTOMER .......... "; USING "#####,#.##"; changedue#



     LPRINT "  T O W N   O F   T R O Y"
     LPRINT
     LPRINT "  A/R RECEIPT FOR PAYMENT"
     LPRINT "        "; DATE$
     LPRINT "    Env. #:"; ENVELOPE
     LPRINT "   "; CAT$
     LPRINT
     LPRINT "Customer: "; LEFT$(nme$, 18)
     LPRINT "  Acct #: "; acctrecord
     LPRINT
     LPRINT " Bill Amount: "; USING "$####,#.##"; amount#
     LPRINT " Amount Paid: "; USING "$####,#.##"; amtreceived#
     LPRINT "  Change Due: "; USING "$####,#.##"; changedue#
     LPRINT
     LPRINT "Acct Balance: "; USING "$####,#.##"; owed# - amount#
     LPRINT LEFT$(MESSAGE$, 29)
     LPRINT "    T H A N K   Y O U  !"
     FOR JJ = 1 TO 5
     LPRINT " "
     NEXT JJ
     LPRINT CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)
     CLOSE
     LOCATE 20, 1: COLOR 11
     PRINT "PRESS <ENTER> TO CONTINUE "; : fl = 1: GOSUB 62000
     GOTO 1000

1900.2 COLOR 11: LOCATE 19, 1: PRINT "Payment Codes Are As Follows: "
 COLOR 14: LOCATE 20, 1: PRINT "1= Reg Payment - CASH      2= Reg Payment - CHECK      3= ADJUSTMENT"
 RETURN


1950.2 IF PCODE = 1 THEN dd$ = "Reg Payment - CASH"
     IF PCODE = 2 THEN dd$ = "Reg Payment - CHECK"
     IF PCODE = 3 THEN dd$ = "ADJUSTMENT/CREDIT"
     LOCATE 19, 1: PRINT STRING$(79, 32)
     PRINT STRING$(79, 32): PRINT STRING$(79, 32)
     PRINT STRING$(79, 32); : RETURN
arnamesearch:
 LOCATE 3, 1:
 COLOR 11: PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : fl = 25
 GOSUB 62000: IF cf = 1 THEN ggo = 0: RETURN
 nn$ = in$
 nn = LEN(nn$)
 CLOSE 4: OPEN "arlname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24: FIELD 4, 20 AS lastnam$, 4 AS mrec$
 CLOSE 1: GOSUB 60000.2
 FOR LL = 1 TO LOF(4) / 24: GET 4, LL
 IF LEFT$(lastnam$, nn) = nn$ THEN GOTO arn1
arn3:
 NEXT LL
 ggo = 0
 RETURN
arn1:
 GET 1, CVS(mrec$)
 LOCATE 6, 1: COLOR 11
 arnme$ = nme$
 COLOR 10
 PRINT "Full Account Name: "; arnme$
 PRINT "   Account Number: "; CVS(mrec$)
 PRINT "         Catagory: "; CATEGORY$
 PRINT
arn2:
 LOCATE 9, 1: COLOR 11
 PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
 IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO arn2
 IF in$ = "Q" THEN ggo = 0: RETURN
 IF in$ = "Y" THEN ggo = 1: record = CVS(mrec$): RETURN
 IF in$ = "N" THEN GOTO arn3
 GOTO arn2
 

3000 RUN "pospay2"
4000 RUN "pospost"
5000 scren$ = "4.05.00": CLS
5020 LOCATE 1, 30: COLOR 11: PRINT "Set Message Line"
5030 LOCATE 3, 1: PRINT "Enter Message To Print (ESC-End) .. "; : fl = 40: GOSUB 62000
5040 OPEN "R", 8, "message.dat", 40: FIELD 8, 40 AS mes$
5050 GET 8, 1:
5060 LSET mes$ = in$
5070 PUT 8, 1: CLOSE 8: GOTO 120

6000 d$ = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 2): PDATE$ = d$
6010 scren$ = "5.06.00": CLS : GOSUB 63500
6020 COLOR 11: LOCATE 1, 30
     PRINT "Payment Entry Setup"
6030 LOCATE 3, 1: COLOR 11
6100 CLOSE 1: OPEN "R", 1, "tbmpaid.dat", 256
     FIELD 1, 1 AS chargecurrent$, 6 AS paydate$, 4 AS ENVELOPE$, 2 AS TAXYEAR$
6105 IF LOF(1) / 256 = 0 THEN LSET chargecurrent$ = "": LSET paydate$ = "": LSET ENVELOPE$ = MKS$(0): LSET TAXYEAR$ = MKI$(0): PUT 1, 1
6110 GET 1, 1
6200 FOR LL = 3 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
     LOCATE 3, 1: COLOR 11
     PRINT "1- Charge Current Year Interest (Y/N) .. "; chargecurrent$
     PRINT "2- Payment Entry Date Default (mmddyy) . "; paydate$
     PRINT "3- Envelope Number ..................... "; CVS(ENVELOPE$)
     PRINT "4- Tax Year to Enter (YY) .............. "; CVI(TAXYEAR$)
6210 LOCATE 8, 1: COLOR 11
     PRINT "Enter Item # to Change (1-4) or <ESC>-OK .. "; : fl = -1
     GOSUB 62000
     IF cf = 1 THEN CLOSE : GOTO 1000
     IF VAL(in$) < 1 OR VAL(in$) > 4 THEN 6210
     ON VAL(in$) GOSUB 6310, 6320, 6330, 6340
     GOTO 6210
6310  LOCATE 3, 42: fl = 1: GOSUB 62210: IF NOT (in$ = "Y" OR in$ = "N") THEN 6310
      LSET chargecurrent$ = in$: PUT 1, 1: RETURN
6320  LOCATE 4, 42: fl = -6: GOSUB 62000: IF LEN(in$) <> 6 THEN 6320
      LSET paydate$ = in$: PUT 1, 1: RETURN
6330  LOCATE 5, 43: fl = -5: GOSUB 62000
      LSET ENVELOPE$ = MKS$(VAL(in$)): PUT 1, 1: RETURN
6340  LOCATE 6, 43: fl = -2: GOSUB 62000: IF LEN(in$) <> 2 THEN 6340
      LSET TAXYEAR$ = MKI$(VAL(in$)): PUT 1, 1: RETURN
7000 scren$ = "4.06.00": CLS : GOSUB 63500
     LOCATE 1, 30: COLOR 11: PRINT "Operator Maintenance";
     LOCATE 3, 1: COLOR 11
     PRINT "(A)dd New Operator : (C)hange Operator Password : (P)rint Operator List .. "; : fl = 1
     GOSUB 62210: IF cf = 1 THEN 10
     IF NOT (in$ = "A" OR in$ = "C" OR in$ = "P") THEN 7000
     IF in$ = "A" THEN GOSUB addop
     IF in$ = "C" THEN GOSUB changeop
     IF in$ = "P" THEN GOSUB printop
     GOTO 7000
addop:
FOR LL = 3 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL: LOCATE 3, 1: COLOR 11
GOSUB opfile
PRINT "Operator ID# Assigned: "; LOF(8) / 64 + 1
PRINT
PRINT "Enter the Operator Name: "; : fl = 30: GOSUB 62000
LSET OPNAME$ = UCASE$(in$)
COLOR 11
PRINT "Enter the Operator Password (Must by 5 characters): "; : fl = 5
GOSUB 63000
passwd$ = in$
COLOR 11
PRINT "Please Verify by"
PRINT "Enter the Operator Password (Must by 5 characters): "; : fl = 5
GOSUB 63000
IF in$ <> passwd$ THEN
 COLOR 11
 PRINT "Passwords did not match : please begin again!"
 PRINT "Press <ENTER> to Continue "; : fl = 1: GOSUB 62000: CLOSE : GOTO addop
END IF
LSET oppassword$ = in$
LSET blk$ = ""
PUT 8, LOF(8) / 64 + 1
COLOR 11
PRINT "Operator Entered Successfully": PRINT CHR$(7)
RETURN

changeop:
FOR LL = 3 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL: LOCATE 3, 1: COLOR 11
GOSUB opfile
COLOR 11
PRINT "Enter the Operator ID #: "; : fl = -2: GOSUB 62000
opnumber = VAL(in$)
IF opnumber < 1 OR opnumber > LOF(8) / 64 THEN CLOSE : GOTO changeop
GET 8, opnumber: COLOR 11: PRINT "Operator Name: "; OPNAME$: PRINT

PRINT "Enter the Operator Password (Must by 5 characters): "; : fl = 5
GOSUB 63000
passwd$ = in$
COLOR 11
PRINT "Please Verify by"
PRINT "Enter the Operator Password (Must by 5 characters): "; : fl = 5
GOSUB 63000
IF in$ <> passwd$ THEN
 COLOR 11
 PRINT "Passwords did not match : please begin again!"
 PRINT "Press <ENTER> to Continue "; : fl = 1: GOSUB 62000: CLOSE : GOTO changeop
END IF
LSET oppassword$ = in$
LSET blk$ = ""
PUT 8, opnumber
COLOR 11
PRINT "Operator Changed Successfully": PRINT CHR$(7)
RETURN

printop:
FOR LL = 3 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL: LOCATE 3, 1: COLOR 11
GOSUB opfile
LOCATE 5, 1: PRINT "PRINT OPERATOR FILE (Y/N) .. "; : fl = 1: GOSUB 62210
IF in$ = "Y" THEN
 LPRINT "CASH RECEIPTS SYSTEM : VALID OPERATOR FILE : DATE: "; DATE$
 LPRINT
 LPRINT "ID#    OPERATOR NAME"
 LPRINT "---    -------------"
 FOR x = 1 TO LOF(8) / 64: GET 8, x
 LPRINT USING "##"; x;
  LPRINT TAB(8); OPNAME$
 NEXT x
 LPRINT CHR$(12);
 CLOSE
END IF
CLOSE
RETURN

8000 scren$ = "4.06.00": CLS : GOSUB 63500
     LOCATE 1, 30: COLOR 11: PRINT "Misc Code Maintenace"
     COLOR 11
     LOCATE 3, 1: PRINT "(C)hange or (P)rint Codes (F2-Init:ESC-Menu) .. "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN 120
     IF cf = 12 THEN GOSUB codeinit: GOTO 8000
     IF NOT (in$ = "C" OR in$ = "P") THEN 8000
     IF in$ = "P" THEN 8500
     GOSUB misccodefile
8010 LOCATE 3, 1: COLOR 11: PRINT STRING$(79, 32): LOCATE 3, 1
     PRINT "Enter Code Number (1-12) or <ESC>-Menu .. "; : fl = -2
     GOSUB 62000: IF cf = 1 THEN 8000
     XX = VAL(in$)
     IF XX < 1 OR XX > 12 THEN 8010
     GET 15, XX
     LOCATE 6, 1
     COLOR 11
     PRINT "Enter Code Desc.  .. "; : fl = 30: GOSUB 62000: LSET CODEDESC$ = in$: LSET active$ = "Y": PUT 15, XX
     PRINT "Updated"
     PRINT CHR$(7);
     LOCATE 6, 1: PRINT STRING$(79, 32): LOCATE 7, 1: PRINT STRING$(79, 32): LOCATE 8, 1: PRINT STRING$(79, 32)
     GOTO 8010

8500 REM print them
     GOSUB misccodefile
     FOR x = 1 TO 12: GET 15, x
     IF active$ = "Y" THEN
      PRINT "Code # "; x; " Set To "; CODEDESC$
     END IF
     NEXT x
     LOCATE 23, 1: PRINT "End of Codes"
     PRINT "Press ENTER to Continue "; : fl = 1
     GOSUB 62000
     CLOSE
     GOTO 8000


codeinit:
GOSUB misccodefile
FOR x = 1 TO 12: GET 15, x: LSET active$ = "N": LSET CODEDESC$ = "": PUT 15, x: NEXT x
COLOR 11
PRINT "Initialized"
PRINT CHR$(7);
CLOSE 15: RETURN

misccodefile:
CLOSE 15
OPEN "R", 15, "posmisc.dat", 256
FIELD 15, 30 AS CODEDESC$, 1 AS active$, 225 AS nu$
RETURN

9000 scren$ = "4.07.00": CLS : GOSUB 63500
     LOCATE 1, 30: COLOR 11: PRINT "Set Envelope #"
     COLOR 11
     LOCATE 3, 1: PRINT "Enter the Envelope Number (ESC-Menu) .. "; : fl = -7
     GOSUB 62000: IF cf = 1 THEN 10
     ENV = VAL(in$): ENVELOPE = ENV
     OPEN "poenv.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #19 LEN = 128
     FIELD 19, 4 AS ENVELOPE$, 124 AS blank$
     LSET ENVELOPE$ = MKS$(ENV): LSET blank$ = ""
     PUT 19, 1
     CLOSE 19: GOTO 120

setenv:
     OPEN "poenv.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #19 LEN = 128
     FIELD 19, 4 AS ENVELOPE$, 124 AS blank$
     IF LOF(19) > 0 THEN GET 19, 1: ENVELOPE = CVS(ENVELOPE$): CLOSE 19: setenv = 1: RETURN
     CLOSE 19
     scren$ = "4.07.00": CLS : GOSUB 63500
     LOCATE 1, 30: COLOR 11: PRINT "Set Envelope #"
     COLOR 11
     LOCATE 3, 1: PRINT "Enter the Envelope Number (ESC-Menu) .. "; : fl = -7
     GOSUB 62000: IF cf = 1 THEN 10
     ENV = VAL(in$)
     OPEN "poenv.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #19 LEN = 128
     FIELD 19, 4 AS ENVELOPE$, 124 AS blank$
     LSET ENVELOPE$ = MKS$(ENV): LSET blank$ = ""
     PUT 19, 1
     CLOSE 19: setenv = 1: RETURN




tempfile:
TOPERATORFILE$ = "tbpay" + filen$ + ".dat"
TOPERATORFILE1$ = "tbpay1" + filen$ + ".dat"
                       

CLOSE 8: CLOSE 9
OPEN TOPERATORFILE$ FOR RANDOM ACCESS READ WRITE SHARED AS #8 LEN = 128
 FIELD 8, 6 AS PDATE$, 4 AS pamt$, 4 AS PTAXYR$, 20 AS PDESC1$, 2 AS ptype$, 4 AS Ptax$, 4 AS Pint$, 4 AS Padv$, 4 AS Pcol$, 4 AS plate$, 4 AS prec$, 4 AS Pctx$, 4 AS Pcin$, 4 AS Pcad$, 4 AS Pcco$, 4 AS pcla$, 4 AS pdisamt$, 6 AS ppaydate$, 2 AS  _
penvnum$
 nexrecord = LOF(8) / 128 + 1
OPEN TOPERATORFILE1$ FOR RANDOM ACCESS READ WRITE SHARED AS #9 LEN = 256
 FOR tloop = 0 TO 9
 FIELD 9, (24 * tloop) AS dummy4$, 4 AS ptx$(tloop), 4 AS pin$(tloop), 4 AS pad$(tloop), 4 AS pco$(tloop), 4 AS pla$(tloop): NEXT tloop
RETURN

opfile:
CLOSE 8: OPEN "posoper.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #8 LEN = 64
         FIELD 8, 30 AS OPNAME$, 5 AS oppassword$, 29 AS blk$
         RETURN
60000 CLOSE 1: OPEN "wbcust.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 256
60010 FIELD 1, 25 AS lname$, 25 AS firstname$, 30 AS add1$, 30 AS add2$, 20 AS city$, 2 AS st$, 9 AS zp$, 12 AS PHONE$, 2 AS BOOK$, 12 AS meter$, 2 AS RATE$, 8 AS WON$, 8 AS MINSTALL$, 8 AS AVGUSE$, 2 AS AVGCT$, 8 AS curbal$, 8 AS pastdue$, 1 AS  _
cycle$, 5 AS NOTUSED$, 8 AS CURREAD$, 8 AS PASTREAD$, 8 AS firsttr$, 8 AS lasttr$, 4 AS waterdep$
60015 CLOSE 8: OPEN "wbcust2.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #8 LEN = 64
      FIELD 8, 8 AS IREVENUE1$, 8 AS IREVENUE2$, 8 AS IREVENUE3$, 8 AS irevenue4$, 8 AS irevenue5$, 8 AS IREVENUE6$, 8 AS IREVENUE7$, 8 AS IREVENUE8$
60030 RETURN
61000 CLOSE 2: OPEN "wbartran.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 50
61010 FIELD 2, 8 AS TDATE$, 8 AS tamt$, 2 AS TTYPE$, 15 AS TDESC$, 4 AS tmrec$, 4 AS tnexttr$, 4 AS tcuramt$
61020 RETURN

60000.1 CLOSE 1: OPEN "tbcust.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 512
      FIELD 1, 7 AS acct$, 5 AS pin$, 174 AS filler$
      FIELD 1, 186 AS dummy$, 39 AS pdesc$, 22 AS pmap$, 35 AS psub$, 4 AS plot$, 35 AS pcounty$, 4 AS psize$, 8 AS pland$, 8 AS pbldg$
      FIELD 1, 341 AS dummy1$, 8 AS pbus$, 8 AS pres$, 8 AS ppub$, 4 AS persexp$, 4 AS scexp$, 4 AS homeexp$, 4 AS useexp$
      FIELD 1, 381 AS dummy2$, 4 AS curlateamt$, 4 AS curtaxamt$, 4 AS curintamt$, 4 AS curadvamt$, 4 AS curcolamt$, 4 AS ctaxnotice$
      FIELD 1, 405 AS dummy3$, 4 AS ftr$, 4 AS ltr$, 11 AS filler1$, 1 AS LATE$, 4 AS curprinciple$
      FIELD 1, 429 AS dummy4$, 4 AS CUSTOMERRECORD$, 4 AS nexparcel$, 75 AS blank$
      CLOSE 3: OPEN "tbcust1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #3 LEN = 280
      FOR tloop = 0 TO 9
      FIELD 3, (24 * tloop) AS dummy4$, 4 AS ptaxamt$(tloop), 4 AS pintamt$(tloop), 4 AS padvamt$(tloop), 4 AS pcolamt$(tloop), 4 AS plateamt$(tloop), 4 AS ptaxnotice$(tloop): NEXT tloop
      FOR tloop = 0 TO 9
      FIELD 3, (240 + (4 * tloop)) AS dummy5$, 4 AS pprinciple$(tloop)
      NEXT tloop
      CLOSE 14: OPEN "tbcust2.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = 256
      FIELD 14, 7 AS acct1$, 25 AS firstnam$, 35 AS lnam$, 35 AS A1$, 35 AS A2$, 20 AS cy$, 2 AS st$, 10 AS zp$, 12 AS tp$, 4 AS firstparcel$, 4 AS LASTPARCEL$, 11 AS ss$
      RETURN

61000.1 CLOSE 2: OPEN "tbartran.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 96
61010.1 FIELD 2, 6 AS TDATE$, 4 AS tamt$, 2 AS ttaxyr$, 20 AS TDESC$, 2 AS TTYPE$, 4 AS taprin$, 4 AS taint$, 4 AS taadv$, 4 AS tacol$, 4 AS talate$, 4 AS tmrec$, 4 AS tnexttr$, 4 AS tbal$, 4 AS tdisamt$, 2 AS tenvnum$, 24 AS tblank$
61020.1 RETURN

60000.2 CLOSE 1: OPEN "arcust.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 512
60010.2 FIELD 1, 40 AS nme$, 30 AS add1$, 30 AS add2$, 20 AS city$, 2 AS st$
      FIELD 1, 122 AS du1$, 9 AS zp$, 12 AS PHONE$, 8 AS curbal$, 8 AS pastdue$
      FIELD 1, 159 AS du2$, 1 AS active$, 4 AS firsttr$, 4 AS lasttr$
      FIELD 1, 168 AS du3$, 30 AS contact$, 2 AS CATEGORY$, 4 AS stfee$
      FIELD 1, 204 AS du4$, 12 AS license$, 6 AS validdate$, 290 AS blank$
      RETURN
61000.2 CLOSE 2: OPEN "artran.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 128
      FIELD 2, 8 AS TDATE$, 4 AS tcuramt$, 4 AS TrunAMT$, 2 AS TTYPE$, 25 AS TDESC$
      FIELD 2, 43 AS dum1$, 4 AS tmrec$, 4 AS tnexttr$, 77 AS blank1$
      RETURN


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
62140 IF z$ = "." AND INSTR(in$, ".") = 0 THEN GOTO 62180
62150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF z$ < "0" OR z$ > "9" THEN 62030
62180 PRINT z$; : in$ = in$ + z$: ZL = ZL + 1
62190 IF ABS(fl) = 1 THEN 62050 ELSE 62030
62200 LOCATE RO, PO: PRINT in$: RETURN
62210 GOSUB 62000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
62220 RETURN
63000 cf = 0: in$ = "": COLOR 14: RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
63010 IF INKEY$ <> "" THEN 63010
63020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
63030 z$ = INKEY$: IF z$ = "" THEN 63030
63040 IF z$ <> CHR$(13) THEN 63060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
63050 GOSUB 63200: RETURN
63060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
63065 IF ASC(z$) = 0 THEN 63030
63070 IF LEN(z$) > 1 THEN cf = ASC(RIGHT$(z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
63080 IF z$ = " " THEN 63130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 63000
63090 IF z$ <> CHR$(8) THEN 63130 ELSE IF ZL = 0 THEN 63030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 63110 ELSE IF z$ = "," THEN 63120
63100 IF z$ = "." THEN ZD = 0: GOTO 63110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
63110 in$ = LEFT$(in$, LEN(in$) - 1)
63120 ZL = ZL - 1: GOTO 63030
63130 IF ABS(fl) = ZL THEN 63030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 63180
63140 IF z$ = "." AND ZD = 0 THEN ZD = 1: GOTO 63180
63150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 63190
63160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 63180
63170 IF z$ < "0" OR z$ > "9" THEN 63030
63180 PRINT "*"; : in$ = in$ + z$: ZL = ZL + 1
63190 IF ABS(fl) = 1 THEN 63050 ELSE 63030
63200 PRINT "": RETURN

63210 GOSUB 63000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
63220 RETURN

63500 LOCATE 1, 1: COLOR 10: PRINT "Post Date: "; sysdate$; TAB(70); scren$: PRINT STRING$(79, "_"): RETURN
fixamts:
fix1:

 LOCATE 15, 1: COLOR 11: PRINT "Enter Item Number to Edit (ESC-When Finished) .. "; : fl = -1: GOSUB 62000
 IF cf = 1 THEN LOCATE 15, 1: PRINT STRING$(79, 32): RETURN
 IF VAL(in$) < 1 OR VAL(in$) > 6 THEN GOTO fix1
 IF VAL(in$) = 1 THEN LOCATE 6, 63: fl = -10: GOSUB 62000: LSET IREVENUE1$ = MKD$(VAL(in$)): PUT 8, acctrecord: GOTO fix1
 IF VAL(in$) = 2 THEN LOCATE 7, 63: fl = -10: GOSUB 62000: LSET IREVENUE2$ = MKD$(VAL(in$)): PUT 8, acctrecord: GOTO fix1
 IF VAL(in$) = 3 THEN LOCATE 8, 63: fl = -10: GOSUB 62000: LSET IREVENUE3$ = MKD$(VAL(in$)): PUT 8, acctrecord: GOTO fix1
 IF VAL(in$) = 4 THEN LOCATE 9, 63: fl = -10: GOSUB 62000: LSET irevenue4$ = MKD$(VAL(in$)): PUT 8, acctrecord: GOTO fix1
 IF VAL(in$) = 5 THEN LOCATE 10, 63: fl = -10: GOSUB 62000: LSET irevenue5$ = MKD$(VAL(in$)): PUT 8, acctrecord: GOTO fix1
 IF VAL(in$) = 6 THEN LOCATE 11, 63: fl = -10: GOSUB 62000: LSET IREVENUE6$ = MKD$(VAL(in$)): PUT 8, acctrecord:  GOTO fix1
networkerror:
     REM ERROR CODES
     Number = ERR
     SELECT CASE ERR
          CASE IS = 70
               LOCATE 23, 1: PRINT STRING$(79, 32): LOCATE 23, 1: COLOR 28: PRINT "Network Message:"; : COLOR 15
               PRINT "This RECORD is Busy : Wait? (Y/N) ..."; : fl = 1: GOSUB 62210
               SELECT CASE in$
                    CASE IS = "Y"
                         start = TIMER
                         DO
                              time = TIMER - start
                         LOOP WHILE time < 10
                         RESUME
                    CASE IS = "N"
                         RESUME NEXT
                    CASE ELSE
                         GOSUB 62210
               END SELECT

          CASE IS = 24, 25
               LOCATE 23, 1: PRINT STRING$(79, 32): LOCATE 23, 1: COLOR 28
               PRINT "Error Message:"; : COLOR 15: PRINT "Device Time-out! Put Printer On-Line: Press ENTER to Resume Printing";
               fl = 1: GOSUB 62000: LOCATE 23, 1: PRINT STRING$(79, 32)
               RESUME
          CASE IS = 27
               LOCATE 23, 1: PRINT STRING$(79, 32): LOCATE 23, 1: COLOR 28
               PRINT "Error Message:"; : COLOR 15: PRINT "Printer Is OUT OF PAPER! Press ENTER to Resume Printing";
               fl = 1: GOSUB 62000: LOCATE 23, 1: PRINT STRING$(79, 32)
               RESUME
          CASE ELSE
               CLS
               LOCATE 15, 1: COLOR 28: PRINT "Error Message: "; : COLOR 15: PRINT "Error "; Number; " Has Occured in Line "; ERL
               PRINT "Write Down Message and Hit <ESC> Key to exit."
               fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10 ELSE GOSUB 62000
     END SELECT

SETOPERATOR:
COLOR 11
LOCATE 1, 34: PRINT "Enter Payments"
   GOSUB opfile
IF LOF(8) = 0 THEN
 COLOR 11
 PRINT "No Operator File FOUND"; CHR$(7);
 PRINT "The operator file must be set up : Press <ENTER> to Continue "; : fl = 1
 GOSUB 62000
 GOTO 7000
END IF
LOCATE 3, 1: COLOR 11
PRINT "Enter Your Operator Number Now .. "; : fl = -2
GOSUB 62000
IF cf = 1 THEN RUN "CPMENU"
record = VAL(in$)
IF record < 1 OR record > LOF(8) / 64 THEN CLOSE : GOTO 120
GET 8, record
COLOR 11
PRINT "Enter Your Password Now .. "; : fl = 5: GOSUB 63000
IF in$ = oppassword$ THEN
  OPERATORNUMBER = record
  operatorname$ = OPNAME$
  filen$ = STR$(OPERATORNUMBER)
  filen$ = RIGHT$(filen$, LEN(filen$) - 1)
  OPERATORFILE$ = "pospay" + filen$ + ".dat"
  CLOSE
 ELSE
  PRINT "Invalid Password : Please Begin Again"
  PRINT "Press <ENTER> to Continue "; : fl = 1: GOSUB 62000: CLOSE : GOTO 10
END IF
scren$ = "4.00.00": sysdate$ = LEFT$(DATE$, 6) + RIGHT$(DATE$, 2)
CLS
GOSUB 63500
OPSET = 1
RETURN


Tnamesearch:
LOCATE 3, 1:
FOR LL = 3 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
LOCATE 3, 1
COLOR 11: PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : fl = 25
 GOSUB 62000: IF cf = 1 THEN ggo = 0: RETURN
 nn$ = in$
 nn = LEN(nn$)
 IF nn > 8 THEN nn = 8
 CLOSE 4: OPEN "tblname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24: FIELD 4, 20 AS lastnam$, 4 AS mrec$
 CLOSE 1: GOSUB 60000.1

FOR LL = 1 TO LOF(4) / 24: GET 4, LL
IF LEFT$(lastnam$, nn) = LEFT$(nn$, nn) THEN GOTO n1
Nn3:
NEXT LL
ggo = 0: RETURN
n1:
GET 14, CVS(mrec$)
LOCATE 6, 1: COLOR 11
kk = INSTR(firstnam$, "  ")
IF kk = 0 THEN kk = 25
IF kk = 1 THEN nme$ = lnam$ ELSE nme$ = LEFT$(firstnam$, kk) + lnam$
COLOR 10
FOR lll = 6 TO 9: LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
LOCATE 6, 1
PRINT "Full Account Name: "; nme$
PRINT "   Account Number: "; acct1$
ACCTNUMBER$ = acct1$
account$ = acct1$
PRINT
Nn2:
LOCATE 9, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO Nn2
IF in$ = "Q" THEN ggo = 0: RETURN
IF in$ = "Y" THEN ggo = 1: record = CVS(mrec$): RETURN
IF in$ = "N" THEN GOTO Nn3
GOTO Nn2

findacct:
     CLOSE 4
     OPEN "TBCUSTAC.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 16
     FIELD 4, 12 AS met$, 4 AS ARREC$
     IF LEN(account$) < 12 THEN account$ = account$ + STRING$(12 - LEN(account$), 32)
     IF LOF(4) / 16 = 0 THEN RETURN
     HI = LOF(4) / 16: LO = 1: ME = HI
     FOR XX = 1 TO 16: GET 4, ME
     IF account$ = met$ THEN ng = 1: record = CVS(ARREC$): CLOSE 4: RETURN
     IF account$ > met$ THEN LO = ME ELSE HI = ME
     ME = INT((HI - LO) / 2) + LO
     NEXT XX: CLOSE 4
     CLOSE 4
     ng = 0
     RETURN

