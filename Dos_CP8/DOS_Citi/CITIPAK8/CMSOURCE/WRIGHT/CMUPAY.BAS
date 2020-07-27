    DEFINT A-Z
'*** Citipak Include Files Needed ***
'$INCLUDE: 'CMFILES.BI'
'$INCLUDE: 'ubtrans.BI'
'$INCLUDE: 'gl.bi'

    OPEN "CMUTIL.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #15
    FIELD 15, 2 AS OperNum$, 10 AS PDate$
    GET 15, 1
    OperRecNumber = CVI(OperNum$)
    PostDate$ = PDate$
    DIM translink!(700), totm@(299), messline$(20), m$(20), Apply!(10)
    Scren$ = "4.00.00"
    SysDate$ = DATE$

10 CLS
    GOSUB 63500
    GOSUB OldSetOperatorName:

1000 REM
     FOR ll = 3 TO 23: LOCATE ll, 1: PRINT STRING$(79, 32): NEXT ll
     FOR JJ = 1 TO 8: Apply!(JJ) = 0: NEXT JJ
     pdesc$ = ""
     GOSUB openfile
     GET 14, 1
     max = CVI(REVTRACK$)
     CLOSE 14
1020 LOCATE 1, 30: COLOR 11: PRINT "Utility Payment Entry"; : LOCATE 3, 1
1025 DEPFLAG = 0
1030 COLOR 11: PRINT "Enter Acct # Please (<TAB>-Name Search : <ESC>-End) .. "; : fl = -5
     GOSUB 62000
     IF cf = 1 THEN RUN "CMMENU"
     IF cf = 2 THEN GOSUB namesearch: IF GGO = 1 THEN GOTO 1080 ELSE CLOSE : GOTO 1000

1040 Record = VAL(in$)
1080 CLOSE 1: GOSUB 60000
     IF Record < 1 OR Record > LOF(1) / 1024 THEN CLOSE : GOTO 1000
      GET 1, Record: AcctRecord! = Record
      GOSUB openfile: GET 14, 1
     IF VAL(LEFT$(meter$, 2)) = 0 THEN
      LOCATE 23, 1: COLOR 11: PRINT "Account Deleted : Press <ENTER> to Continue "; : fl = 1
      GOSUB 62000: CLOSE : GOTO 1000
     END IF
1090 LOCATE 3, 1: PRINT STRING$(79, 32)
     B = INSTR(meter$, "/"): IF B = 0 THEN m$ = meter$: GOTO 1095
     m$ = LEFT$(meter$, B - 1)
1095 COLOR 15: LOCATE 3, 1
      PRINT "  Account Number: "; USING "#####"; Record;
     LSET IRev1$ = MKD$((INT((CVD(IRev1$) * 100) + .5) / 100))
     LSET IRev2$ = MKD$((INT((CVD(IRev2$) * 100) + .5) / 100))
     LSET IRev3$ = MKD$((INT((CVD(IRev3$) * 100) + .5) / 100))
     LSET IRev4$ = MKD$((INT((CVD(IRev4$) * 100) + .5) / 100))
     LSET IRev5$ = MKD$((INT((CVD(IRev5$) * 100) + .5) / 100))
     LSET IRev6$ = MKD$((INT((CVD(IRev6$) * 100) + .5) / 100))
     LSET irev7$ = MKD$((INT((CVD(irev7$) * 100) + .5) / 100))
     LSET IREV8$ = MKD$((INT((CVD(IREV8$) * 100) + .5) / 100))
     PUT 1, Record
     LOCATE 3, 45
     PRINT "( CURRENT: "; USING "####,#.##"; CVD(CURBAL$); : PRINT ")"
     LOCATE 4, 1: PRINT "Account Location: "; m$;
     LOCATE 4, 45
     PRINT "(PAST DUE: "; USING "####,#.##"; CVD(pastdue$); : PRINT ")"
     LOCATE 5, 1
1100 IF CVS(messfirst$) > 0 THEN
      GOSUB reviewmessage
     END IF
     COLOR 11: PRINT "Customer Account of ..."; TAB(45); "Balance = "; USING "####.##"; CVD(CURBAL$) + CVD(pastdue$): owed# = CVD(CURBAL$) + CVD(pastdue$): owed# = INT((owed# * 100) + .5) / 100
1110 COLOR 14
     kk = INSTR(firstname$, "  "): IF kk = 0 THEN kk = 25
     IF LEFT$(firstname$, 1) = " " THEN
       nme$ = lname$
      ELSE
       nme$ = LEFT$(firstname$, kk) + lname$
     END IF
1120 IF messageflag = 1 THEN COLOR 27 ELSE COLOR 14
1135 PRINT TAB(5); LEFT$(nme$, 39);
     COLOR 14: PRINT TAB(45); "1-"; LEFT$(rev1name$, 15); USING "$####,#.##"; CVD(IRev1$)
1140 PRINT TAB(5); add1$; : IF max < 2 THEN PRINT  ELSE PRINT TAB(45); "2-"; LEFT$(rev2name$, 15); USING "$####,#.##"; CVD(IRev2$)
1150 PRINT TAB(5); add2$; : IF max < 3 THEN PRINT  ELSE PRINT TAB(45); "3-"; LEFT$(rev3name$, 15); USING "$####,#.##"; CVD(IRev3$)
1160 PRINT TAB(5); CITY$; " "; ST$; " "; ZP$; : IF max < 4 THEN PRINT  ELSE PRINT TAB(45); "4-"; LEFT$(rev4name$, 15); USING "$####,#.##"; CVD(IRev4$)
     IF max < 5 THEN 1165 ELSE PRINT TAB(45); "5-"; LEFT$(rev5name$, 15); USING "$####,#.##"; CVD(IRev5$)
     IF max < 6 THEN 1165 ELSE PRINT TAB(45); "6-"; LEFT$(rev6name$, 15); USING "$####,#.##"; CVD(IRev6$)
     IF max < 7 THEN 1165 ELSE PRINT TAB(45); "7-"; LEFT$(rev7name$, 15); USING "$####,#.##"; CVD(irev7$)
     IF max < 8 THEN 1165 ELSE PRINT TAB(45); "8-"; LEFT$(rev8name$, 15); USING "$####,#.##"; CVD(IREV8$)
1165 TOTREV# = CVD(IRev1$) + CVD(IRev2$) + CVD(IRev3$) + CVD(IRev4$) + CVD(IRev5$) + CVD(IRev6$) + CVD(irev7$) + CVD(IREV8$)
     TOTREV# = INT((TOTREV# * 100) + .5) / 100
1170 LOCATE 12, 1: COLOR 11: PRINT STRING$(44, "-")
1180 LOCATE 13, 1: COLOR 11: PRINT "Correct Acct (Y/N)(<F6>-Cash : <F7>-Check)";
     LOCATE 14, 1: COLOR 11: PRINT "(F4-Messages)(F5-History)(<TAB>-EDIT AMT):"; : fl = 1: GOSUB 62210:
     IF cf = 18 THEN
       LOCATE 19, 1
       PRINT "Enter New Cur Bal Amt .. "; : fl = -9: GOSUB 62000
       LSET CURBAL$ = MKD$(VAL(in$)): PUT 1, AcctRecord!
       LOCATE 19, 1: PRINT STRING$(79, 32)
       GOTO 1090
     END IF


     IF cf = 2 THEN GOSUB fixamts: GOTO 1090
     IF cf = 14 THEN
      GOSUB messagecenter
      Scren$ = "4.01.00": CLS : GOSUB 63500
      LOCATE 1, 34: COLOR 11: PRINT "Payment Entry"; : LOCATE 3, 1
      GOTO 1090
     END IF

     IF cf = 15 THEN
      GOSUB displayhistory
     Scren$ = "4.01.00": CLS : GOSUB 63500
     LOCATE 1, 34: COLOR 11: PRINT "Payment Entry"; : LOCATE 3, 1

      GOTO 1090
     END IF
     IF cf = 16 THEN amount@ = owed#: pcode = 1: GOTO 1285
     IF cf = 17 THEN
         amount@ = owed#: pcode = 2:
         LOCATE 15, 1: COLOR 11: PRINT "Enter Description (Ck #) ............ "; : fl = 15: GOSUB 62000: pdesc$ = in$
         GOTO 1285
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
     FOR JJ = 16 TO 23: LOCATE JJ, 1: PRINT STRING$(79, 32): NEXT JJ
     GOTO 1170
1205 LOCATE 13, 1: PRINT STRING$(44, 32)
     LOCATE 14, 1: PRINT STRING$(44, 32)
1210 IF CVS(waterdep$) > 0 THEN COLOR 10: LOCATE 23, 1: PRINT "Deposit on File in the Amount of :"; USING "###.##"; CVS(waterdep$);
1220 GOSUB 1900: LOCATE 13, 1: COLOR 11
     PRINT "Enter Payment Code (See Below 1-7) .. ";
     fl = -1: GOSUB 62000
     IF cf = 1 THEN CLOSE : GOTO 1000 ELSE pcode = VAL(in$)
     IF pcode < 1 OR pcode > 7 THEN 1220
1222 GOSUB 1950: COLOR 11: LOCATE 14, 55: PRINT dd$
1250 IF pcode = 3 THEN 1260
     LOCATE 14, 1: COLOR 11: PRINT "Enter Amount ... "; : PRINT "("; USING "$####.##"; owed#; : PRINT ") ......... "; : fl = -8: GOSUB 62000
     IF LEN(in$) = 0 THEN amount@ = owed#: LOCATE 14, 39: PRINT USING "$####.##"; amount@:  ELSE amount@ = VAL(in$): LOCATE 14, 39: PRINT USING "$####.##"; amount@
     GOTO 1260.0
1260 amount@ = CVS(waterdep$): amount@ = INT((amount@ * 100) + .5) / 100
1260.0 IF pcode = 4 OR pcode = 5 THEN
1260.1 LOCATE 15, 1: COLOR 11: PRINT "Enter Deposit Date (mmddyy) ......... "; : fl = 6: GOSUB 62000
       IF LEN(in$) <> 6 THEN GOTO 1260.1
       des$ = in$
       ELSE
1261 LOCATE 15, 1: COLOR 11: PRINT "Enter Description (Ck #) ............ "; : fl = 15: GOSUB 62000: pdesc$ = in$
     END IF
1265 LOCATE 16, 1: COLOR 11: PRINT "Enter Amount Tendered ............... "; : fl = -9: GOSUB 62000
     tendered@ = VAL(in$)
     LOCATE 16, 39:  PRINT USING "$####.##"; tendered@

     LOCATE 16, 55: COLOR 11
     PRINT "Change Due: "; : COLOR 15: PRINT USING "####.##"; tendered@ - amount@
1270 LOCATE 17, 1: COLOR 11: PRINT "Is the Above Correct (Y/N) .......... "; : fl = 1: GOSUB 62210: IF NOT (in$ = "Y" OR in$ = "N") THEN 1270 ELSE LOCATE 17, 39: PRINT in$
1280 IF in$ = "N" THEN
     FOR j = 13 TO 14: LOCATE j, 1: PRINT STRING$(44, 32): NEXT j
     FOR j = 14 TO 18: LOCATE j, 1: PRINT STRING$(79, 32): NEXT j
     GOTO 1200
     END IF
1285 IF amount@ = owed# THEN Apply!(1) = CVD(IRev1$): Apply!(2) = CVD(IRev2$): Apply!(3) = CVD(IRev3$): Apply!(4) = CVD(IRev4$): Apply!(5) = CVD(IRev5$): Apply!(6) = CVD(IRev6$): Apply!(7) = CVD(irev7$): Apply!(8) = CVD(IREV8$): GOTO 1295
     left@ = amount@
     IF pcode = 4 OR pcode = 5 THEN 1295
1286 FOR j = 12 TO 13: LOCATE j, 1: PRINT STRING$(44, 32): NEXT j
     FOR j = 14 TO 23: LOCATE j, 1: PRINT STRING$(79, 32): NEXT j
     GOSUB openfile
     GET 14, 1
     max = CVI(REVTRACK$)

     LOCATE 12, 1: COLOR 15: PRINT "Please Break Down the Payment"
     COLOR 10: LOCATE 13, 1: PRINT "Amount Paid by Customer: "; USING "$####.##"; amount@
     COLOR 11: LOCATE 14, 1: PRINT "1- Amount to Apply :"; LEFT$(rev1name$, 15)
      IF max < 2 THEN 1288 ELSE LOCATE 15, 1: PRINT "2- Amount to Apply :"; LEFT$(rev2name$, 15)
      IF max < 3 THEN 1288 ELSE LOCATE 16, 1: PRINT "3- Amount to Apply :"; LEFT$(rev3name$, 15)
      IF max < 4 THEN 1288 ELSE LOCATE 17, 1: PRINT "4- Amount to Apply :"; LEFT$(rev4name$, 15)
      IF max < 5 THEN 1288 ELSE LOCATE 18, 1: PRINT "5- Amount to Apply :"; LEFT$(rev5name$, 15)
      IF max < 6 THEN 1288 ELSE LOCATE 19, 1: PRINT "6- Amount to Apply :"; LEFT$(rev6name$, 15)
      IF max < 7 THEN 1288 ELSE LOCATE 20, 1: PRINT "7- Amount to Apply :"; LEFT$(rev7name$, 15)
      IF max < 8 THEN 1288 ELSE LOCATE 21, 1: PRINT "8- Amount to Apply :"; LEFT$(rev8name$, 15)

1288 GOSUB a1
     IF amount@ - left@ = 0 THEN LOCATE 23, 1: COLOR 13: PRINT "Amount has been fully distributed " ELSE LOCATE 23, 1: PRINT STRING$(65, 32)
     LOCATE 23, 1: COLOR 11: PRINT "Selection # (ESC-When Done) ... "; : fl = -1: GOSUB 62000
     IF cf = 1 THEN 1291
     LOCATE 16, 50: PRINT STRING$(28, 32)
     IF VAL(in$) < 1 OR VAL(in$) > 8 THEN 1288
     ON VAL(in$) GOSUB a2, a3, a4, a5, a6, a7, a8, a9: GOTO 1288
a1:
left@ = 0: FOR JJ = 1 TO 8
left@ = left@ + Apply!(JJ)
NEXT JJ: left@ = INT((left@ * 100) + .5) / 100
LOCATE 15, 50: COLOR 10: PRINT "Amount Left to Dist. "; USING "$####.##"; amount@ - left@: RETURN

a2: LOCATE 14, 36: fl = -9: GOSUB 62000: Apply!(1) = VAL(in$): LOCATE 14, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a3: LOCATE 15, 36: fl = -9: GOSUB 62000: Apply!(2) = VAL(in$): LOCATE 15, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a4: LOCATE 16, 36: fl = -9: GOSUB 62000: Apply!(3) = VAL(in$): LOCATE 16, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a5: LOCATE 17, 36: fl = -9: GOSUB 62000: Apply!(4) = VAL(in$): LOCATE 17, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a6: LOCATE 18, 36: fl = -9: GOSUB 62000: Apply!(5) = VAL(in$): LOCATE 18, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a7: LOCATE 19, 36: fl = -9: GOSUB 62000: Apply!(6) = VAL(in$): LOCATE 19, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a8: LOCATE 20, 36: fl = -9: GOSUB 62000: Apply!(7) = VAL(in$): LOCATE 20, 36: PRINT USING "$####.##"; VAL(in$): RETURN
a9: LOCATE 21, 36: fl = -9: GOSUB 62000: Apply!(8) = VAL(in$): LOCATE 21, 36: PRINT USING "$####.##"; VAL(in$): RETURN
1291
1292 IF left@ = amount@ THEN 1295 ELSE LOCATE 16, 50: PRINT CHR$(7); : COLOR 13: PRINT "Not Fully Distributed Yet!": GOTO 1288
1295 GOSUB printutilityreceipt
     d$ = LEFT$(PostDate$, 6) + RIGHT$(PostDate$, 2)
     PDate$ = d$: TM$ = LEFT$(TIME$, 2) + MID$(TIME$, 4, 2)
     postingdate$ = d$
     GOSUB 60000
4081 GET 1, AcctRecord!
      LOCK #1, AcctRecord!
       lockflag = 1

     IF pcode = 3 THEN LSET waterdep$ = MKS$(CVS(waterdep$) - amount@): PUT 1, AcctRecord!
     IF pcode = 4 OR pcode = 5 THEN
       LSET waterdep$ = MKS$(CVS(waterdep$) + amount@)
       LSET depdate$ = pdesc$
       PUT 1, AcctRecord!: GOTO 4100
     END IF

4090 GOSUB 4700
     newbalance@ = CVD(CURBAL$) + CVD(pastdue$)
     newbalance@ = INT((newbalance@ * 100) + .5) / 100
     IF newbalance@ = 0 AND acctstatus$ = "B" THEN LSET acctstatus$ = "I"
     PUT 1, AcctRecord!
4100 GOSUB 4600
4101 GOSUB 61000: REM post to transaction file
     LSET tdate$ = postingdate$
     LSET tdesc$ = pdesc$
     LSET tcuramt$ = MKS$(amount@)
     LSET POSTTIME$ = TM$
     LSET OPERNUMB$ = Operator$
     LSET tglpcode$ = "N"
     LSET tamt$ = MKD$(newbalance@)
     IF pcode = 1 OR pcode = 4 THEN LSET tcustloc$ = "1"
     IF pcode = 2 OR pcode = 5 THEN LSET tcustloc$ = "2"
4105 IF pcode < 3 THEN pcode = 3: GOTO 4110
4106 IF pcode = 3 THEN pcode = 4: GOTO 4110
     IF pcode = 4 OR pcode = 5 THEN pcode = 6: GOTO 4110
     IF pcode = 6 THEN pcode = 5
     IF pcode = 7 THEN pcode = 7
4110 LSET ttype$ = MKI$(pcode)
4120 LSET tmrec$ = MKS$(AcctRecord!)
     LSET tnexttr$ = MKS$(0)
     LSET wbrev1$ = MKS$(Apply!(1)): LSET wbrev2$ = MKS$(Apply!(2))
     LSET wbrev3$ = MKS$(Apply!(3)): LSET wbrev4$ = MKS$(Apply!(4))
     LSET wbrev5$ = MKS$(Apply!(5)): LSET wbrev6$ = MKS$(Apply!(6))
     LSET wbrev7$ = MKS$(Apply!(7)): LSET wbrev8$ = MKS$(Apply!(8))
4130 PUT 2, LOF(2) / 128 + 1: TRECORD! = LOF(2) / 128
     IF CVD(firsttr$) = 0 THEN
      LSET firsttr$ = MKD$(TRECORD!)
       LSET lasttr$ = MKD$(TRECORD!)
        PUT 1, AcctRecord!
         CLOSE 2
         GOTO 4200
      END IF
4150 prev! = CVD(lasttr$): LSET lasttr$ = MKD$(TRECORD!): PUT 1, AcctRecord!
     GET 2, prev!: LSET tnexttr$ = MKS$(TRECORD!): PUT 2, prev!: CLOSE 2
4200 IF lockflag = 1 THEN
      lockflag = 0: UNLOCK #1, AcctRecord!
     END IF
     CLOSE
     RUN

4600 REM reduce individual amount
     GET 1, AcctRecord!
      LSET IRev1$ = MKD$(CVD(IRev1$) - Apply!(1))
      LSET IRev2$ = MKD$(CVD(IRev2$) - Apply!(2))
      LSET IRev3$ = MKD$(CVD(IRev3$) - Apply!(3))
      LSET IRev4$ = MKD$(CVD(IRev4$) - Apply!(4))
      LSET IRev5$ = MKD$(CVD(IRev5$) - Apply!(5))
      LSET IRev6$ = MKD$(CVD(IRev6$) - Apply!(6))
      LSET irev7$ = MKD$(CVD(irev7$) - Apply!(7))
      LSET IREV8$ = MKD$(CVD(IREV8$) - Apply!(8))
     PUT 1, AcctRecord!
     RETURN

4700 IF pcode = 6 THEN
      LSET CURBAL$ = MKD$(CVD(CURBAL$) - amount@)
      RETURN
     END IF
     a4# = amount@: E# = CVD(pastdue$): j# = E# - a4#: y = SGN(j#)
4710 IF y = -1 THEN LSET pastdue$ = MKD$(0): RB# = a4# - E#: GOTO 4740
4720 IF y = 0 THEN LSET pastdue$ = MKD$(0): RETURN
4730 IF y = 1 THEN LSET pastdue$ = MKD$(j#): RETURN
4740 E1# = CVD(CURBAL$): J1# = E1# - RB#: Y1 = SGN(J1#)
4750 IF Y1 = -1 THEN LSET CURBAL$ = MKD$(J1#): RETURN
4760 IF Y1 = 0 THEN LSET CURBAL$ = MKD$(0): RETURN
4770 IF Y1 = 1 THEN LSET CURBAL$ = MKD$(J1#): RETURN



1900 COLOR 11: LOCATE 19, 1: PRINT "Payment Codes Are As Follows: "
1910 COLOR 14: LOCATE 20, 1: PRINT "1= Reg Payment - CASH      2= Reg Payment - CHECK      3= Apply Deposit"
1920 LOCATE 21, 1: PRINT "4= Deposit Payment - CASH  5= Deposit Payment - CHECK  6= ADJUSTMENT"
     LOCATE 22, 1: PRINT "7= Direct Deposit"
1930 RETURN
1950 IF pcode = 1 THEN dd$ = "Reg Payment - CASH"
     IF pcode = 2 THEN dd$ = "Reg Payment - CHECK"
     IF pcode = 3 THEN dd$ = "Apply Deposit"
     IF pcode = 4 THEN dd$ = "Deposit Payment - CASH"
     IF pcode = 5 THEN dd$ = "Deposit Payment - CHECK"
     IF pcode = 6 THEN dd$ = "ADJUSTMENT/CREDIT"
     IF pcode = 7 THEN dd$ = "Direct/Draft"
1960 LOCATE 19, 1: PRINT STRING$(79, 32)
     PRINT STRING$(79, 32): PRINT STRING$(79, 32)
     PRINT STRING$(79, 32); : RETURN




fixamts:
GOSUB openfile
GET 14, 1
max = CVI(REVTRACK$)
fix1:
 LOCATE 15, 1: COLOR 11: PRINT "Enter Item Number to Edit (ESC-When Finished) .. "; : fl = -1: GOSUB 62000
 IF cf = 1 THEN LOCATE 15, 1: PRINT STRING$(79, 32): RETURN
 IF VAL(in$) < 1 OR VAL(in$) > max THEN GOTO fix1
 IF VAL(in$) = 1 THEN LOCATE 6, 63: fl = -10: GOSUB 62000: LSET IRev1$ = MKD$(VAL(in$)): PUT 1, AcctRecord!: GOTO fix1
 IF VAL(in$) = 2 THEN LOCATE 7, 63: fl = -10: GOSUB 62000: LSET IRev2$ = MKD$(VAL(in$)): PUT 1, AcctRecord!: GOTO fix1
 IF VAL(in$) = 3 THEN LOCATE 8, 63: fl = -10: GOSUB 62000: LSET IRev3$ = MKD$(VAL(in$)): PUT 1, AcctRecord!: GOTO fix1
 IF VAL(in$) = 4 THEN LOCATE 9, 63: fl = -10: GOSUB 62000: LSET IRev4$ = MKD$(VAL(in$)): PUT 1, AcctRecord!: GOTO fix1
 IF VAL(in$) = 5 THEN LOCATE 10, 63: fl = -10: GOSUB 62000: LSET IRev5$ = MKD$(VAL(in$)): PUT 1, AcctRecord!: GOTO fix1
 IF VAL(in$) = 6 THEN LOCATE 11, 63: fl = -10: GOSUB 62000: LSET IRev6$ = MKD$(VAL(in$)): PUT 1, AcctRecord!:  GOTO fix1
 IF VAL(in$) = 7 THEN LOCATE 12, 63: fl = -10: GOSUB 62000: LSET irev7$ = MKD$(VAL(in$)): PUT 1, AcctRecord!:  GOTO fix1
 IF VAL(in$) = 8 THEN LOCATE 13, 63: fl = -10: GOSUB 62000: LSET IREV8$ = MKD$(VAL(in$)): PUT 1, AcctRecord!:  GOTO fix1
60000 CLOSE 1: OPEN "wbcust.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 1024
60010 FIELD 1, 25 AS lname$, 25 AS firstname$, 30 AS add1$, 30 AS add2$, 30 AS servadd$
      FIELD 1, 140 AS du1$, 20 AS CITY$, 2 AS ST$, 9 AS ZP$, 12 AS phone$
      FIELD 1, 183 AS du2$, 2 AS BOOK$, 12 AS meter$, 1 AS BillCycle$
      FIELD 1, 198 AS DU3$, 8 AS firsttr$, 8 AS lasttr$, 2 AS CUSTTYPE$
      FIELD 1, 216 AS du4$, 8 AS CURBAL$, 8 AS pastdue$, 1 AS acctstatus$
      FIELD 1, 233 AS du5$, 12 AS ssn$, 12 AS drvlic$
      FIELD 1, 257 AS DU6$, 6 AS WMINSTALL$, 6 AS GMINSTALL$
      FIELD 1, 269 AS du7$, 4 AS wAVGUSE$, 2 AS wAVGCT$, 4 AS gavguse$, 2 AS gavgct$
      FIELD 1, 281 AS du8$, 8 AS wcurread$, 8 AS wpastread$, 6 AS wcurreaddate$
      FIELD 1, 303 AS DU9$, 6 AS wpastreaddate$, 10 AS wmeternum$, 2 AS wminimum$
      FIELD 1, 321 AS du10$, 1 AS wmeterest$, 1 AS gmeterest$
      FIELD 1, 323 AS du11$, 8 AS gcurread$, 8 AS gpastread$, 6 AS gcurreaddate$
      FIELD 1, 345 AS du12$, 6 AS gpastreaddate$, 10 AS gmeternum$, 2 AS gminimum$
      FIELD 1, 363 AS DU13$, 4 AS waterdep$, 4 AS gasdep$
      FIELD 1, 371 AS du14$, 8 AS IRev1$, 8 AS IRev2$, 8 AS IRev3$
      FIELD 1, 395 AS du15$, 8 AS IRev4$, 8 AS IRev5$, 8 AS IRev6$
      FIELD 1, 419 AS DU16$, 8 AS irev7$, 8 AS IREV8$
      FIELD 1, 435 AS du17$, 2 AS rate1$, 2 AS rate2$, 2 AS rate3$
      FIELD 1, 441 AS DU18$, 2 AS rate4$, 2 AS rate5$, 2 AS rate6$
      FIELD 1, 449 AS du19$, 2 AS rate7$, 2 AS rate8$
      FIELD 1, 453 AS DU20$, 1 AS workorder$, 4 AS workordernumber$
      FIELD 1, 458 AS du21$, 6 AS workorderbeg$, 6 AS workorderend$
      FIELD 1, 470 AS DU22$, 4 AS messfirst$, 4 AS messlast$
      FIELD 1, 478 AS du23$, 30 AS comment$, 4 AS seqnum$
      FIELD 1, 512 AS du24$, 4 AS firstwo$, 4 AS lastwo$
      FIELD 1, 520 AS du25$, 6 AS mtrtime$, 8 AS etime$, 3 AS prorate$
      FIELD 1, 537 AS du26$, 4 AS amtowed$, 4 AS amtpd$, 4 AS payment$, 2 AS rs$
      FIELD 1, 551 AS du27$, 4 AS meterseq$, 4 AS garbpickups$, 1 AS WTAXRATE$, 1 AS GTAXRATE$, 1 AS bankdraft$, 462 AS NotUsed$
60030 RETURN

60100 CLOSE 6: OPEN "wbcust1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #6 LEN = 256
60110 FIELD 6, 4 AS RAMT1$, 4 AS RAMT2$, 4 AS RAMT3$, 25 AS RDESC1$, 25 AS RDESC2$, 25 AS RDESC3$, 9 AS RACCT1$, 9 AS RACCT2$, 9 AS RACCT3$, 4 AS OAMT1$, 4 AS OAMT2$, 25 AS ODESC1$, 25 AS ODESC2$, 9 AS OACCT1$, 9 AS OACCT2$, 66 AS blank$
60120 RETURN
61000 CLOSE 2: OPEN "wbartran.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 128
61010 FIELD 2, 8 AS tdate$, 8 AS tamt$, 2 AS ttype$, 15 AS tdesc$, 4 AS tmrec$, 4 AS tnexttr$, 4 AS tcuramt$
      FIELD 2, 45 AS tu1$, 4 AS wbrev1$, 4 AS wbrev2$, 4 AS wbrev3$, 4 AS wbrev4$, 4 AS wbrev5$, 4 AS wbrev6$, 4 AS wbrev7$, 4 AS wbrev8$
      FIELD 2, 77 AS tu2$, 1 AS testimate$, 8 AS tcurread$, 8 AS tpastread$
      FIELD 2, 94 AS TU3$, 8 AS tconsump$, 2 AS tbook$, 2 AS tcusttype$, 1 AS tcustloc$: REM tcustloc$ = 1 for cash and 2 for check in payment transaction
      FIELD 2, 107 AS TU4$, 1 AS tstatus$, 2 AS OPERNUMB$, 4 AS POSTTIME$, 1 AS tglpcode$, 13 AS wbblank$
61020 RETURN
61500 CLOSE 15: OPEN "wbmessag.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #15 LEN = 1600
61510 FOR xx = 0 TO 19: FIELD 15, (80 * xx) AS DUM$, 80 AS messline$(xx): NEXT xx: RETURN
      RETURN


62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(X): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
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
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 62 THEN cf = 14: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 63 THEN cf = 15: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 64 THEN cf = 16: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 65 THEN cf = 17: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 66 THEN cf = 18: RETURN
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
63000 cf = 0: in$ = "": COLOR 14: RO = CSRLIN: PO = POS(X): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
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

63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SysDate$; TAB(70); Scren$: PRINT STRING$(79, "_"): RETURN

openfile:
CLOSE 14
OPEN "wbinfo.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #14 LEN = 512
FIELD 14, 30 AS dispname$, 20 AS defcity$, 2 AS defst$, 5 AS defzip$
FIELD 14, 57 AS du1$, 2 AS REVTRACK$
FIELD 14, 59 AS du2$, 22 AS rev1name$, 1 AS rev1code$
FIELD 14, 82 AS DU3$, 22 AS rev2name$, 1 AS rev2code$
FIELD 14, 105 AS du4$, 22 AS rev3name$, 1 AS rev3code$
FIELD 14, 128 AS du5$, 22 AS rev4name$, 1 AS rev4code$
FIELD 14, 151 AS DU6$, 22 AS rev5name$, 1 AS rev5code$
FIELD 14, 174 AS du7$, 22 AS rev6name$, 1 AS rev6code$
FIELD 14, 197 AS du8$, 22 AS rev7name$, 1 AS rev7code$
FIELD 14, 220 AS DU9$, 22 AS rev8name$, 1 AS rev8code$
FIELD 14, 243 AS du10$, 269 AS blank$
RETURN


displayhistory:
     dev$ = "scrn:": OPEN "o", 10, dev$: WIDTH #10, 255
     CLS : GOSUB 39000
31200 tr! = CVD(firsttr$): GOSUB 61000
     GOSUB translink
31300 IF counter = 0 THEN 32000
     firstone = counter
31350 FOR kkk = counter TO 1 STEP -1
31400 GET 2, translink!(kkk)
31450 REM
31460 IF lc >= 18 THEN
      firstone = kkk
      LOCATE 22, 1: COLOR 11
      PRINT "Press <F1>-Next Page  :  <F2>-Prev Page (ESC-Quit) .. "; : fl = 1: GOSUB 62000
      IF cf = 1 THEN CLOSE : RETURN
      IF cf = 11 THEN CLS : GOSUB 39000: GOTO 31500

      IF cf = 12 THEN
       counter = firstone + 24
       IF counter > maxcount THEN counter = maxcount
       CLS : GOSUB 39000: GOTO 31350
      END IF
      ELSE GOTO 31500
      END IF

31500 PRINT #10, tdate$; TAB(10); tdesc$; TAB(30);
31600 IF CVI(ttype$) = 1 THEN PRINT #10, "Charge"; : GOTO 31700
31620 IF CVI(ttype$) = 2 THEN PRINT #10, "Late Chrge"; : GOTO 31700
31640 IF CVI(ttype$) = 3 THEN PRINT #10, "Payment"; : GOTO 31700
31660 IF CVI(ttype$) = 4 THEN PRINT #10, "Pay by Dep"; : GOTO 31700
31670 IF CVI(ttype$) = 5 THEN PRINT #10, "Adjustment"; : GOTO 31700
     IF CVI(ttype$) = 6 THEN PRINT #10, "Dep. Pymt"; : GOTO 31700
     IF CVI(ttype$) = 7 THEN PRINT #10, "Dir. Dep."; : GOTO 31700
31680 PRINT #10, "";
31700 PRINT #10, TAB(45); USING "#####.##"; CVS(tcuramt$);
31710 PRINT #10, TAB(60); USING "$$#####,.##"; CVD(tamt$): lc = lc + 1
31720 NEXT kkk
32000 LOCATE 22, 1
      COLOR 11
      PRINT "That's All, Press any key to Continue "; : fl = 1: GOSUB 62000: CLOSE 10: RETURN


translink:
counter = 0
trlink1:
IF tr! = 0 THEN maxcount = counter: RETURN
 counter = counter + 1
 translink!(counter) = tr!: GET 2, tr!
 tr! = CVS(tnexttr$): GOTO trlink1


39000 PRINT #10, TAB(center); TOWN$
     PRINT #10, TAB(24); "Utility Customer Account History"
39100 PRINT #10, "Date: "; DATE$:
     PRINT #10, " Loc # "; m$; TAB(60); "Acct #"; Record: PRINT #10, ""
39200 PRINT #10, "Account of ... "; : B = INSTR(firstname$, "  "): IF B = 0 THEN B = 20
39300 PRINT #10, LEFT$(firstname$, B); lname$; TAB(55); "Balance: "; USING "$####.##"; (CVD(CURBAL$) + CVD(pastdue$)): PRINT #10, ""
39400 PRINT #10, "  Date"; TAB(10); "Description"; TAB(30); "Type"; TAB(44); "  Amount"; TAB(63); "Balance": PRINT #10, STRING$(70, "="): lc = 6
     RETURN


namesearch:
LOCATE 3, 1:
COLOR 11: PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : fl = 25
 GOSUB 62000: IF cf = 1 THEN GGO = 0: RETURN
 nn$ = in$
2 nn = LEN(nn$)
 CLOSE 4: OPEN "wblname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24: FIELD 4, 20 AS lastnam$, 4 AS mrec$
 CLOSE 1: GOSUB 60000
FOR ll = 1 TO LOF(4) / 24: GET 4, ll
IF LEFT$(lastnam$, nn) = nn$ THEN GOTO n1
n3:
NEXT ll
GGO = 0: RETURN
n1:
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
PRINT "  Location Number: "; m$
PRINT "   Account Number: "; CVS(mrec$)
ACCTNUMBER$ = m$
PRINT
n2:
LOCATE 9, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO n2
IF in$ = "Q" THEN GGO = 0: RETURN
IF in$ = "Y" THEN GGO = 1: Record = CVS(mrec$): RETURN
IF in$ = "N" THEN GOTO n3
GOTO n2

printutilityreceipt:
LOCATE 23, 1: PRINT STRING$(79, 32)
LOCATE 23, 1: PRINT CHR$(7);
LOCATE 23, 1: COLOR 15
PRINT "Print RECEIPT (Y/N) .. "; : fl = 1
GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO printutilityreceipt
IF in$ = "N" THEN RETURN
LPRINT
LPRINT "CITY OF HOHENWALD"
LPRINT "UTILITIES RECEIPT"
LPRINT "    "; DATE$
LPRINT
LPRINT
LPRINT "Account#: "; AcctRecord!
LPRINT "Customer: "; LEFT$(nme$, 30)
LPRINT "Serv. At: "; LEFT$(servadd$, 30)
LPRINT
LPRINT
REM 14 LINES PRINTED TO HERE
IF pcode = 4 OR pcode = 5 THEN
 LPRINT "Amount Paid on"
 LPRINT "Deposit: "; USING "####,#.##"; amount@
 LPRINT
 LPRINT
ELSE
LPRINT "Amount Owed: "; USING "####,#.##"; owed#
IF pcode = 6 THEN
   LPRINT "Amount Adj.: "; USING "####,#.##"; amount@
  ELSE
   LPRINT "Amount Paid: "; USING "####,#.##"; amount@
END IF
LPRINT
  IF (owed# - amount@) >= 0 AND (owed# - amount@ < .01) THEN
    LPRINT "Account Paid in Full"
   ELSE
    LPRINT "Acct Balance: "; USING "####,#.##"; owed# - amount@
  END IF

END IF

LPRINT
LPRINT "LOC # "; LEFT$(meter$, 8)
LPRINT "Operator: "; filen$
LPRINT
LPRINT
LPRINT TAB(7); "T H A N K   Y O U !"
LPRINT
LPRINT
LPRINT
LPRINT
GOTO printutilityreceipt


messagecenter:

REM initialize values
 FOR ll = 1 TO 20: m$(ll) = "": NEXT ll: flag = 0
IF CVS(messfirst$) = 0 THEN
 nomessage = 1
 ELSE
 GOSUB 61500
 rec = CVS(messfirst$)
 GET 15, rec
 FOR ll = 0 TO 19: m$(ll + 1) = messline$(ll): NEXT ll
 flag = 1
END IF

CLS
COLOR 11
PRINT "* * * * * * * * * * * * *  M E S S A G E    C E N T E R    * * * * * * * * *"
FOR X = 1 TO 20
COLOR 11
PRINT USING "##"; X;
PRINT "-";
IF flag = 1 THEN COLOR 15: PRINT LEFT$(m$(X), 76) ELSE PRINT ""
COLOR 11
NEXT X
LOCATE 23, 1: COLOR 11
PRINT "Press <ENTER> to Continue "; : fl = 1: GOSUB 62000
CLOSE 15: RETURN

reviewmessage:
 GOSUB 61500
 rec = CVS(messfirst$)
 IF rec = 0 THEN RETURN
 GET 15, rec
 IF ASC(LEFT$(messline$(0), 1)) > 32 THEN messageflag = 1 ELSE messageflag = 0
 RETURN

OldSetOperatorName:
  Action = 1
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM AS CMFile LEN = CMOperRecLen
  GET CMFile, OperRecNumber, CMOperRec(1)
  LOCATE 1, 55: COLOR 15: PRINT "Oper: "; LEFT$(CMOperRec(1).OperatorName, 19)
  LOCATE 1, 1: COLOR 11: PRINT "Post Date: "; : COLOR 15: PRINT PostDate$
  Operator = CMOperRec(1).OperatorNumber
  Operator$ = STR$(Operator)
  Operator$ = RIGHT$(Operator$, LEN(Operator$) - 1)
  CLOSE CMFile
  RETURN


