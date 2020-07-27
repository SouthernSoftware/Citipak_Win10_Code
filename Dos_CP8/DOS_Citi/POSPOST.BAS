10 REM (C)1989 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
   CLEAR
11 DIM tc#(20), td#(20), tn$(20), ta$(20), trec(20), ta#(20), tot#(100), translink!(1500)
   ON ERROR GOTO networkerror
110 SCREN$ = "4.00.00": sysdate$ = LEFT$(DATE$, 6) + RIGHT$(DATE$, 2)
120 CLS : GOSUB 63500
    IF OPSET = 1 THEN 130 ELSE GOSUB SETOPERATOR
130 LOCATE 5, 1: COLOR 11
4000 REM
4001 d$ = LEFT$(DATE$, 6) + RIGHT$(DATE$, 2): pdate$ = d$
4010 SCREN$ = "4.04.00": CLS : GOSUB 63500
4020 LOCATE 1, 30: COLOR 11: PRINT "Post Payments to Accounts"
4030 LOCATE 3, 1: COLOR 11: PRINT "Press <ENTER> to Post or (<ESC>=Abort) .. "; : fl = -6: GOSUB 62000
     IF cf = 1 THEN 10
4035 dat$ = in$
4040 OPEN OPERATORFILE$ FOR RANDOM ACCESS READ WRITE AS #9 LEN = 128: LOCK #9
     FIELD 9, 2 AS psource$, 8 AS pmeter$, 4 AS PREC$, 8 AS pamt$, 15 AS pdesc$, 2 AS ptype$, 30 AS pname$, 4 AS TFDIST$, 4 AS TLDIST$, 4 AS trev1$, 4 AS trev2$, 4 AS TREV3$, 4 AS trev4$, 4 AS trev5$, 4 AS trev6$, 8 AS pdate$


4050 FOR mainloop = 1 TO 3: REM 1=util:2=tax:3=ar
     IF mainloop = 1 THEN
      IF LOF(9) / 128 = 0 THEN GOTO 4210
     END IF
     IF mainloop = 2 THEN GOSUB taxpost: GOTO 4210
     IF mainloop = 3 THEN GOSUB arpost: GOTO 4210

4060 FOR X = 1 TO LOF(9) / 128
4065 GET 9, X
     IF CVI(psource$) = 99 THEN 4200: REM indicates previously posted
4066 IF mainloop = 1 AND CVI(psource$) = 1 THEN 4068
4068 REM check for deleted account
     IF VAL(pmeter$) = 0 THEN 4200
     IF CVS(PREC$) = 0 THEN 4200
4070 LOCATE 8, 1: COLOR 11: PRINT "Now Posting Account of .. "; : COLOR 14: PRINT pname$
4080 GOSUB 60000: GET 1, CVS(PREC$)
     IF CVI(ptype$) = 4 OR CVI(ptype$) = 5 THEN LSET waterdep$ = MKS$(CVS(waterdep$) + CVD(pamt$)): PUT 1, CVS(PREC$): GOTO 4100
     IF CVI(ptype$) = 3 THEN LSET waterdep$ = MKD$(CVS(waterdep$) - CVD(pamt$)): PUT 1, CVS(PREC$)
4090 GOSUB 4700: PUT 1, CVS(PREC$): acctbal = CVD(curbal$) + CVD(pastdue$)
4100 GOSUB 61000: LSET TDATE$ = d$: LSET tamt$ = MKD$(acctbal): LSET tcuramt$ = MKS$(CVD(pamt$))
4105 IF CVI(ptype$) < 3 THEN pcode = 3: GOTO 4110
4106 IF CVI(ptype$) = 3 THEN pcode = 4: GOTO 4110
     IF CVI(ptype$) = 4 OR CVI(ptype$) = 5 THEN pcode = 6: GOTO 4110
     IF CVI(ptype$) = 6 THEN pcode = 5
4110 LSET TTYPE$ = MKI$(pcode): LSET TDESC$ = pdesc$
4120 LSET tmrec$ = MKS$(CVS(PREC$)): LSET tnexttr$ = MKS$(0)
4130 PUT 2, LOF(2) / 50 + 1: TRECORD = LOF(2) / 50
4140 IF CVD(firsttr$) = 0 THEN LSET firsttr$ = MKD$(TRECORD): LSET lasttr$ = MKD$(TRECORD): PUT 1, CVS(PREC$): CLOSE 1: CLOSE 2: GOTO 4200
4150 prev = CVD(lasttr$): LSET lasttr$ = MKD$(TRECORD): PUT 1, CVS(PREC$): GET 2, prev: LSET tnexttr$ = MKS$(TRECORD): PUT 2, prev: CLOSE 2: CLOSE 1
     GET 8, CVS(PREC$)
     LSET IREVENUE1$ = MKD$(CVD(IREVENUE1$) - CVS(trev1$))
     LSET IREVENUE2$ = MKD$(CVD(IREVENUE2$) - CVS(trev2$))
     LSET IREVENUE3$ = MKD$(CVD(IREVENUE3$) - CVS(TREV3$))
     LSET irevenue4$ = MKD$(CVD(irevenue4$) - CVS(trev4$))
     LSET irevenue5$ = MKD$(CVD(irevenue5$) - CVS(trev5$))
     LSET IREVENUE6$ = MKD$(CVD(IREVENUE6$) - CVS(trev6$))
     PUT 8, CVS(PREC$)
4200 LSET psource$ = MKI$(99): PUT 9, X
4206 NEXT X
4210 NEXT mainloop
     CLOSE
     KILL OPERATORFILE$
     RUN "pospay"
    
4700 a4# = CVD(pamt$): E# = CVD(pastdue$): j# = E# - a4#: y = SGN(j#)
4710 IF y = -1 THEN LSET pastdue$ = MKD$(0): RB# = a4# - E#: GOTO 4740
4720 IF y = 0 THEN LSET pastdue$ = MKD$(0): RETURN
4730 IF y = 1 THEN LSET pastdue$ = MKD$(j#): RETURN
4740 E1# = CVD(curbal$): J1# = E1# - RB#: Y1 = SGN(J1#)
4750 IF Y1 = -1 THEN LSET curbal$ = MKD$(J1#): RETURN
4760 IF Y1 = 0 THEN LSET curbal$ = MKD$(0): RETURN
4770 IF Y1 = 1 THEN LSET curbal$ = MKD$(J1#): RETURN


taxpost:
     GOSUB tempfile: GOSUB 60000.1
     IF LOF(8) / 128 = 0 THEN CLOSE 8: CLOSE 9: RETURN
4060.1 FOR X = 1 TO LOF(8) / 128: GET 8, X: GET 9, X
4065.1 IF CVS(PREC$) <= 0 THEN GOTO 4205.1
     GET 1, CVS(PREC$): GET 3, CVS(PREC$)
     GET 14, CVS(CUSTOMERRECORD$)
     kk = INSTR(firstnam$, "  ")
     IF kk = 0 THEN nme$ = lnam$ ELSE nme$ = LEFT$(firstnam$, kk) + lnam$
4070.1 LOCATE 8, 1
     COLOR 11: PRINT "Now Posting Account of .. "; : COLOR 14: PRINT nme$
4090.1 GOSUB 4700.1
     GOSUB figurebalance
4100.1 acctbal# = curbal# + pastdue#
     GOSUB 61000.1
     LSET TDATE$ = ppaydate$
     LSET tamt$ = MKS$(CVS(pamt$))
     IF CVI(ptype$) = 1 THEN typ = 2 ELSE typ = 4: REM payment=2 release=4
4110.1 LSET TTYPE$ = MKI$(typ)
     LSET ttaxyr$ = MKI$(CVS(PTAXYR$))
     LSET TDESC$ = PDESC1$
     LSET taprin$ = MKS$(CVS(Ptax$)): LSET taint$ = MKS$(CVS(Pint$))
     LSET taadv$ = MKS$(CVS(Padv$)): LSET tacol$ = MKS$(CVS(Pcol$))
     LSET talate$ = MKS$(CVS(plate$))
     LSET tenvnum$ = MKI$(CVI(penvnum$))
4120.1 LSET tmrec$ = MKS$(CVS(PREC$)): LSET tnexttr$ = MKS$(0)
     LSET tbal$ = MKS$(acctbal#): LSET tdisamt$ = MKS$(CVS(pdisamt$))
     LSET tblank$ = ""
4130.1 PUT 2, LOF(2) / 96 + 1: TRECORD = LOF(2) / 96
4140.1 IF CVS(ftr$) = 0 THEN LSET ftr$ = MKS$(TRECORD): LSET ltr$ = MKS$(TRECORD): PUT 1, CVS(PREC$): GOTO 4200.1
4150.1 prev = CVS(ltr$): LSET ltr$ = MKS$(TRECORD): PUT 1, CVS(PREC$): GET 2, prev: LSET tnexttr$ = MKS$(TRECORD): PUT 2, prev
4200.1 REM
4205.1 NEXT X
     CLOSE
     KILL toperatorfile$
     KILL toperatorfile1$
     RETURN

4700.1 REM UPDATE PRIOR AND CUR YEAR BALANCE
     LSET curtaxamt$ = MKS$((CVS(curtaxamt$) - CVS(Pctx$)))
     LSET curintamt$ = MKS$((CVS(curintamt$) + CVS(Pcin$)))
     LSET curadvamt$ = MKS$((CVS(curadvamt$) - CVS(Pcad$)))
     LSET curcolamt$ = MKS$((CVS(curcolamt$) - CVS(Pcco$)))
     LSET curlateamt$ = MKS$((CVS(curlateamt$) - CVS(pcla$)))
     PUT 1, CVS(PREC$)
PRIORYEAR:
    FOR ll = 0 TO 9
    LSET ptaxamt$(ll) = MKS$((CVS(ptaxamt$(ll)) - CVS(ptx$(ll))))
    LSET pintamt$(ll) = MKS$((CVS(pintamt$(ll)) + CVS(pin$(ll))))
    LSET padvamt$(ll) = MKS$((CVS(padvamt$(ll)) - CVS(pad$(ll))))
    LSET pcolamt$(ll) = MKS$((CVS(pcolamt$(ll)) - CVS(pco$(ll))))
    LSET plateamt$(ll) = MKS$((CVS(plateamt$(ll)) - CVS(pla$(ll))))
    NEXT ll
    PUT 3, CVS(PREC$)
    RETURN

arpost:

     d$ = LEFT$(DATE$, 6) + RIGHT$(DATE$, 2): pdate$ = d$
     CLOSE 3
     OPEN aroperatorfile$ FOR RANDOM ACCESS READ WRITE SHARED AS #3 LEN = 256
     FIELD 3, 4 AS PREC$, 8 AS pamt$, 15 AS pdesc$, 2 AS ptype$, 30 AS pname$, 4 AS TFDIST$, 4 AS TLDIST$
     IF LOF(3) / 256 = 0 THEN CLOSE : RETURN
     LOCK #3
     FOR X = 1 TO LOF(3) / 256:
     GET 3, X: IF CVS(PREC$) = 0 THEN GOTO 4205.2
     LOCATE 8, 1: COLOR 11: PRINT "Now Posting Account of .. "; : COLOR 14: PRINT pname$
     GOSUB 60000.2
     GET 1, CVS(PREC$)
     acctrecord = CVS(PREC$)
     
4090.2 GOSUB 4700.2
       arnewbALANCE# = CVD(curbal$) + CVD(pastdue$): PUT 1, acctrecord
4100.2 GOSUB 61000.2: LSET tcuramt$ = MKS$(CVD(pamt$))
     LSET TDATE$ = d$: LSET TrunAMT$ = MKS$(arnewbALANCE#)
4105.2 IF CVI(ptype$) < 3 THEN pcode = 3: GOTO 4110.2
     IF CVI(ptype$) = 3 THEN pcode = 5
4110.2 LSET TTYPE$ = MKI$(pcode): LSET TDESC$ = pdesc$
4120.2 LSET tmrec$ = MKS$(acctrecord): LSET tnexttr$ = MKS$(0)
4130.2 PUT 2, LOF(2) / 128 + 1: TRECORD = LOF(2) / 128
4140.2 IF CVS(firsttr$) = 0 THEN
       LSET firsttr$ = MKS$(TRECORD)
       LSET lasttr$ = MKS$(TRECORD)
       PUT 1, acctrecord: CLOSE 1: CLOSE 2: GOTO 4205.2
     END IF
4150.2 prev = CVS(lasttr$): LSET lasttr$ = MKS$(TRECORD): PUT 1, acctrecord
     GET 2, prev: LSET tnexttr$ = MKS$(TRECORD): PUT 2, prev: CLOSE 2
     CLOSE 1
4205.2 LSET PREC$ = MKS$(0): PUT 3, X
     NEXT X
     CLOSE
     KILL aroperatorfile$
     RETURN
4700.2 a4# = CVD(pamt$): E# = CVD(pastdue$): j# = E# - a4#: y = SGN(j#)
4710.2 IF y = -1 THEN LSET pastdue$ = MKD$(0): RB# = a4# - E#: GOTO 4740.2
4720.2 IF y = 0 THEN LSET pastdue$ = MKD$(0): RETURN
4730.2 IF y = 1 THEN LSET pastdue$ = MKD$(j#): RETURN
4740.2 E1# = CVD(curbal$): J1# = E1# - RB#: Y1 = SGN(J1#)
4750.2 IF Y1 = -1 THEN LSET curbal$ = MKD$(J1#): RETURN
4760.2 IF Y1 = 0 THEN LSET curbal$ = MKD$(0): RETURN
4770.2 IF Y1 = 1 THEN LSET curbal$ = MKD$(J1#): RETURN


tempfile:
toperatorfile$ = "tbpay" + filen$ + ".dat"
toperatorfile1$ = "tbpay1" + filen$ + ".dat"
CLOSE 8: CLOSE 9
OPEN toperatorfile$ FOR RANDOM ACCESS READ WRITE SHARED AS #8 LEN = 128
 FIELD 8, 6 AS pdate$, 4 AS pamt$, 4 AS PTAXYR$, 20 AS PDESC1$, 2 AS ptype$, 4 AS Ptax$, 4 AS Pint$, 4 AS Padv$, 4 AS Pcol$, 4 AS plate$, 4 AS PREC$, 4 AS Pctx$, 4 AS Pcin$, 4 AS Pcad$, 4 AS Pcco$, 4 AS pcla$, 4 AS pdisamt$, 6 AS ppaydate$, 2 AS  _
penvnum$
 nexrecord = LOF(8) / 128 + 1
OPEN toperatorfile1$ FOR RANDOM ACCESS READ WRITE SHARED AS #9 LEN = 256
 FOR tloop = 0 TO 9
 FIELD 9, (24 * tloop) AS dummy4$, 4 AS ptx$(tloop), 4 AS pin$(tloop), 4 AS pad$(tloop), 4 AS pco$(tloop), 4 AS pla$(tloop): NEXT tloop
RETURN
figurebalance:
     curbal# = 0: pastdue# = 0
     curbal# = (CVS(curtaxamt$) + CVS(curadvamt$) + CVS(curcolamt$) + CVS(curlateamt$))
     curbal# = INT((curbal# * 100) + .5) / 100
     FOR ll = 0 TO 9
     pastdue# = pastdue# + (CVS(ptaxamt$(ll)) + CVS(padvamt$(ll)) + CVS(plateamt$(ll)) + CVS(pcolamt$(ll)))
     pastdue# = INT((pastdue# * 100) + .5) / 100
     NEXT ll
     RETURN

opfile:
CLOSE 8: OPEN "posoper.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #8 LEN = 64
         FIELD 8, 30 AS OPNAME$, 5 AS oppassword$, 29 AS blk$
         RETURN
60000 CLOSE 1: OPEN "wbcust.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 256
60010 FIELD 1, 25 AS LNAME$, 25 AS FIRSTNAME$, 30 AS add1$, 30 AS add2$, 20 AS city$, 2 AS st$, 9 AS zp$, 12 AS PHONE$, 2 AS BOOK$, 12 AS meter$, 2 AS RATE$, 8 AS WON$, 8 AS MINSTALL$, 8 AS AVGUSE$, 2 AS AVGCT$, 8 AS curbal$, 8 AS pastdue$, 1 AS  _
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
      FIELD 1, 168 AS du3$, 30 AS contact$, 2 AS Category$, 4 AS stfee$
      FIELD 1, 204 AS du4$, 12 AS license$, 6 AS validdate$, 290 AS blank$
60030.2 RETURN
61000.2 CLOSE 2: OPEN "artran.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 128
61010.2 FIELD 2, 8 AS TDATE$, 4 AS tcuramt$, 4 AS TrunAMT$, 2 AS TTYPE$, 25 AS TDESC$
      FIELD 2, 43 AS dum1$, 4 AS tmrec$, 4 AS tnexttr$, 77 AS blank1$
61020.2 RETURN


62000 cf = 0: in$ = "": COLOR 14: RO = CSRLIN: PO = POS(X): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
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

63500 LOCATE 1, 1: COLOR 10: PRINT "Post Date: "; sysdate$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN
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
   GOSUB opfile
IF LOF(8) = 0 THEN
 COLOR 11
 PRINT "No Operator File FOUND"; CHR$(7);
 PRINT "The operator file must be set up : Press <ENTER> to Continue "; : fl = 1
 GOSUB 62000
 RUN "pospay"
END IF
LOCATE 3, 1: COLOR 11
PRINT "Enter Your Operator Number Now .. "; : fl = -2
GOSUB 62000
IF cf = 1 THEN RUN "pospay"
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
  aroperatorfile$ = "arpay" + filen$ + ".dat"
  CLOSE
 ELSE
  PRINT "Invalid Password : Please Begin Again"
  PRINT "Press <ENTER> to Continue "; : fl = 1: GOSUB 62000: CLOSE : GOTO 10
END IF
SCREN$ = "4.00.00": sysdate$ = LEFT$(DATE$, 6) + RIGHT$(DATE$, 2)
CLS
GOSUB 63500
OPSET = 1
RETURN

