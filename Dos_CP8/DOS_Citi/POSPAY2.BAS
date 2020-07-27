10 REM (C)1989 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
   CLEAR
11 DIM SHARED tc#(20), td#(20), tn$(20), ta$(20), trec(20), ta#(20), tot#(100), dacct(800), dcramt#(800), ddramt#(800), PDATE$(800), totamt(12), misc#(30)
   DIM MISCCODE$(12)
   ON ERROR GOTO networkerror
   CLOSE : OPEN "R", 1, "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
   currentyear$ = year$: CLOSE
60 CLS : sysdate$ = DATE$: c = 0
120 CLS : GOSUB 63500
   GOSUB opfile
   LOCATE 3, 1: COLOR 11
   PRINT "Enter Your Operator Number Now .. "; : fl = -2
   GOSUB 62000
   IF CF = 1 THEN RUN "POSPAY"
   record = VAL(in$)
   IF record < 1 OR record > LOF(8) / 64 THEN CLOSE : GOTO 120
   GET 8, record
   COLOR 11
   PRINT "Enter Your Password Now .. "; : fl = 5: GOSUB 63000
   IF CF = 1 THEN RUN "POSPAY"
  IF in$ = oppassword$ THEN
   OPERATORNUMBER = record
   operatorname$ = opname$
   filen$ = STR$(OPERATORNUMBER)
   filen$ = RIGHT$(filen$, LEN(filen$) - 1)
   OPERATORFILE$ = "pospay" + filen$ + ".dat"
   TOPERATORFILE$ = "tbpay" + filen$ + ".dat"
   TOPERATORFILE1$ = "tbpay1" + filen$ + ".dat"
   aroperatorfile$ = "arpay" + filen$ + ".dat"
   CLOSE
  ELSE
   PRINT "Invalid Password : Please Begin Again"
   PRINT "Press <ENTER> to Continue "; : fl = 1: GOSUB 62000: CLOSE : GOTO 10
  END IF
   SCREN$ = "4.00.00": sysdate$ = LEFT$(DATE$, 6) + RIGHT$(DATE$, 2)
   CLS
   GOSUB 63500
130 LOCATE 5, 1
    COLOR 11

3000 REM
3010 SCREN$ = "4.03.00": CLS : GOSUB 63500
     totwater# = 0: totsewer# = 0: totmisc# = 0
     totsewertax# = 0: totpenalty# = 0: tmisc# = 0
     FOR J = 1 TO 10: tot#(J) = 0: NEXT J
3020 LOCATE 1, 30: COLOR 11: PRINT "Print Payment Edit List"
     LOCATE 3, 1: PRINT "Payment Edit List for Operator "; : COLOR 15: PRINT operatorname$
3030 LOCATE 5, 1: COLOR 11
     PRINT "Press Enter to Print (<ESC>=Abort) .. "; : fl = 1: GOSUB 62000
     IF CF = 1 THEN RUN "pospay"
3050 OPEN OPERATORFILE$ FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 128
     FIELD 2, 2 AS psource$, 8 AS pmeter$, 4 AS prec$, 8 AS pamt$, 15 AS pdesc$, 2 AS ptype$, 30 AS pname$, 4 AS TFDIST$, 4 AS TLDIST$, 4 AS trev1$, 4 AS trev2$, 4 AS trev3$, 4 AS trev4$, 4 AS trev5$, 4 AS trev6$, 4 AS trev7$, 4 AS trev8$, 1 AS  _
ttyp$, 8 AS PDATE$
3060 tot# = 0
     FOR mainloop = 1 TO 4
    
     IF mainloop = 1 THEN
      IF LOF(2) / 128 = 0 THEN CLOSE : GOTO 3090 ELSE GOSUB 3900
     END IF
     IF mainloop = 2 THEN
       GOSUB 3000.1: CLOSE 2
       OPEN OPERATORFILE$ FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 128
       FIELD 2, 2 AS psource$, 8 AS pmeter$, 4 AS prec$, 8 AS pamt$, 15 AS pdesc$, 2 AS ptype$, 30 AS pname$, 4 AS TFDIST$, 4 AS TLDIST$, 4 AS trev1$, 4 AS trev2$, 4 AS trev3$, 4 AS trev4$, 4 AS trev5$, 4 AS trev6$, 4 AS trev7$, 4 AS trev8$, 1 AS  _
ttyp$, 8 AS PDATE$
       GOTO 3090
     END IF
     IF mainloop = 3 THEN
      GOSUB printar
      CLOSE 2
       OPEN OPERATORFILE$ FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 128
       FIELD 2, 2 AS psource$, 8 AS pmeter$, 4 AS prec$, 8 AS pamt$, 15 AS pdesc$, 2 AS ptype$, 30 AS pname$, 4 AS TFDIST$, 4 AS TLDIST$, 4 AS trev1$, 4 AS trev2$, 4 AS trev3$, 4 AS trev4$, 4 AS trev5$, 4 AS trev6$, 4 AS trev7$, 4 AS trev8$, 1 AS  _
ttyp$, 8 AS PDATE$
      GOTO 3090
     END IF

     IF mainloop = 4 THEN GOSUB 3910
    
     REM mainloop1=utility:mainloop2=tax:mainloop3=a/r:mainloop4=misc

     FOR X = 1 TO LOF(2) / 128: GET 2, X
     IF CVI(psource$) = 0 THEN 3080
     IF mainloop = 1 AND CVI(psource$) = 1 THEN 3061
     IF mainloop = 4 AND CVI(psource$) = 4 THEN 3061 ELSE 3080
3061 IF CVI(psource$) = 99 THEN 3080: REM already posted but not cleared out
3065 typ = 1
3070 GOSUB 3100
3080 NEXT X
     IF mainloop = 1 THEN GOSUB 3200
     IF mainloop = 4 THEN GOSUB 3250
     LPRINT CHR$(12);
3090 NEXT mainloop
     CLOSE
     LPRINT "*** Total Drawer Totals "
     LPRINT
     LPRINT "Date: "; DATE$
     LPRINT "OPERATOR # "; OPERATORNUMBER
     LPRINT "OPERATOR NAME: "; operatorname$
     LPRINT
     LPRINT "Total Cash ........... "; USING "$$#####,#.##"; grandcash#
     LPRINT "Total Checks ......... "; USING "$$#####,#.##"; grandchk#
     LPRINT "-------------------------------------"
     LPRINT "Grand Total Recpts ... "; USING "$$#####,#.##"; grandcash# + grandchk#

     LPRINT
     LPRINT "** NOTE: Check Total Includes All Tax Receipts"
     LPRINT CHR$(12);
     RUN "pospay"

3100 b = INSTR(pmeter$, "/"): IF b = 0 THEN b = 9
3105 IF lc >= 50 THEN LPRINT CHR$(12); : GOSUB 3900
3110 IF CVI(psource$) = 1 THEN LPRINT LEFT$(pmeter$, b - 1);
     IF CVI(psource$) = 4 THEN LPRINT USING "####"; CVS(prec$);
     LPRINT TAB(10); pname$; TAB(40); USING "####,.##"; CVD(pamt$);
      LPRINT TAB(50); pdesc$; TAB(66); "";
3115 IF CVI(ptype$) < 1 OR CVI(ptype$) > 10 THEN LPRINT "NOT DEFINED": GOTO 3130
3120 IF CVI(ptype$) = 1 THEN LPRINT "Reg Pymt CASH"
     IF CVI(ptype$) = 2 THEN LPRINT "Reg Pymt CHECK"
     IF CVI(ptype$) = 3 THEN LPRINT "Deposit Applied"
     IF CVI(ptype$) = 4 THEN LPRINT "Dep Pymt CASH"
     IF CVI(ptype$) = 5 THEN LPRINT "Dep Pymt CHECK"
     IF CVI(ptype$) = 6 THEN LPRINT "ADJUSTMENT"
     IF CVI(ptype$) = 7 THEN LPRINT "Dir Deposit"
     IF CVI(ptype$) = 9 THEN LPRINT "MISC Pymt CASH"
     IF CVI(ptype$) = 10 THEN LPRINT "MISC Pymt CHK"
3125 tot#(CVI(ptype$)) = tot#(CVI(ptype$)) + CVD(pamt$)
3130 typ = 1
     totwater#(typ) = totwater#(typ) + CVS(trev1$)
     totsewer#(typ) = totsewer#(typ) + CVS(trev2$)
     totpenalty#(typ) = totpenalty#(typ) + CVS(trev3$)
     tottrash#(typ) = tottrash#(typ) + CVS(trev4$)
     totmisc#(typ) = totmisc#(typ) + CVS(trev5$)
     lc = lc + 1
    
     IF CVI(psource$) = 4 THEN
      misc#(CVS(prec$)) = misc#(CVS(prec$)) + CVD(pamt$)
     END IF

     RETURN

3200 LPRINT CHR$(12);
3201 LPRINT "Cash Receipts  :  Recap Page"; TAB(55); "Date: "; DATE$
     LPRINT : LPRINT
3205 LPRINT "Payment Recap:"
     LPRINT "Total Cash Receipts ...... "; USING "$#####,#.##"; tot#(1) + tot#(4)
       grandcash# = grandcash# + tot#(1) + tot#(4)
     LPRINT "Total Check Receipts ..... "; USING "$#####,#.##"; tot#(2) + tot#(5)
       grandchk# = grankchk# + tot#(2) + tot#(5)
     LPRINT "                           --------------"
     LPRINT "Bank Deposit Total       = "; USING "$#####,#.##"; tot#(1) + tot#(2) + tot#(4) + tot#(5)
     LPRINT
     LPRINT
     LPRINT "Total Deposits Applied ... "; USING "$#####,#.##"; tot#(3)
     LPRINT "Total Adjustments Made ... "; USING "$#####,#.##"; tot#(6)
     LPRINT "                           --------------"
     LPRINT "NET TOTAL ................ "; USING "$#####,#.##"; tot#(1) + tot#(2) + tot#(3) + tot#(4) + tot#(5) + tot#(6)
     LPRINT
     LPRINT : LPRINT : LPRINT STRING$(79, "=")
     LPRINT "Breakdown by Revenue Source": LPRINT
     LPRINT "Total Water Revenue ............ "; USING "$#####,#.##"; totwater#(1)
     LPRINT "Total Sewer  Revenue ........... "; USING "$#####,#.##"; totsewer#(1)
     LPRINT "Total Penalty Revenue .......... "; USING "$#####,#.##"; totpenalty#(1)
     LPRINT "Total Trash Revenue ............ "; USING "$#####,#.##"; tottrash#(1)
     LPRINT "Total Misc. Revenue ............ "; USING "$#####,#.##"; totmisc#(1)
     FOR ll = 1 TO 2
      totwater#(ll) = 0: totsewer#(ll) = 0
      totgarbage#(ll) = 0: totsewertax#(ll) = 0
      totpenalty#(ll) = 0: totmisc#(ll) = 0
     NEXT ll
3220 RETURN
3250 LPRINT CHR$(12)
     LPRINT "Misc Cash Receipts  :  Recap Page"; TAB(55); "Date: "; DATE$
     LPRINT : LPRINT
     LPRINT "Payment Recap:"
     LPRINT "Total Cash Receipts ...... "; USING "$#####,#.##"; tot#(9)
     grandcash# = grandcash# + tot#(9)
     LPRINT "Total Check Receipts ..... "; USING "$#####,#.##"; tot#(10)
     grandchk# = grandchk# + tot#(10)
     LPRINT "                           --------------"
     LPRINT "Bank Deposit Total       = "; USING "$#####,#.##"; tot#(9) + tot#(10)
     LPRINT
     LPRINT
     CLOSE 15
     OPEN "R", 15, "posmisc.dat", 256
     FIELD 15, 30 AS codedesc$, 1 AS active$, 225 AS nu$
     FOR X = 1 TO 12
     LPRINT "Misc Code #"; X;
     GET 15, X
     IF active$ = "Y" THEN
      LPRINT TAB(20); codedesc$;
     END IF
     LPRINT TAB(55); USING "$#####,#.##"; misc#(X)
     NEXT X
     RETURN
    
3900 PAGE = PAGE + 1
   LPRINT TAB(27); "Utility Payment Entry Register"
   LPRINT "Date: "; DATE$; TAB(65); "PAGE- "; PAGE
   LPRINT "OPERATOR # "; OPERATORNUMBER
   LPRINT "OPERATOR NAME: "; operatorname$
   LPRINT
   LPRINT "  Acct #"; TAB(10); "Account Name"; TAB(42); "Amount"; TAB(50); "Description"; TAB(66); "Type": LPRINT STRING$(79, "="): lc = 4: RETURN
3910 PAGE2 = PAGE2 + 1
   LPRINT TAB(27); "Misc. Payment Entry Register"
   LPRINT "Date: "; DATE$; TAB(65); "PAGE- "; PAGE2
   LPRINT "OPERATOR # "; OPERATORNUMBER
   LPRINT "OPERATOR NAME: "; operatorname$
   LPRINT
   LPRINT "Code #"; TAB(10); "Account Name"; TAB(42); "Amount"; TAB(50); "Description"
   LPRINT STRING$(79, "="): lc = 4: RETURN



taxfile:
CLOSE 8: CLOSE 9
OPEN TOPERATORFILE$ FOR RANDOM ACCESS READ WRITE SHARED AS #8 LEN = 128
 FIELD 8, 6 AS PDATE$, 4 AS pamt$, 4 AS PTAXYR$, 20 AS PDESC1$, 2 AS ptype$, 4 AS Ptax$, 4 AS Pint$, 4 AS Padv$, 4 AS Pcol$, 4 AS plate$, 4 AS prec$, 4 AS Pctx$, 4 AS Pcin$, 4 AS Pcad$, 4 AS Pcco$, 4 AS pcla$, 4 AS pdisamt$, 6 AS ppaydate$, 2 AS  _
penvnum$
 nexrecord = LOF(8) / 128 + 1
OPEN TOPERATORFILE1$ FOR RANDOM ACCESS READ WRITE SHARED AS #9 LEN = 256
 FOR tloop = 0 TO 9
 FIELD 9, (24 * tloop) AS dummy4$, 4 AS ptx$(tloop), 4 AS pin$(tloop), 4 AS pad$(tloop), 4 AS pco$(tloop), 4 AS pla$(tloop): NEXT tloop
RETURN

3000.1 REM tax print
      tot# = 0: tot1# = 0: lc = 0
3030.1 GOSUB taxfile: GOSUB 3900.1
     IF LOF(8) / 128 = 0 THEN CLOSE 8: CLOSE 9: LPRINT CHR$(12); : RETURN
3060.1 tot# = 0: GOSUB 60000.1
     FOR X = 1 TO LOF(8) / 128
     GET 8, X: GET 9, X
     IF CVS(prec$) = 0 THEN GOTO 3080.1
3065.1 IF lc >= 56 THEN PRINT #10, CHR$(12); : GOSUB 3900.1
3070.1 GOSUB 3100.1
3080.1 NEXT X
     GOSUB 3200.1
3090.1 RETURN


3100.1 IF lc >= 50 THEN LPRINT CHR$(12); : GOSUB 3900.1
     GET 1, CVS(prec$): GET 14, CVS(CUSTOMERRECORD$)
     KK = INSTR(firstnam$, "  ")
     IF KK = 0 THEN nme$ = lnam$ ELSE nme$ = LEFT$(firstnam$, KK) + lnam$
3110.1 LPRINT "Payment Date: "; LEFT$(ppaydate$, 2); "/"; MID$(ppaydate$, 3, 2); "/"; RIGHT$(ppaydate$, 2)
     LPRINT acct$; " "; pin$; TAB(15); LEFT$(nme$, 23); TAB(40); USING "####,.##"; CVS(pamt$);
      LPRINT TAB(50); USING "####,.##"; CVS(pdisamt$);
       LPRINT TAB(61); LEFT$(PDESC1$, 14); CVI(penvnum$)
3120.1 IF CVI(ptype$) = 1 THEN
      LPRINT "Payment"; : tot# = tot# + CVS(pamt$): tot# = INT((tot# * 100) + .5) / 100
      ELSE
      LPRINT "Release"; : tot1# = tot1# + CVS(pamt$): tot1# = INT((tot1# * 100) + .5) / 100
     END IF
     tot2# = tot2# + CVS(pdisamt$)
     LPRINT TAB(12); USING "####,.##"; CVS(Ptax$);
      LPRINT TAB(24); USING "####,.##"; CVS(Pint$);
       LPRINT TAB(38); USING "####,.##"; CVS(Padv$);
        LPRINT TAB(50); USING "####,.##"; CVS(Pcol$);
         LPRINT TAB(65); USING "####,.##"; CVS(plate$)
     GOSUB breakdownyr
     LPRINT STRING$(79, "=")
3130.1 lc = lc + 4
     tottax = tottax + CVS(Ptax$): tottax = INT((tottax * 100) + .5) / 100
     totint = totint + CVS(Pint$): totint = INT((totint * 100) + .5) / 100
     totadv = totadv + CVS(Padv$): totadv = INT((totadv * 100) + .5) / 100
     totcol = totcol + CVS(Pcol$): totcol = INT((totcol * 100) + .5) / 100
     totdog = totdog + CVS(plate$): totdog = INT((totdog * 100) + .5) / 100
3140.1 RETURN

3200.1 LPRINT CHR$(12);
     LPRINT "Totals Page : Tax Receipts System"
     LPRINT "Date: "; DATE$
     LPRINT
     LPRINT "Total Amount Paid ...... "; USING "$$######,.##"; tot#
      grandchk# = grandchk# + tot#
     LPRINT "Total Amount Discounts . "; USING "$$######,.##"; tot2#
     LPRINT "Total Amount Released .. "; USING "$$######,.##"; tot1#
     LPRINT "Total A/R Reduction .... "; USING "$$######,.##"; tot# + tot1# + tot2#
     LPRINT : LPRINT : LPRINT
     LPRINT TAB(28); "* Breakdown by Year *"
     LPRINT
     LPRINT "Year"; TAB(10); "Tax Paid"; TAB(20); "Interest"; TAB(30); "Adv Paid"; TAB(40); "Col Cost"; TAB(55); "Dog Fee Pd"; TAB(70); " Total"
     LPRINT STRING$(79, "-")
     yr = VAL(currentyear$)
     LPRINT yr; TAB(10); USING "#####.##"; cyrtx;
      LPRINT TAB(20); USING "#####.##"; cyrin;
       LPRINT TAB(30); USING "#####.##"; cyrad;
        LPRINT TAB(40); USING "#####.##"; cyrco;
        LPRINT TAB(54); USING "#####.##"; cyrla;
        LPRINT TAB(69); USING "#####,#.##"; (cyrtx + cyrin + cyrad + cyrco + cyrla)
     FOR ll = 0 TO 9
     LPRINT yr - ll - 1;
      LPRINT TAB(10); USING "#####.##"; yrtax(ll);
       LPRINT TAB(20); USING "#####.##"; yrin(ll);
        LPRINT TAB(30); USING "#####.##"; yrad(ll);
         LPRINT TAB(40); USING "#####.##"; yrco(ll);
          LPRINT TAB(54); USING "#####.##"; yrla(ll);
           LPRINT TAB(69); USING "#####,#.##"; (yrtax(ll) + yrin(ll) + yrad(ll) + yrco(ll) + yrla(ll))
      NEXT ll
     

      REM TOTAL UP NOW
      FOR ll = 0 TO 9
      tax# = tax# + yrtax(ll)
      tax# = INT((tax# * 100) + .5) / 100
      intr# = intr# + yrin(ll)
      intr# = INT((intr# * 100) + .5) / 100
      adv# = adv# + yrad(ll)
      adv# = INT((adv# * 100) + .5) / 100
      col# = col# + yrco(ll)
      col# = INT((col# * 100) + .5) / 100
      la# = la# + yrla(ll)
      la# = INT((la# * 100) + .5) / 100
      NEXT ll
      tax# = tax# + cyrtx: tax# = INT((tax# * 100) + .5) / 100
      intr# = intr# + cyrin: intr# = INT((intr# * 100) + .5) / 100
      adv# = adv# + cyrad: adv# = INT((adv# * 100) + .5) / 100
      col# = col# + cyrco: col# = INT((col# * 100) + .5) / 100
      la# = la# + cyrla: la# = INT((la# * 100) + .5) / 100
      LPRINT STRING$(79, "-")
      LPRINT TAB(9); USING "######.##"; tax#;
       LPRINT TAB(19); USING "######.##"; intr#;
        LPRINT TAB(29); USING "######.##"; adv#;
         LPRINT TAB(39); USING "######.##"; col#;
          LPRINT TAB(53); USING "######.##"; la#;
           LPRINT TAB(69); USING "#####,#.##"; tax# + intr# + adv# + col# + la#
            LPRINT CHR$(12);
3220.1 RETURN


3900.1 page1 = page1 + 1
   LPRINT TAB(24); "Town of Troy : Tax Payment List"
   LPRINT "Date: "; DATE$; TAB(65); "Page- "; page1
   LPRINT "OPERATOR # "; OPERATORNUMBER
   LPRINT "OPERATOR NAME: "; operatorname$
3910.1 LPRINT "Acct/Parcel"; TAB(15); "Account Name"; TAB(40); "Amt Paid"; TAB(50); "Discount"; TAB(61); "Description  Env#"
       LPRINT "Type"; TAB(13); "Tax Paid"; TAB(24); "Interest Pd"; TAB(38); "Adv Paid"; TAB(50); "Col Cost Pd    Dog Fee Pd"
       LPRINT STRING$(79, "="): lc = 4: RETURN


breakdownyr:
yr = VAL(currentyear$): year = yr
IF CVS(Pctx$) <> 0 OR CVS(Pcin$) <> 0 OR CVS(Pcad$) <> 0 OR CVS(Pcco$) <> 0 THEN
   LPRINT "Yr- "; yr; " Tax- "; USING "#####.##"; CVS(Pctx$);
    LPRINT " Int- "; USING "####.##"; CVS(Pcin$);
     LPRINT "  Adv- "; USING "###.##"; CVS(Pcad$);
      LPRINT " Collect- "; USING "###.##"; CVS(Pcco$);
       LPRINT " Dog- "; USING "###.##"; CVS(pcla$): lc = lc + 1
      cyrtx = cyrtx + CVS(Pctx$)
      cyrtx = INT((cyrtx * 100) + .5) / 100
      cyrin = cyrin + CVS(Pcin$)
      cyrin = INT((cyrin * 100) + .5) / 100
      cyrad = cyrad + CVS(Pcad$)
      cyrad = INT((cyrad * 100) + .5) / 100
      cyrco = cyrco + CVS(Pcco$)
      cyrco = INT((cyrco * 100) + .5) / 100
      cyrla = cyrla + CVS(pcla$)
      cyrla = INT((cyrla * 100) + .5) / 100

END IF
FOR ll = 0 TO 9
yr = year - ll - 1
IF CVS(ptx$(ll)) <> 0 OR CVS(pin$(ll)) <> 0 OR CVS(pad$(ll)) <> 0 OR CVS(pco$(ll)) <> 0 OR CVS(pla$(ll)) <> 0 THEN
   LPRINT "Yr- "; yr; " Tax- "; USING "#####.##"; CVS(ptx$(ll));
    LPRINT " Int- "; USING "####.##"; CVS(pin$(ll));
     LPRINT "  Adv- "; USING "###.##"; CVS(pad$(ll));
      LPRINT " Collect- "; USING "###.##"; CVS(pco$(ll));
       LPRINT " Dog- "; USING "###.##"; CVS(pla$(ll)): lc = lc + 1
      yrtax(ll) = yrtax(ll) + CVS(ptx$(ll))
       yrtax(ll) = INT((yrtax(ll) * 100) + .5) / 100
      yrin(ll) = yrin(ll) + CVS(pin$(ll))
       yrin(ll) = INT((yrin(ll) * 100) + .5) / 100
      yrad(ll) = yrad(ll) + CVS(pad$(ll))
       yrad(ll) = INT((yrad(ll) * 100) + .5) / 100
      yrco(ll) = yrco(ll) + CVS(pco$(ll))
       yrco(ll) = INT((yrco(ll) * 100) + .5) / 100
      yrla(ll) = yrla(ll) + CVS(pla$(ll))
       yrla(ll) = INT((yrla(ll) * 100) + .5) / 100
END IF
NEXT ll
RETURN

printar:
     FOR jk = 1 TO 6: artot#(jk) = 0: NEXT jk
     CLOSE 2: OPEN aroperatorfile$ FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 256
     FIELD 2, 4 AS prec$, 8 AS pamt$, 15 AS pdesc$, 2 AS ptype$, 30 AS pname$, 4 AS TFDIST$, 4 AS TLDIST$
     IF LOF(2) / 256 = 0 THEN CLOSE 2: RETURN
     GOSUB 3900.2
3060.2 artot# = 0:
     FOR X = 1 TO LOF(2) / 256: GET 2, X:
3064.2 IF CVS(prec$) = 0 THEN GOTO 3080.2
     typ = 1
3070.2 GOSUB 3100.2
3080.2 NEXT X: GOSUB 3200.2: LPRINT CHR$(12); : CLOSE 2: RETURN

3100.2 IF lc >= 50 THEN LPRINT CHR$(12); : GOSUB 3900.2
3110.2 LPRINT CVS(prec$); TAB(10); LEFT$(pname$, 24); TAB(35); USING "####,.##"; CVD(pamt$);
     LPRINT TAB(50); " "; pdesc$; TAB(66); "";
3115.2 IF CVI(ptype$) < 1 OR CVI(ptype$) > 3 THEN LPRINT "NOT DEFINED": GOTO 3135.2
3120.2 IF CVI(ptype$) = 1 THEN LPRINT "Reg Pymt CASH"
     IF CVI(ptype$) = 2 THEN LPRINT "Reg Pymt CHECK"
     IF CVI(ptype$) = 3 THEN LPRINT "ADJUSTMENT"
3125.2 totar#(CVI(ptype$)) = totar#(CVI(ptype$)) + CVD(pamt$)
3135.2 lc = lc + 1
3170.2 RETURN

3200.2 LPRINT STRING$(79, "-"): LPRINT CHR$(12): LPRINT
3201.2 LPRINT "A/R Cash Receipts  :  Recap Page"; TAB(55); "Date: "; DATE$
     LPRINT : LPRINT
3205.2 LPRINT "Payment Recap:"
3210.2 LPRINT "Total Cash Receipts ...... "; USING "$#####,#.##"; totar#(1)
       grandcash# = grandcash# + totar#(1)
     LPRINT "Total Check Receipts ..... "; USING "$#####,#.##"; totar#(2)
       grandchk# = grandchk# + totar#(2)
     LPRINT "                           --------------"
     LPRINT "Bank Deposit Total       = "; USING "$#####,#.##"; totar#(1) + totar#(2)
     LPRINT
     LPRINT
     LPRINT "Total Adjustments Made ... "; USING "$#####,#.##"; totar#(3)
     LPRINT "                           --------------"
     LPRINT "NET TOTAL ................ "; USING "$#####,#.##"; totar#(1) + totar#(2) + totar#(3)
     LPRINT : LPRINT : LPRINT STRING$(79, "=")
   
3220.2 RETURN
3900.2 GOSUB openprintinfo:
      IF LOF(12) = 0 THEN cname$ = "" ELSE GET 12, 1: cname$ = companyname$
      CLOSE 12
      KK = INSTR(cname$, "  ")
      IF KK = 1 THEN tabnumB = 40 ELSE tabnumB = 40 - (INT(KK / 2))
      LPRINT TAB(tabnumB); cname$
      LPRINT TAB(30); "A/R Payment Listing":
      LPRINT "Date: "; DATE$; TAB(65); "Page- "; page1
      LPRINT "OPERATOR # "; OPERATORNUMBER
      LPRINT "OPERATOR NAME: "; operatorname$
3910.2 LPRINT "  Acct #"; TAB(10); "Account Name"; TAB(38); "Amount"; TAB(53); "Description"; TAB(66); "Type": LPRINT STRING$(79, "="): lc = 4: RETURN


opfile:
CLOSE 8: OPEN "posoper.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #8 LEN = 64
         FIELD 8, 30 AS opname$, 5 AS oppassword$, 29 AS blk$
         RETURN

60000 CLOSE 1: OPEN "wbcust.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 256
60010 FIELD 1, 25 AS lname$, 25 AS firstname$, 30 AS ADD1$, 30 AS add2$, 20 AS CITY$, 2 AS st$, 9 AS zp$, 12 AS PHONE$, 2 AS BOOK$, 12 AS meter$, 2 AS RATE$, 8 AS WON$, 8 AS MINSTALL$, 8 AS AVGUSE$, 2 AS AVGCT$, 8 AS CURBAL$, 8 AS PASTDUE$, 1 AS  _
cycle$, 5 AS NOTUSED$, 8 AS CURREAD$, 8 AS PASTREAD$, 8 AS firsttr$, 8 AS lasttr$, 4 AS waterdep$
60015 CLOSE 8: OPEN "wbcust2.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #8 LEN = 64
      FIELD 8, 8 AS IREVENUE1$, 8 AS IREVENUE2$, 8 AS IREVENUE3$, 8 AS irevenue4$, 8 AS irevenue5$, 8 AS IREVENUE6$, 8 AS IREVENUE7$, 8 AS IREVENUE8$
60030 RETURN
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
60030.1 RETURN

62000 CF = 0: in$ = "": COLOR 14: RO = CSRLIN: PO = POS(X): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN CF = 1: RETURN
      IF z$ = CHR$(9) THEN CF = 2: RETURN
62065 IF ASC(z$) = 0 THEN 62030
62070 IF LEN(z$) > 1 THEN CF = ASC(RIGHT$(z$, 1)) - 57: IF CF > 0 THEN RETURN ELSE CF = CF + 27: RETURN
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
63000 CF = 0: in$ = "": COLOR 14: RO = CSRLIN: PO = POS(X): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
63010 IF INKEY$ <> "" THEN 63010
63020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
63030 z$ = INKEY$: IF z$ = "" THEN 63030
63040 IF z$ <> CHR$(13) THEN 63060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
63050 GOSUB 63200: RETURN
63060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN CF = 1: RETURN
      IF z$ = CHR$(9) THEN CF = 2: RETURN
63065 IF ASC(z$) = 0 THEN 63030
63070 IF LEN(z$) > 1 THEN CF = ASC(RIGHT$(z$, 1)) - 57: IF CF > 0 THEN RETURN ELSE CF = CF + 27: RETURN
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
 IF CF = 1 THEN LOCATE 15, 1: PRINT STRING$(79, 32): RETURN
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
               fl = 1: GOSUB 62000: IF CF = 1 THEN CLOSE : GOTO 10 ELSE GOSUB 62000
     END SELECT

MISCCODE$(1) = "Misc."
MISCCODE$(2) = ""
MISCCODE$(3) = ""
MISCCODE$(4) = ""
MISCCODE$(5) = ""
MISCCODE$(6) = ""
MISCCODE$(7) = ""
MISCCODE$(8) = ""
MISCCODE$(9) = ""
MISCCODE$(10) = ""
MISCCODE$(11) = ""
MISCCODE$(12) = ""

openprintinfo:
     OPEN "R", 12, "bacompny.dat", 128
     FIELD 12, 40 AS companyname$, 88 AS blank$
     RETURN

