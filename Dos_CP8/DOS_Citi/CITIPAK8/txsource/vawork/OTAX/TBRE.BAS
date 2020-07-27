10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbcustrep
30 REM Customer Report Menu
40 REM Last Revision Date: 8-24-87
50 REM by Nick Picerno
60 CLS : CLEAR
     CLOSE : OPEN "R", 1, "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
     CURRENTYEAR$ = year$: CLOSE
     ON ERROR GOTO networkerror

     DIM SHARED BMTAX(10, 1), sdTAX(10, 1), ADVERT(10, 1), COLECT(10, 1), INTEREST(10, 1)
     DIM SHARED totpay#(2, 2, 12), totprin#(2, 2, 12), totsd#(2, 2, 12), totint#(2, 2, 12), totadv#(2, 2, 12), totcol#(2, 2, 12)
     DIM SHARED yrdesc(11)

     DEFDBL A-C, I, S

100 REM begin program
110 SCREN$ = "2.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500
130 LOCATE 4, 1: COLOR 11
200 PRINT TAB(31); "Tax Billing System"
210 PRINT TAB(30); "  Customer Reports "
220 PRINT
230 PRINT TAB(5); "1. Print Master Customer Listing";
240        PRINT TAB(45); "6. Sr. Citizen Report"
PRINT
250 PRINT TAB(5); "2. Print Customer Balances";
     PRINT TAB(45); "7. Balance by Specific Tax Year"
260 PRINT
270 PRINT TAB(5); "3. Customer Account Inquiry";
     PRINT TAB(45); "8. Print/Charge Advertising List"
280 PRINT
290 PRINT TAB(5); "4. Print Account Valuation Listing";
     PRINT TAB(45); "9. Deliquent Notice System"
300 PRINT
302 PRINT TAB(5); "5. Print Mailing Labels";
     PRINT TAB(44); "10. Monthly Recap Report"
310 LOCATE 20, 1: COLOR 11: PRINT TAB(25); "Selection No. (ESC-Master Menu) : "; : fl = -2: GOSUB 62000
320 IF CF = 1 THEN CLOSE : RUN "tbmenu"
325 IF LEN(in$) = 0 THEN 310
330 IF VAL(in$) < 1 OR VAL(in$) > 10 THEN 310
340 ON VAL(in$) GOTO 1000, 2000, 3000, 4000, 5000, 6000, 7000, 8000, 9000, 10000

1000 REM
1010 SCREN$ = "2.01.00": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 32: PRINT "Customer Account Listing"
1020 LOCATE 4, 1
1030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount or (L)ast Name (A,L or ESC-ABORT) ..... "; : fl = 1: GOSUB 62210
1040 IF CF = 1 THEN 10
1050 IF NOT (in$ = "A" OR in$ = "L") THEN 1020
1060 SORTSPEC$ = in$: LOCATE 4, 79: PRINT SORTSPEC$
1070 LOCATE 6, 1
1080 COLOR 11: PRINT "Output Specification: Output to (S)creen or (P)rinter (S,P or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
1090 IF CF = 1 THEN 10
1100 IF NOT (in$ = "S" OR in$ = "P") THEN 1070
1105 devspec$ = in$: LOCATE 6, 77: PRINT devspec$
1110 LOCATE 7, 1: COLOR 11: PRINT STRING$(79, "="): COLOR 10: LOCATE 13, 23: PRINT "Now Processing : Please be Patient!": COLOR 14
     LOCATE 23, 1: PRINT "Press <ESC> to stop printing . . . . "
1120 IF devspec$ = "S" THEN DEV$ = "SCRN:" ELSE DEV$ = "lpt1:"
1130 OPEN "O", 10, DEV$: WIDTH #10, 255
1140 GOSUB 60000: IF devspec$ = "S" THEN CLS : GOSUB 1900 ELSE GOSUB 1900
1150 IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10
1160 IF SORTSPEC$ = "A" THEN
      OPEN "tbacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 31
      FIELD 4, 27 AS met$, 4 AS mrec$: ENDOFFILE = LOF(4) / 31
     ELSE
      OPEN "tblname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
      FIELD 4, 20 AS lastnam$, 4 AS mrec$: ENDOFFILE = LOF(4) / 24
     END IF
1170 FOR x = 1 TO ENDOFFILE: GET 4, x: IF CVS(mrec$) = 0 THEN 1230
1180 GET 1, CVS(mrec$): B = INSTR(acct$, " ")
     IF B = 0 THEN M$ = LEFT$(acct$, 8) ELSE M$ = LEFT$(acct$, B - 1)
     k = INSTR(lnam$, "  "): IF k = 0 THEN k = 31
     IF k = 1 THEN
      nme$ = firstnam$
     ELSE
       nme$ = LEFT$(lnam$, k - 1) + ", " + firstnam$
     END IF
1190 PRINT #10, acct$; " "; pin$; TAB(31); LEFT$(nme$, 30);
1192 k = INSTR(tp$, "  "): IF k = 0 OR k = 1 THEN PRINT #10, "": GOTO 1200
1194 IF k = 8 THEN PRINT #10, TAB(65); LEFT$(tp$, 3); "-"; MID$(tp$, 4, 4): GOTO 1200
1196 IF k = 11 THEN PRINT #10, TAB(65); LEFT$(tp$, 3); "-"; MID$(tp$, 4, 3); "-"; MID$(tp$, 6, 4)
1200 PRINT #10, LEFT$(A1$, 28); TAB(30); LEFT$(A2$, 28); TAB(60); cy$
     PRINT #10, st$; TAB(6); zp$; TAB(20); LEFT$(pdesc$, 38); TAB(60); LEFT$(pmap$, 19)
     PRINT #10, LEFT$(psub$, 33); TAB(35); USING "#####"; CVS(plot$);
     PRINT #10, TAB(42); LEFT$(pcounty$, 26); TAB(70); USING "######.#"; CVS(psize$)
     PRINT #10, STRING$(79, "-"): lc = lc + 5
1201 z$ = INKEY$: IF z$ = "" THEN 1202 ELSE IF ASC(z$) = 27 THEN GOSUB 33000: IF ok2continue = 1 THEN 1210 ELSE 1240
1202 IF devspec$ = "S" THEN 1210 ELSE 1220
1210 IF lc >= 18 THEN LOCATE 23, 1: COLOR 11: PRINT "Press any key for next page (ESC-Quit)!"; : fl = 1: GOSUB 62000: IF CF = 1 THEN CLOSE : GOTO 10 ELSE CLS : GOSUB 1900: GOTO 1230
1220 IF lc >= 55 THEN PRINT #10, CHR$(12); : GOSUB 1900
1230 NEXT x: IF devspec$ = "S" THEN LOCATE 22, 1: COLOR 11: PRINT "All done, press any key for menu "; : fl = 1: GOSUB 62000 ELSE PRINT #10, CHR$(12);
1240 CLOSE
1250 GOTO 10

1750 CLOSE 4
     OPEN "tbacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 31
     FIELD 4, 27 AS met$, 4 AS ARREC$
     IF LEN(account$) < 12 THEN account$ = account$ + STRING$(12 - LEN(account$), 32)
     IF LEN(parcel$) < 15 THEN parcel$ = parcel$ + STRING$(15 - LEN(parcel$), 32)
     thisone$ = account$ + parcel$
     IF LEN(thisone$) < 27 THEN thisone$ = thisone$ + STRING$(27 - LEN(thisone$), 32)
1752 IF LOF(4) / 31 = 0 THEN RETURN
1755 HI = LOF(4) / 31: LO = 1: me = HI
1760 FOR XX = 1 TO 16: GET 4, me
1765 IF thisone$ = met$ THEN NG = 1: record = CVS(ARREC$): CLOSE 4: RETURN
1770 IF thisone$ > met$ THEN LO = me ELSE HI = me
1775 me = INT((HI - LO) / 2) + LO
1780 NEXT XX: CLOSE 4
1782 CLOSE 4: NG = 0: RETURN
1785 REM

1900 p = p + 1
     PRINT #10, TAB(27); "Tax Customer Account Listing": PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "####"; p
     PRINT #10, "":
1910 PRINT #10, "Acct. #"; TAB(11); "Parcel #"; TAB(31); "Account Name"; TAB(65); "Phone #"
     PRINT #10, "Address Line 1"; TAB(30); "Address Line 2"; TAB(60); "City"
     PRINT #10, "State  Zip Code"; TAB(20); "Property Desc"; TAB(60); "Map Block Lot"
     PRINT #10, "Subdivision"; TAB(35); "Lot #"; TAB(42); "County"; TAB(70); "  Size"
     PRINT #10, STRING$(79, "="): lc = 7: RETURN
2000 REM
2001 TOT1# = 0: TOT2# = 0: TOT3# = 0
     FOR ll = 0 TO 9: TOTTAX#(ll) = 0: NEXT ll
2010 SCREN$ = "2.02.00": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 32: PRINT "Customer Balance Listing"
2020 LOCATE 4, 1
     LOCATE 4, 1: COLOR 11: PRINT "Do you wish to print the detail report (show receipts) (Y/N) .. "; : fl = 1
     GOSUB 62210: IF NOT (in$ = "Y" OR in$ = "N") THEN 2020
     IF in$ = "N" THEN 2023
     GOTO printdetail
2023 LOCATE 4, 1: PRINT STRING$(79, 32): LOCATE 4, 1
2030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount or (L)ast Name (A,L or ESC-ABORT) ..... "; : fl = 1: GOSUB 62210
2040 IF CF = 1 THEN 10
2050 IF NOT (in$ = "A" OR in$ = "L") THEN 2023
2060 SORTSPEC$ = in$: LOCATE 4, 79: PRINT SORTSPEC$
2070 LOCATE 6, 1
2080 COLOR 11: PRINT "Output Specification: Output to (S)creen or (P)rinter (S,P or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
2090 IF CF = 1 THEN 10
2100 IF NOT (in$ = "S" OR in$ = "P") THEN 2070
2105 devspec$ = in$: LOCATE 6, 77: PRINT devspec$
2108 LOCATE 8, 1: COLOR 11: PRINT "Print (A)ll Customers or (B)alance Customers (A,B or <ESC>-End) .. "; : fl = 1: GOSUB 62210
     IF CF = 1 THEN 10
     IF NOT (in$ = "A" OR in$ = "B") THEN 2108
     type$ = in$
2110 LOCATE 10, 1: COLOR 11: PRINT "Print Past Due by Year (Y/N) .. "; : fl = 1: GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN 2110
     brkdwn$ = in$
     LOCATE 11, 1: COLOR 11: PRINT STRING$(79, "="): LOCATE 13, 23: PRINT "Now Processing : Please be Patient!": COLOR 14
2120 IF devspec$ = "S" THEN DEV$ = "SCRN:" ELSE DEV$ = "lpt1:"
2130 OPEN "O", 10, DEV$: WIDTH #10, 255
2140 GOSUB 60000: GOSUB 61000: IF devspec$ = "S" THEN CLS : GOSUB 2900 ELSE GOSUB 2900
2150 IF LOF(1) / 256 = 0 THEN CLOSE : GOTO 10
2160 IF SORTSPEC$ = "A" THEN OPEN "tbacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 31: FIELD 4, 27 AS met$, 4 AS mrec$: ENDOFFILE = LOF(4) / 22 ELSE OPEN "tblname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24: FIELD 4, 20 AS  _
lastnam$, 4 AS mrec$: ENDOFFILE = LOF(4) / 24
2170 FOR x = 1 TO ENDOFFILE: GET 4, x: IF CVS(mrec$) = 0 THEN 2230
2180 GET 1, CVS(mrec$): GET 3, CVS(mrec$): B = INSTR(acct$, " "): IF B = 0 THEN M$ = LEFT$(acct$, 8) ELSE M$ = LEFT$(acct$, B - 1)
     k = INSTR(lnam$, "  "): IF k = 0 THEN k = 31
     IF k = 1 THEN
      nme$ = firstnam$
     ELSE
       nme$ = LEFT$(lnam$, k - 1) + ", " + firstnam$
     END IF
     pastdue = 0: curbal = 0
     curbal = CVS(curtaxamt$) + CVS(curintamt$) + CVS(curadvamt$) + CVS(curcolamt$) + CVS(curlateamt$)
     FOR ll = 0 TO 9
     pastdue = pastdue + (CVS(PTAXAMT$(ll)) + CVS(PintAMT$(ll)) + CVS(PADVAMT$(ll)) + CVS(PCOLAMT$(ll)) + CVS(PLATEAMT$(ll)))
     NEXT ll
     curbal = INT((curbal * 100) + .5) / 100
     pastdue = INT((pastdue * 100) + .5) / 100
     IF type$ = "A" THEN 2190
     IF curbal > 0 OR pastdue > 0 THEN 2190 ELSE 2230
2190 PRINT #10, acct$ + " " + pin$; TAB(30); LEFT$(nme$, 22);
       PRINT #10, TAB(54); USING "######.##"; curbal;
        PRINT #10, TAB(68); USING "######.##"; pastdue: lc = lc + 1
     IF brkdwn$ = "Y" THEN 2191 ELSE 2195
2191 FOR ll = 0 TO 4: TOTTAX#(ll) = TOTTAX#(ll) + CVS(PTAXAMT$(ll)) + CVS(PintAMT$(ll)) + CVS(PADVAMT$(ll)) + CVS(PCOLAMT$(ll)) + CVS(PLATEAMT$(ll))
      PRINT #10, TAB((ll * 15) + 1); USING "######.##"; (CVS(PTAXAMT$(ll)) + CVS(PintAMT$(ll)) + CVS(PADVAMT$(ll)) + CVS(PCOLAMT$(ll)) + CVS(PLATEAMT$(ll))); : NEXT ll
     FOR ll = 5 TO 9: TOTTAX#(ll) = TOTTAX#(ll) + CVS(PTAXAMT$(ll)) + CVS(PintAMT$(ll)) + CVS(PADVAMT$(ll)) + CVS(PCOLAMT$(ll)) + CVS(PLATEAMT$(ll))
      PRINT #10, TAB(((ll - 5) * 15) + 1); USING "######.##"; (CVS(PTAXAMT$(ll)) + CVS(PintAMT$(ll)) + CVS(PADVAMT$(ll)) + CVS(PCOLAMT$(ll)) + CVS(PLATEAMT$(ll))); : NEXT ll
     PRINT #10, "": lc = lc + 2
2195 TOT1# = TOT1# + curbal: TOT2# = TOT2# + pastdue
2200 IF devspec$ = "S" THEN 2210 ELSE 2220
2210 IF lc >= 15 THEN LOCATE 22, 1: COLOR 11: PRINT "Press any key for next page (ESC-Quit)!"; : fl = 1: GOSUB 62000: IF CF = 1 THEN CLOSE : GOTO 10 ELSE CLS : GOSUB 2900: GOTO 2230
2220 IF lc >= 55 THEN PRINT #10, CHR$(12); : GOSUB 2900
2230 NEXT x: PRINT #10, STRING$(79, "-"):
PRINT #10, "Totals"; TAB(52); USING "########.##"; TOT1#;
PRINT #10, TAB(66); USING "########.##"; TOT2#
IF brkdwn$ = "Y" THEN 2232 ELSE 2235
2232 FOR ll = 0 TO 4: PRINT #10, TAB((ll * 15) + 1); USING "######.##"; TOTTAX#(ll); : NEXT ll
     FOR ll = 5 TO 9: PRINT #10, TAB(((ll - 5) * 15) + 1); USING "######.##"; TOTTAX#(ll); : NEXT ll: PRINT #10, ""
2235 IF devspec$ = "S" THEN LOCATE 22, 1: COLOR 11: PRINT "All done, press any key for menu "; : fl = 1: GOSUB 62000 ELSE PRINT #10, CHR$(12);
2240 CLOSE
2250 GOTO 10
2900 p = p + 1
     PRINT #10, TAB(27); "Tax Customer Balance Listing": PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "####"; p
     PRINT #10, "":
2910 PRINT #10, "Acct  /  Parcel #"; TAB(30); "Account Name"; TAB(54); "Current"; TAB(68); "Previous"
     IF brkdwn$ = "Y" THEN 2920 ELSE PRINT #10, STRING$(79, "="): lc = 2: RETURN
2920 FOR ll = 0 TO 4: PRINT #10, TAB((ll * 15) + 1); "Year: "; VAL(CURRENTYEAR$) - (ll + 1); : NEXT ll
2930 FOR ll = 5 TO 9: PRINT #10, TAB(((ll - 5) * 15) + 1); "Year: "; VAL(CURRENTYEAR$) - (ll + 1); : NEXT ll
2940 PRINT #10, "": PRINT #10, STRING$(79, "="): lc = 4: RETURN

printdetail:
OPEN "O", 10, "lpt1": WIDTH #10, 255

LOCATE 8, 1: COLOR 11
PRINT "This Report Requires Wide Paper or Condensed Print (Ready Y/N) .. "; : fl = 1
GOSUB 62210: IF in$ = "Y" THEN GOTO pr1
CLOSE : GOTO 10
pr1:
GOSUB 60000: GOSUB 61000: GOSUB prhead
IF LOF(1) / 256 = 0 THEN CLOSE : GOTO 10
OPEN "tblname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24: FIELD 4, 20 AS lastnam$, 4 AS mrec$: ENDOFFILE = LOF(4) / 24
FOR x = 1 TO ENDOFFILE: GET 4, x: IF CVS(mrec$) = 0 THEN 2230
GET 1, CVS(mrec$): GET 3, CVS(mrec$): B = INSTR(acct$, " "): IF B = 0 THEN M$ = LEFT$(acct$, 8) ELSE M$ = LEFT$(acct$, B - 1)
 k = INSTR(lnam$, "  ")
 IF k = 0 THEN nme$ = firstnam$ ELSE nme$ = LEFT$(lnam$, k - 1) + ", " + firstnam$
 curbal = CVS(curtaxamt$) + CVS(curintamt$) + CVS(curadvamt$) + CVS(curcolamt$) + CVS(curlateamt$)
 pastdue = 0
 FOR ll = 0 TO 9
   pastdue = pastdue + (CVS(PTAXAMT$(ll)) + CVS(PintAMT$(ll)) + CVS(PADVAMT$(ll)) + CVS(PCOLAMT$(ll)) + CVS(PLATEAMT$(ll)))
 NEXT ll
 curbal = INT((curbal * 100) + .5) / 100
 pastdue = INT((pastdue * 100) + .5) / 100
 IF curbal > 0 OR pastdue > 0 THEN 2990 ELSE 2991
2990 PRINT #10, acct$ + " " + pin$; TAB(30); LEFT$(nme$, 28);
 PRINT #10, TAB(60); USING "######.##"; curbal + pastdue: lc = lc + 1
 yr = VAL(CURRENTYEAR$)
 IF CVS(curtaxamt$) = 0 AND CVS(curintamt$) = 0 AND CVS(curlateamt$) = 0 AND CVS(curadvamt$) = 0 AND CVS(curcolamt$) = 0 THEN GOTO pr01
 PRINT #10, USING "######"; CVS(ctaxnotice$);
  PRINT #10, TAB(12); USING "####"; yr;
  PRINT #10, TAB(35); USING "#####.##"; CVS(curtaxamt$);
  PRINT #10, TAB(50); USING "#####.##"; CVS(curlateamt$);
   PRINT #10, TAB(65); USING "#####.##"; CVS(curintamt$);
    PRINT #10, TAB(80); USING "#####.##"; CVS(curadvamt$);
     PRINT #10, TAB(95); USING "#####.##"; CVS(curcolamt$);
      PRINT #10, TAB(110); USING "#####.##"; (CVS(curtaxamt$) + CVS(curlateamt$) + CVS(curintamt$) + CVS(curadvamt$) + CVS(curcolamt$))
       lc = lc + 1
pr01:
FOR ll = 0 TO 9
IF CVS(PTAXAMT$(ll)) = 0 AND CVS(PintAMT$(ll)) = 0 AND CVS(PADVAMT$(ll)) = 0 AND CVS(PCOLAMT$(ll)) = 0 AND CVS(PLATEAMT$(ll)) = 0 THEN GOTO pr02
PRINT #10, USING "######"; CVS(ptaxnotice$(ll));
 PRINT #10, TAB(12); USING "####"; yr - (ll + 1);
  PRINT #10, TAB(35); USING "#####.##"; CVS(PTAXAMT$(ll));
   PRINT #10, TAB(50); USING "#####.##"; CVS(PLATEAMT$(ll));
    PRINT #10, TAB(65); USING "#####.##"; CVS(PintAMT$(ll));
     PRINT #10, TAB(80); USING "#####.##"; CVS(PADVAMT$(ll));
      PRINT #10, TAB(95); USING "#####.##"; CVS(PCOLAMT$(ll));
       PRINT #10, TAB(110); USING "#####.##"; (CVS(PTAXAMT$(ll)) + CVS(PLATEAMT$(ll)) + CVS(PintAMT$(ll)) + CVS(PADVAMT$(ll)) + CVS(PCOLAMT$(ll)))
        lc = lc + 1

pr02:
NEXT ll

REM add up here based upon year and county
BMTAX(0, 0) = BMTAX(0, 0) + CVS(curtaxamt$)
BMTAX(0, 0) = INT((BMTAX(0, 0) * 100) + .5) / 100

sdTAX(0, 0) = sdTAX(0, 0) + CVS(curlateamt$)
sdTAX(0, 0) = INT((sdTAX(0, 0) * 100) + .5) / 100

INTEREST(0, 0) = INTEREST(0, 0) + CVS(curintamt$)
INTEREST(0, 0) = INT((INTEREST(0, 0) * 100) + .5) / 100

ADVERT(0, 0) = ADVERT(0, 0) + CVS(curadvamt$)
ADVERT(0, 0) = INT((ADVERT(0, 0) * 100) + .5) / 100

COLECT(0, 0) = COLECT(0, 0) + CVS(curcolamt$)
COLECT(0, 0) = INT((COLECT(0, 0) * 100) + .5) / 100

FOR ll1 = 0 TO 9
BMTAX(ll1 + 1, 0) = BMTAX(ll1 + 1, 0) + CVS(PTAXAMT$(ll1))
BMTAX(ll1 + 1, 0) = INT((BMTAX(ll1 + 1, 0) * 100) + .5) / 100
sdTAX(ll1 + 1, 0) = sdTAX(ll1 + 1, 0) + CVS(PLATEAMT$(ll1))
sdTAX(ll1 + 1, 0) = INT((sdTAX(ll1 + 1, 0) * 100) + .5) / 100
INTEREST(ll1 + 1, 0) = INTEREST(ll1 + 1, 0) + CVS(PintAMT$(ll1))
INTEREST(ll1 + 1, 0) = INT((INTEREST(ll1 + 1, 0) * 100) + .5) / 100
ADVERT(ll1 + 1, 0) = ADVERT(ll1 + 1, 0) + CVS(PADVAMT$(ll1))
ADVERT(ll1 + 1, 0) = INT((ADVERT(ll1 + 1, 0) * 100) + .5) / 100
COLECT(ll1 + 1, 0) = COLECT(ll1 + 1, 0) + CVS(PCOLAMT$(ll1))
COLECT(ll1 + 1, 0) = INT((COLECT(ll1 + 1, 0) * 100) + .5) / 100
NEXT ll1

PRINT #10, STRING$(131, "-")
 lc = lc + 1
IF lc >= 50 THEN PRINT #10, CHR$(12); : GOSUB prhead


2991 NEXT x
     GOSUB printtotals: PRINT #10, CHR$(12);
     CLOSE
     GOTO 10
printtotals:
PRINT #10, CHR$(12)
PRINT #10, "Customer Detail Tax Balance Listing": PRINT #10, ""
PRINT #10, "*** GRAND TOTALS"
PRINT #10,
PRINT #10, "TAX YEAR"; TAB(12); "  TAX AMT"; TAB(27); "LATE AMT"; TAB(43); "INTEREST"; TAB(56); "ADVERTISING"; TAB(74); "COLLECT"; TAB(90); "TOTAL AMOUNT"

PRINT #10, STRING$(110, "=")
BMTAX = 0: sdTAX = 0: INTEREST = 0: ADVERT = 0: COLECT = 0
FOR ll = 0 TO 10
A = 0: B = 1: yr = VAL(CURRENTYEAR$)
PRINT #10, yr - ll; TAB(10); USING "$#######,#.##"; BMTAX(ll, A) + BMTAX(ll, B);
 PRINT #10, TAB(25); USING "$#######,#.##"; sdTAX(ll, A) + sdTAX(ll, B);
 PRINT #10, TAB(40); USING "$#######,#.##"; INTEREST(ll, A) + INTEREST(ll, B);
 PRINT #10, TAB(55); USING "$#######,#.##"; ADVERT(ll, A) + ADVERT(ll, B);
 PRINT #10, TAB(70); USING "$#######,#.##"; COLECT(ll, A) + COLECT(ll, B);
 subtot = BMTAX(ll, A) + sdTAX(ll, A) + INTEREST(ll, A) + COLECT(ll, A) + ADVERT(ll, A)
 subtot1 = BMTAX(ll, B) + sdTAX(ll, B) + INTEREST(ll, B) + COLECT(ll, B) + ADVERT(ll, B)
 subtot = INT((subtot * 100) + .5) / 100
 subtot1 = INT((subtot1 * 100) + .5) / 100
 PRINT #10, TAB(90); USING "$#######,#.##"; (subtot + subtot1)
 BMTAX = BMTAX + BMTAX(ll, A) + BMTAX(ll, B)
 sdTAX = sdTAX + sdTAX(ll, A) + sdTAX(ll, B)
 INTEREST = INTEREST + INTEREST(ll, A) + INTEREST(ll, B)
 ADVERT = ADVERT + ADVERT(ll, A) + ADVERT(ll, B)
 COLECT = COLECT + COLECT(ll, A) + COLECT(ll, B)
BMTAX = INT((BMTAX * 100) + .5) / 100
sdTAX = INT((sdTAX * 100) + .5) / 100
INTEREST = INT((INTEREST * 100) + .5) / 100
ADVERT = INT((ADVERT * 100) + .5) / 100
COLECT = INT((COLECT * 100) + .5) / 100


NEXT ll
PRINT #10, STRING$(110, "-")
PRINT #10, "TOTALS"; TAB(10); USING "$#######,#.##"; BMTAX;
 PRINT #10, TAB(25); USING "$#######,#.##"; sdTAX;
  PRINT #10, TAB(40); USING "$#######,#.##"; INTEREST;
   PRINT #10, TAB(55); USING "$#######,#.##"; ADVERT;
    PRINT #10, TAB(70); USING "$#######,#.##"; COLECT;
     PRINT #10, TAB(90); USING "$#######,#.##"; BMTAX + sdTAX + INTEREST + ADVERT + COLECT

 RETURN

prhead:
p = p + 1
PRINT #10, TAB(37); "Detailed Tax Customer Balance Listing": PRINT #10, "Date: "; DATE$; TAB(99); "Page - "; USING "####"; p
PRINT #10, "":
PRINT #10, "Acct No  /    Parcel No"; TAB(30); "Account Name"; TAB(60); "Account Balance"
PRINT #10, "RECEIPT #"; TAB(12); "YEAR"; TAB(35); "   Tax Amt"; TAB(50); " Late Amt"; TAB(65); "Interest Amt"; TAB(80); "Adv Cost"; TAB(95); "Col Cost"; TAB(110); "Total Due"
PRINT #10, STRING$(131, "="): lc = 4: RETURN

3000 REM
3010 SCREN$ = "2.03.00": CLS : GOSUB 63500
3020 COLOR 11: LOCATE 1, 32: PRINT "Customer Account Inquiry";
        LOCATE 3, 1: CLOSE
        COLOR 11: PRINT "Output Specification: Output to (S)creen or (P)rinter (S,P or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
        IF CF = 1 THEN 10
        IF NOT (in$ = "S" OR in$ = "P") THEN 3020
        devspec$ = in$: LOCATE 3, 77: PRINT devspec$
        IF devspec$ = "S" THEN DEV$ = "SCRN:" ELSE DEV$ = "lpt1:"
        OPEN "O", 10, DEV$: WIDTH #10, 255
3030 LOCATE 3, 1: COLOR 11: PRINT STRING$(79, 32): LOCATE 3, 1
        PRINT "Enter Customer Account Number (TAB-Last Name : ESC-Menu) .. "; : fl = 12
        GOSUB 62000:
        IF CF = 1 THEN
         CLOSE : GOTO 10
        END IF
        IF CF = 2 THEN
         GOSUB namesearch
         IF ggo = 1 THEN 3075
         IF ggo = 0 THEN LOCATE 10, 1: PRINT "Account Not Found : Press Any Key to Continue "; : fl = 1: GOSUB 62000: GOTO 3000
        END IF
3040 account$ = in$
3041 COLOR 11
     PRINT "Enter Parcel Number  ... "; : fl = 15: GOSUB 62000
     IF CF = 1 THEN 10
     IF LEN(in$) = 0 THEN parcel$ = "all"
     parcel$ = in$
3065 IF parcel$ = "all" THEN
      GOSUB 3750: IF NG = 0 THEN LOCATE 8, 1: PRINT "Customer Account Not Located, Press any key to try Again!"; : fl = 1: GOSUB 62000: CLOSE : GOTO 3000
      GOTO 3075
     END IF
     GOSUB 1750
     IF NG = 0 THEN LOCATE 8, 1: PRINT "Customer Account Not Located, Press any key to try Again!"; : fl = 1: GOSUB 62000: CLOSE : GOTO 3000
3075 GOSUB 60000: LOCK #1, record : LOCK #3, record
     GET 1, record: GET 3, record: LOCATE 23, 1: PRINT STRING$(79, 32);
     account$ = acct$: parcel$ = pin$
     GOSUB figurebalance

3110 IF DEV$ = "SCRN:" THEN CLS : GOSUB 3900 ELSE GOSUB 3900
3120 TR = CVS(ftr$): GOSUB 61000
3130 IF TR = 0 THEN 3200
3140 GET 2, TR
3145 IF devspec$ = "S" THEN 3146 ELSE 3147
3146 IF lc >= 18 THEN COLOR 11: LOCATE 22, 1: PRINT "Press any key for next page! "; : fl = 1: GOSUB 62000: CLS : GOSUB 3900: GOTO 3150 ELSE 3150
3147 IF lc >= 55 THEN PRINT #10, CHR$(12); : GOSUB 3900
3150 PRINT #10, LEFT$(tdate$, 2); "/"; MID$(tdate$, 3, 2); "/"; MID$(tdate$, 5, 2);
      PRINT #10, TAB(12); USING "####"; CVI(ttaxyr$);
       PRINT #10, TAB(20); tdesc$; TAB(45);
3160 IF CVI(ttype$) = 1 THEN PRINT #10, "Charge"; : GOTO 3170
3162 IF CVI(ttype$) = 2 THEN PRINT #10, "Payment"; : GOTO 3170
3164 IF CVI(ttype$) = 3 THEN PRINT #10, "Misc Chrge"; : GOTO 3170
     IF CVI(ttype$) = 4 THEN PRINT #10, "Release"; : GOTO 3170
     IF CVI(ttype$) = 5 THEN PRINT #10, "Interest"; : GOTO 3170
3168 PRINT #10, "";
3170 PRINT #10, TAB(60); USING "######.##"; CVS(tamt$) + CVS(tdisamt$);
     PRINT #10, TAB(70); USING "######.##"; CVS(tbal$): lc = lc + 1
3172 TR = CVS(TNEXTTR$): GOTO 3130
3200 IF devspec$ = "S" THEN 3210 ELSE 3220
3210 LOCATE 22, 1: COLOR 11: PRINT "That's All, Press any key to Continue "; : fl = 1: GOSUB 62000: CLOSE : GOTO 10
3220 PRINT #10, CHR$(12); : CLOSE : GOTO 10

3750 CLOSE 4
     OPEN "tbacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 16
     FIELD 4, 27 AS met$, 4 AS ARREC$
     IF LEN(account$) < 8 THEN account$ = account$ + STRING$(8 - LEN(account$), 32)
     thisone$ = account$
3752 IF LOF(4) / 22 = 0 THEN RETURN
3755 HI = LOF(4) / 22: LO = 1: me = HI
3760 FOR XX = 1 TO 16: GET 4, me
3765 IF thisone$ = LEFT$(met$, 8) THEN NG = 1: record = CVS(ARREC$): CLOSE 4: RETURN
3770 IF thisone$ > LEFT$(met$, 8) THEN LO = me ELSE HI = me
3775 me = INT((HI - LO) / 2) + LO
3780 NEXT XX: CLOSE 4
3782 CLOSE 4: NG = 0: RETURN
3785 REM

3900 PRINT #10, TAB(26); "Tax Customer - Account History"
3910 PRINT #10, "Date: "; DATE$: PRINT #10, ""
3920 PRINT #10, "Account of ... ";
     B = INSTR(firstnam$, "  "): IF B = 0 THEN B = 30
3930 PRINT #10, LEFT$(firstnam$, B); lnam$;
     PRINT #10, TAB(55); "BALANCE="; USING "$#######,#.##"; curbal# + pastdue#
     PRINT #10, ""
3940 PRINT #10, "  Date"; TAB(10); "Tax Year"; TAB(20); "Description"; TAB(45); "Type"; TAB(60); "Amount"; TAB(70); "  BALANCE"
     PRINT #10, STRING$(79, "="): lc = 6: RETURN
3999 STOP

figurebalance:
     curbal# = 0: pastdue# = 0
     curbal# = (CVS(curtaxamt$) + CVS(curintamt$) + CVS(curadvamt$) + CVS(curcolamt$) + CVS(curlateamt$))
     curbal# = INT((curbal# * 100) + .5) / 100
     FOR ll = 0 TO 9
     pastdue# = pastdue# + (CVS(PTAXAMT$(ll)) + CVS(PintAMT$(ll)) + CVS(PADVAMT$(ll)) + CVS(PCOLAMT$(ll)) + CVS(PLATEAMT$(ll)))
     pastdue# = INT((pastdue# * 100) + .5) / 100
     NEXT ll
     RETURN


namesearch:

LOCATE 3, 1:
COLOR 11: PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : fl = 25
 GOSUB 62000: IF CF = 1 THEN ggo = 0: RETURN
 nn$ = in$
 nn = LEN(nn$)
 CLOSE 4: OPEN "R", 4, "tblname.dat", 24: FIELD 4, 20 AS lastnam$, 4 AS mrec$
 CLOSE 1: GOSUB 60000

FOR ll = 1 TO LOF(4) / 24: GET 4, ll
IF LEFT$(lastnam$, nn) = nn$ THEN GOTO n1
n3:
NEXT ll
ggo = 0: RETURN
n1:
GET 1, CVS(mrec$)
LOCATE 6, 1: COLOR 11
KK = INSTR(firstnam$, "  ")
IF KK = 0 THEN KK = 25
IF KK = 1 THEN nme$ = lnam$ ELSE nme$ = LEFT$(firstnam$, KK) + lnam$
COLOR 10
FOR lll = 6 TO 9: LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
LOCATE 6, 1
PRINT "Full Account Name: "; nme$
PRINT "   Account Number: "; acct$
PRINT "    Parcel Number: "; pin$
ACCTNUMBER$ = acct$
PRINT
n2:
LOCATE 9, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO n2
IF in$ = "Q" THEN ggo = 0: RETURN
IF in$ = "Y" THEN ggo = 1: record = CVS(mrec$): RETURN
IF in$ = "N" THEN GOTO n3
GOTO n2
4000 REM
     totland# = 0: totmv# = 0: totmh# = 0: totfm# = 0: totls# = 0
     totop# = 0: totme# = 0: tothome# = 0: totuse# = 0
4010 SCREN$ = "2.04.00": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 32: PRINT "Customer Valuation Listing"
4020 LOCATE 4, 1
4030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount or (L)ast Name (A,L or ESC-ABORT) ..... "; : fl = 1: GOSUB 62210
4040 IF CF = 1 THEN 10
4050 IF NOT (in$ = "A" OR in$ = "L") THEN 4020
4060 SORTSPEC$ = in$: LOCATE 4, 79: PRINT SORTSPEC$
4070 LOCATE 6, 1
4080 COLOR 11: PRINT "Output Specification: Output to (S)creen or (P)rinter (S,P or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
4090 IF CF = 1 THEN 10
4100 IF NOT (in$ = "S" OR in$ = "P") THEN 4070
4105 devspec$ = in$: LOCATE 6, 77: PRINT devspec$
4110 LOCATE 7, 1: COLOR 11: PRINT STRING$(79, "="): LOCATE 13, 23: PRINT "Now Processing : Please be Patient!": COLOR 14
4120 IF devspec$ = "S" THEN DEV$ = "SCRN:" ELSE DEV$ = "lpt1:"
4130 OPEN "O", 10, DEV$: WIDTH #10, 255
4140 GOSUB 60000: IF devspec$ = "S" THEN CLS : GOSUB 4900 ELSE GOSUB 4900
4150 IF LOF(1) / 256 = 0 THEN CLOSE : GOTO 10
4160 IF SORTSPEC$ = "A" THEN OPEN "tbacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 31: FIELD 4, 27 AS met$, 4 AS mrec$: ENDOFFILE = LOF(4) / 22 ELSE OPEN "tblname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24: FIELD 4, 20 AS  _
lastnam$, 4 AS mrec$: ENDOFFILE = LOF(4) / 24
4170 FOR x = 1 TO ENDOFFILE: GET 4, x: IF CVS(mrec$) = 0 THEN 4230
4180 GET 1, CVS(mrec$): GET 3, CVS(mrec$)
     k = INSTR(lnam$, "  "): IF k = 0 THEN k = 31
     IF k = 1 THEN
      nme$ = firstnam$
     ELSE
       nme$ = LEFT$(lnam$, k - 1) + ", " + firstnam$
     END IF

4190 PRINT #10, acct$; " "; pin$; TAB(30); LEFT$(nme$, 49)
     PRINT #10, USING "######,#"; CVD(pland$);
      PRINT #10, TAB(15); USING "######"; CVD(mv$);
       PRINT #10, TAB(23); USING "######"; CVD(mh$);
        PRINT #10, TAB(31); USING "######"; CVD(fm$);
        PRINT #10, TAB(39); USING "######"; CVD(ls$);
         PRINT #10, TAB(47); USING "######"; CVS(op$);
          PRINT #10, TAB(55); USING "######"; CVS(me$);
          
     PRINT #10, TAB(65); USING "######,#"; CVS(HOMEEXP$);
          PRINT #10, "": lc = lc + 3
4191 totland# = totland# + CVD(pland$): totland# = INT((totland# * 100) + .5) / 100
     totmv# = totmv# + CVD(mv$): totmv# = INT((totmv# * 100) + .5) / 100
     totmh# = totmh# + CVD(mh$): totmh# = INT((totmh# * 100) + .5) / 100
     totfm# = totfm# + CVD(fm$): totfm# = INT((totfm# * 100) + .5) / 100
     totls# = totls# + CVD(ls$): totls# = INT((totls# * 100) + .5) / 100
     totop# = totop# + CVS(op$): totop# = INT((totop# * 100) + .5) / 100
     totme# = totme# + CVS(me$): totme# = INT((totme# * 100) + .5) / 100
     TOTEXP# = TOTEXP# + CVS(HOMEEXP$): TOTEXP# = INT((TOTEXP# * 100) + .5) / 100

4200 IF devspec$ = "S" THEN 4210 ELSE 4220
4210 IF lc >= 15 THEN COLOR 11: LOCATE 22, 1: PRINT "Press any key for next page (ESC-Quit)!"; : fl = 1: GOSUB 62000: IF CF = 1 THEN CLOSE : GOTO 10 ELSE CLS : GOSUB 4900: GOTO 4230
4220 IF lc >= 55 THEN PRINT #10, CHR$(12); : GOSUB 4900
4230 NEXT x: PRINT #10, STRING$(79, "-"):
     IF devspec$ = "S" THEN
        LOCATE 23, 1: PRINT "Press Any Key For Totals Page"; : fl = 1
        GOSUB 62000: CLS : GOSUB 4500
     END IF
     IF devspec$ = "P" THEN
        PRINT #10, CHR$(12); : GOSUB 4500
     END IF
     IF devspec$ = "S" THEN LOCATE 22, 1: COLOR 11: PRINT "All done, press any key for menu "; : fl = 1: GOSUB 62000 ELSE PRINT #10, CHR$(12);
4240 CLOSE
4250 GOTO 10
4500 PRINT #10, "Tax Billing System : Valuation Summary Page"
     PRINT #10, ""
     PRINT #10, "Total Real Value ........... "; USING "$$#########,#"; totland#
     PRINT #10, "- Total Land Value ......... "; USING "$$#########,#"; totland#
     PRINT #10, ""
4510 PRINT #10, "Total Personal Value ....... "; USING "$$#########,#"; totmv# + totmh# + totfm# + totls# + totop# + totme#
     PRINT #10, "- Total Motor Veh Value .... "; USING "$$#########,#"; totmv#
     PRINT #10, "- Total Mobile Home Value .. "; USING "$$#########,#"; totmh#
     PRINT #10, "- Total Farm Machine Value.. "; USING "$$#########,#"; totfm#
     PRINT #10, "- Total Live Stock Value ... "; USING "$$#########,#"; totls#
     PRINT #10, "- Total Other Pers Value ... "; USING "$$#########,#"; totop#
     PRINT #10, "- Total Machine/Equip Value. "; USING "$$#########,#"; totme#
     PRINT #10, ""
     PRINT #10, "** Total Valuation ......... "; USING "$$#########,#"; totland# + totmv# + totmh# + totfm# + totls# + totop# + totme#
     PRINT #10, ""
4520 PRINT #10, "Total Exemptions ........... "; USING "$$#########,#"; TOTEXP#
     PRINT #10, "- Sr Citizen Exemptions .... "; USING "$$#########,#"; TOTEXP#
     PRINT #10, ""
     PRINT #10, "** Net Total Valuation ..... "; USING "$$#########,#"; totland# + totmv# + totmh# + totfm# + totls# + totop# + totme# - TOTEXP#
4540 RETURN
4900 p = p + 1
     PRINT #10, TAB(27); "Tax Customer Valuation Listing"
     PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "####"; p
     PRINT #10, ""
4910 PRINT #10, " Acct # : Parcel"; TAB(30); "Account Name"
     PRINT #10, "Real Valuation"; TAB(19); "-------- Personal Valuations --------"
           PRINT #10, "    Land"; TAB(19); "MV      MH      FM      LS     OP     ME"; TAB(65); "Exemptions  "
     PRINT #10, STRING$(79, "="): lc = 2: RETURN

4920 REM
5000 REM
5001 lc = 0
5010 SCREN$ = "2.05.00": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 32: PRINT "Customer Mailing Labels"
5020 LOCATE 4, 1
5030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount or (L)ast Name (A,L or ESC-ABORT) ..... "; : fl = 1: GOSUB 62210
5040 IF CF = 1 THEN 10
5050 IF NOT (in$ = "A" OR in$ = "L") THEN 5020
5060 SORTSPEC$ = in$: LOCATE 4, 79: PRINT SORTSPEC$
5070 LOCATE 6, 1
5080 COLOR 11: PRINT "Output Specification: Output to (S)creen or (P)rinter (S,P or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
5090 IF CF = 1 THEN 10
5100 IF NOT (in$ = "S" OR in$ = "P") THEN 5070
5105 devspec$ = in$: LOCATE 6, 77: PRINT devspec$
5110 COLOR 11: LOCATE 7, 1
     COLOR 11: PRINT "Show Account & Parcel Number on Label (Y/N) .. "; : fl = 1
     GOSUB 62210: IF NOT (in$ = "Y" OR in$ = "N") THEN 5110
     onlabel$ = in$
     COLOR 11: LOCATE 12, 1
     PRINT STRING$(79, "="): LOCATE 13, 23: PRINT "Now Processing : Please be Patient!": COLOR 14
5120 IF devspec$ = "S" THEN DEV$ = "SCRN:" ELSE DEV$ = "lpt1:"
5130 OPEN "O", 10, DEV$: WIDTH #10, 255
5140 GOSUB 60000: IF devspec$ = "S" THEN CLS
5150 IF LOF(1) / 512 = 0 THEN CLOSE : GOTO 10
5160 IF SORTSPEC$ = "A" THEN OPEN "tbacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 31: FIELD 4, 27 AS met$, 4 AS mrec$: ENDOFFILE = LOF(4) / 22 ELSE OPEN "tblname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24: FIELD 4, 20 AS  _
lastnam$, 4 AS mrec$: ENDOFFILE = LOF(4) / 24
5170 FOR x = 1 TO ENDOFFILE: GET 4, x
5180 GET 1, CVS(mrec$)
5182 B = INSTR(firstnam$, "  "): IF B = 0 THEN nme$ = lnam$ ELSE nme$ = LEFT$(firstnam$, B) + lnam$
     IF onlabel$ = "N" THEN PRINT #10, "": GOTO 5185
     PRINT #10, acct$; " "; pin$
5185 PRINT #10, LEFT$(nme$, 34)
5186 PRINT #10, A1$
5188 IF LEFT$(A2$, 1) = " " THEN EXTRA = 1 ELSE EXTRA = 0
5190 IF EXTRA = 0 THEN PRINT #10, A2$
5192 PRINT #10, cy$; st$; " "; zp$
5194 IF EXTRA = 1 THEN PRINT #10, " "
5198 lc = lc + 6
5200 IF devspec$ = "S" THEN 5210 ELSE 5220
5210 IF lc >= 15 THEN COLOR 11: LOCATE 22, 1: PRINT "Press any key for next page (ESC-Quit)!"; : fl = 1: GOSUB 62000: IF CF = 1 THEN CLOSE : GOTO 10 ELSE CLS : lc = 0: GOTO 5230
5220 PRINT #10, ""
5230 NEXT x: IF devspec$ = "S" THEN COLOR 11: LOCATE 22, 1: PRINT "All done, press any key for menu "; : fl = 1: GOSUB 62000 ELSE PRINT #10, CHR$(12);
5240 CLOSE
5250 GOTO 10


6000 REM
6001 lc = 0: totloss# = 0: totvalue# = 0
6010 SCREN$ = "2.06.00": CLS : GOSUB 63500: COLOR 11: LOCATE 1, 32: PRINT "Old Age Exemptions"
6020 LOCATE 4, 1
6030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount or (L)ast Name (A,L or ESC-ABORT) ..... "; : fl = 1: GOSUB 62210
6040 IF CF = 1 THEN 10
6050 IF NOT (in$ = "A" OR in$ = "L") THEN 6020
6060 SORTSPEC$ = in$: LOCATE 4, 79: PRINT SORTSPEC$
6070 LOCATE 6, 1
6080 COLOR 11: PRINT "Output Specification: Output to (S)creen or (P)rinter (S,P or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
6090 IF CF = 1 THEN 10
6100 IF NOT (in$ = "S" OR in$ = "P") THEN 6070
6105 devspec$ = in$: LOCATE 6, 77: PRINT devspec$
6110 COLOR 11: LOCATE 8, 1:
     PRINT "Enter Tax Rate (Ex: .48 for 48 cents per hundred) .. "; : fl = -5: GOSUB 62000
     IF CF = 1 THEN 10
     IF VAL(in$) = 0 THEN 6110
     taxrate = VAL(in$)
     PRINT STRING$(79, "="): LOCATE 13, 23: PRINT "Now Processing : Please be Patient!": COLOR 14
6120 IF devspec$ = "S" THEN DEV$ = "SCRN:" ELSE DEV$ = "lpt1:"
6130 OPEN "O", 10, DEV$: WIDTH #10, 255
6140 GOSUB 60000: IF devspec$ = "S" THEN CLS : GOSUB 6900
6150 IF LOF(1) / 256 = 0 THEN CLOSE : GOTO 10
6160 IF SORTSPEC$ = "A" THEN
      OPEN "tbacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 31
      FIELD 4, 27 AS met$, 4 AS mrec$: ENDOFFILE = LOF(4) / 31
      ELSE
      OPEN "tblname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
      FIELD 4, 20 AS lastnam$, 4 AS mrec$: ENDOFFILE = LOF(4) / 24
     END IF
6170 FOR x = 1 TO ENDOFFILE: GET 4, x
6180 GET 1, CVS(mrec$): GET 3, CVS(mrec$): B = INSTR(acct$, " "): IF B = 0 THEN M$ = acct$ ELSE M$ = LEFT$(acct$, B - 1)
6181 IF LEFT$(firstnam$, 1) = " " THEN nme$ = lnam$: GOTO 6185
6182 B = INSTR(firstnam$, "  "): IF B = 0 THEN B = 20
6184 nme$ = LEFT$(firstnam$, B) + lnam$
6185 IF CVS(HOMEEXP$) > 0 THEN GOSUB printline
6190 NEXT x
     PRINT #10, STRING$(79, "-")
     PRINT #10, TAB(55); USING "$######,#.##"; totvalue#;
      PRINT #10, TAB(70); USING "$####,#.##"; totloss#
     IF devspec$ = "S" THEN COLOR 11: LOCATE 22, 1: PRINT "All done, press any key for menu "; : fl = 1: GOSUB 62000 ELSE PRINT #10, CHR$(12);
     CLOSE : GOTO 10
printline:
IF devspec$ = "S" THEN GOTO p1 ELSE GOTO p2
p1:
 IF lc >= 15 THEN COLOR 11: LOCATE 22, 1: PRINT "Press any key for next page (ESC-Quit)!"; : fl = 1: GOSUB 62000: IF CF = 1 THEN CLOSE : GOTO 10 ELSE CLS : GOSUB 6900: GOTO p3
p2:
 IF lc >= 55 THEN PRINT #10, CHR$(12); : GOSUB 6900
p3:
  PRINT #10, TAB(3); ss$; TAB(20); nme$
  PRINT #10, TAB(20); A1$; TAB(55); USING "$######,#.##"; CVS(HOMEEXP$);
  totvalue# = totvalue# + CVS(HOMEEXP$)
  PRINT #10, TAB(70); USING "$#####.##"; (CVS(HOMEEXP$) / 100) * taxrate
  totloss# = totloss# + (CVS(HOMEEXP$) / 100) * taxrate
  PRINT #10, TAB(20); cy$; " "; st$; " "; zp$: lc = lc + 4: PRINT #10, "": RETURN



6900 p = p + 1: PRINT #10, TAB(10); "FORM AV-22C"; TAB(60); "Page "; p
     PRINT #10, "": PRINT #10, ""
     PRINT #10, "Social Security #"; TAB(20); "Name/Address"; TAB(57); "Exempt Amt"; TAB(71); "Tax Loss"
     PRINT #10, STRING$(79, "="): lc = 3: RETURN

7000 REM
7001 TOT1# = 0: TOT2# = 0: TOT3# = 0
     FOR ll = 0 TO 9: TOTTAX#(ll) = 0: NEXT ll
7010 SCREN$ = "2.02.00": CLS : GOSUB 63500: COLOR 11
     LOCATE 1, 25: PRINT "Customer Balance Listing : Specific Year"
7020 LOCATE 4, 1
7023 LOCATE 4, 1: PRINT STRING$(79, 32): LOCATE 4, 1
7030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount or (L)ast Name (A,L or ESC-ABORT) ..... "; : fl = 1: GOSUB 62210
7040 IF CF = 1 THEN 10
7050 IF NOT (in$ = "A" OR in$ = "L") THEN 7023
7060 SORTSPEC$ = in$: LOCATE 4, 79: PRINT SORTSPEC$
7070 LOCATE 6, 1
7080 COLOR 11: PRINT "Output Specification: Output to (S)creen or (P)rinter (S,P or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
7090 IF CF = 1 THEN 10
7100 IF NOT (in$ = "S" OR in$ = "P") THEN 7070
7105 devspec$ = in$: LOCATE 6, 77: PRINT devspec$
7108 LOCATE 8, 1: COLOR 11: PRINT "Enter Tax Year to Print (ESC-End) (yyyy) ... "; : fl = -4: GOSUB 62000
     IF CF = 1 THEN 10
     IF LEN(in$) <> 4 THEN 7108
     taxyear$ = in$
    LOCATE 11, 1
    COLOR 11: PRINT STRING$(79, "="): LOCATE 13, 23
     PRINT "Now Processing : Please be Patient!": COLOR 14

          cyr = VAL(CURRENTYEAR$)
          pyr = VAL(taxyear$)
          IF pyr > cyr THEN
           CLOSE : LOCATE 19, 1: PRINT "Year Specified Exceeds Current Year : Press any key to Continue "; :
           fl = 1: GOSUB 62000: GOTO 7000
          END IF
          fyr = cyr - pyr
7120 IF devspec$ = "S" THEN DEV$ = "SCRN:" ELSE DEV$ = "lpt1:"
7130 OPEN "O", 10, DEV$: WIDTH #10, 255
7140 GOSUB 60000: GOSUB 61000: IF devspec$ = "S" THEN CLS : GOSUB 7900 ELSE GOSUB 7900
7150 IF LOF(1) / 256 = 0 THEN CLOSE : GOTO 10
7160 IF SORTSPEC$ = "A" THEN OPEN "tbacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 31: FIELD 4, 27 AS met$, 4 AS mrec$: ENDOFFILE = LOF(4) / 22 ELSE OPEN "tblname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24: FIELD 4, 20 AS  _
lastnam$, 4 AS mrec$: ENDOFFILE = LOF(4) / 24
7170 FOR x = 1 TO ENDOFFILE: GET 4, x: IF CVS(mrec$) = 0 THEN 7230
7180 GET 1, CVS(mrec$): GET 3, CVS(mrec$): B = INSTR(acct$, " "): IF B = 0 THEN M$ = LEFT$(acct$, 8) ELSE M$ = LEFT$(acct$, B - 1)
     k = INSTR(lnam$, "  "): IF k = 0 THEN k = 31
     IF k = 1 THEN
      nme$ = firstnam$
     ELSE
       nme$ = LEFT$(lnam$, k - 1) + ", " + firstnam$
     END IF
      pastdue = 0: curbal = 0
      curbal = CVS(curtaxamt$) + CVS(curintamt$) + CVS(curadvamt$) + CVS(curcolamt$) + CVS(curlateamt$)
     FOR ll = 0 TO 9
      pastdue = pastdue + (CVS(PTAXAMT$(ll)) + CVS(PintAMT$(ll)) + CVS(PADVAMT$(ll)) + CVS(PCOLAMT$(ll)) + CVS(PLATEAMT$(ll)))
     NEXT ll
     curbal = INT((curbal * 100) + .5) / 100
     pastdue = INT((pastdue * 100) + .5) / 100
IF fyr = 0 THEN
 curTAX = CVS(curtaxamt$) + CVS(curadvamt$) + CVS(curcolamt$) + CVS(curlateamt$)
 curint = CVS(curintamt$)
 IF curTAX + curint > 0 THEN GOTO 7190 ELSE 7230
END IF
ll = fyr - 1
IF ll > 9 THEN ll = 9
curTAX = (CVS(PTAXAMT$(ll)) + CVS(PADVAMT$(ll)) + CVS(PCOLAMT$(ll)) + CVS(PLATEAMT$(ll)))
curint = CVS(PintAMT$(ll))
IF curTAX + curint > 0 THEN GOTO 7190 ELSE 7230
7190 PRINT #10, acct$; " "; pin$; TAB(33); LEFT$(nme$, 20);
       PRINT #10, TAB(54); USING "######.##"; curTAX;
        PRINT #10, TAB(68); USING "######.##"; curint: lc = lc + 1
7195 TOT1# = TOT1# + curTAX: TOT2# = TOT2# + curint
7200 IF devspec$ = "S" THEN 7210 ELSE 7220
7210 IF lc >= 15 THEN LOCATE 22, 1: COLOR 11: PRINT "Press any key for next page (ESC-Quit)!"; : fl = 1: GOSUB 62000: IF CF = 1 THEN CLOSE : GOTO 10 ELSE CLS : GOSUB 7900: GOTO 7230
7220 IF lc >= 55 THEN PRINT #10, CHR$(12); : GOSUB 7900
7230 NEXT x: PRINT #10, STRING$(79, "-"):
PRINT #10, "Totals"; TAB(52); USING "########.##"; TOT1#;
PRINT #10, TAB(66); USING "########.##"; TOT2#
7235 IF devspec$ = "S" THEN LOCATE 22, 1: COLOR 11: PRINT "All done, press any key for menu "; : fl = 1: GOSUB 62000 ELSE PRINT #10, CHR$(12);
7240 CLOSE
7250 GOTO 10
7900 p = p + 1
     PRINT #10, TAB(21); "Tax Customer Balance Listing : By Year": PRINT #10, "Date: "; DATE$; TAB(68); "Page - "; USING "####"; p
     PRINT #10, "Tax Year: "; taxyear$
     PRINT #10, ""
7910 PRINT #10, "Acct #       Parcel #"; TAB(33); "Account Name"; TAB(55); "Tax Due"; TAB(65); "Interest Due"
7940 PRINT #10, STRING$(79, "="): lc = 4: RETURN
8000 TOT1# = 0: TOT2# = 0: TOT3# = 0
     FOR ll = 0 TO 9: TOTTAX#(ll) = 0: NEXT ll
8010 SCREN$ = "2.02.00": CLS : GOSUB 63500
      COLOR 11: LOCATE 1, 32: PRINT "Advertising List"
8020 LOCATE 4, 1: COLOR 11
8030 COLOR 11: PRINT "Sort Specification: Sort by (A)ccount or (L)ast Name (A,L or ESC-ABORT) ..... "; : fl = 1: GOSUB 62210
8040 IF CF = 1 THEN 10
8050 IF NOT (in$ = "A" OR in$ = "L") THEN 8020
8060 SORTSPEC$ = in$: LOCATE 4, 79: PRINT SORTSPEC$
8070 LOCATE 6, 1
8080 COLOR 11: PRINT "Output Specification: Output to (S)creen or (P)rinter (S,P or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
8090 IF CF = 1 THEN 10
8100 IF NOT (in$ = "S" OR in$ = "P") THEN 8070
8105 devspec$ = in$: LOCATE 6, 77: PRINT devspec$
8110 LOCATE 8, 1: COLOR 11
     PRINT "Enter Amount to Charge (if any) Each Account as Printed .. "; : fl = -8: GOSUB 62000
     IF CF = 1 THEN CLOSE : RUN "tbcustre"
     adamt = VAL(in$)
     LOCATE 8, 57: PRINT USING "#####.##"; adamt
8115 LOCATE 10, 1: COLOR 11
     PRINT "Correct (Y/N) .. "; : fl = 1: GOSUB 62210
     IF NOT (in$ = "Y" OR in$ = "N") THEN 8115
     IF in$ = "N" THEN 8000

     LOCATE 12, 1: COLOR 11:
     PRINT STRING$(79, "="): LOCATE 13, 23: PRINT "Now Processing : Please be Patient!": COLOR 14
8120 IF devspec$ = "S" THEN DEV$ = "SCRN:" ELSE DEV$ = "lpt1:"
8130 OPEN "O", 10, DEV$: WIDTH #10, 255
8140 GOSUB 60000: GOSUB 61000: IF devspec$ = "S" THEN CLS : GOSUB 8900 ELSE GOSUB 8900
8150 IF LOF(1) / 256 = 0 THEN CLOSE : GOTO 10
8160 IF SORTSPEC$ = "A" THEN OPEN "tbacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 31: FIELD 4, 27 AS met$, 4 AS mrec$: ENDOFFILE = LOF(4) / 22 ELSE OPEN "tblname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24: FIELD 4, 20 AS  _
lastnam$, 4 AS mrec$: ENDOFFILE = LOF(4) / 24
8170 FOR x = 1 TO ENDOFFILE: GET 4, x: IF CVS(mrec$) = 0 THEN 8230
8180 GET 1, CVS(mrec$): GET 3, CVS(mrec$): B = INSTR(acct$, " "): IF B = 0 THEN M$ = LEFT$(acct$, 8) ELSE M$ = LEFT$(acct$, B - 1)
     acctrecord = CVS(mrec$)
     k = INSTR(lnam$, "  "): IF k = 0 THEN k = 31
     IF k = 1 THEN
      nme$ = firstnam$
     ELSE
       nme$ = LEFT$(lnam$, k - 1) + ", " + firstnam$
     END IF
     pastdue = 0: curbal = 0

     curbal = CVS(curtaxamt$) + CVS(curintamt$) + CVS(curadvamt$) + CVS(curcolamt$) + CVS(curlateamt$)
     curbal = INT((curbal * 100) + .5) / 100
     IF curbal = 0 THEN 8230
8190 PRINT #10, acct$; " "; pin$; TAB(30); LEFT$(nme$, 23);
       PRINT #10, TAB(55); LEFT$(pdesc$, 14);
       PRINT #10, TAB(71); USING "######.##"; curbal: lc = lc + 1
         IF adamt > 0 THEN
          GOSUB chargethem
          totcharge# = totcharge# + adamt
         END IF
8200 TOT1# = TOT1# + curbal
     IF devspec$ = "S" THEN 8210 ELSE 8220
8210 IF lc >= 15 THEN LOCATE 22, 1: COLOR 11: PRINT "Press any key for next page (ESC-Quit)!"; : fl = 1: GOSUB 62000: IF CF = 1 THEN CLOSE : GOTO 10 ELSE CLS : GOSUB 8900: GOTO 8230
8220 IF lc >= 55 THEN PRINT #10, CHR$(12); : GOSUB 8900
8230 NEXT x:
8235 IF devspec$ = "S" THEN
        IF totcharge# > 0 THEN
         PRINT #10, "Total Charged for Advertising ......... "; TAB(52); USING "$$#######,#.##"; totcharge#
        END IF
       LOCATE 22, 1
       COLOR 11: PRINT "All done, press any key for menu ";
       fl = 1: GOSUB 62000
     ELSE
     IF totcharge# > 0 THEN
      PRINT #10, "Total Charged for Advertising ......... "; TAB(52); USING "$$#######,#.##"; totcharge#
     END IF
     PRINT #10, CHR$(12);
     END IF
8240 CLOSE
8250 GOTO 10

chargethem:
 LSET curadvamt$ = MKS$(CVS(curadvamt$) + adamt)
 PUT 1, acctrecord
 GOSUB figurebalance
 newbalance = curbal# + pastdue#
 newbalance = INT((newbalance * 100) + .5) / 100

GOSUB 61000
LSET tdate$ = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 2)
LSET tamt$ = MKS$(adamt): LSET ttaxyr$ = MKI$(tyr)
LSET tdesc$ = "Advertising Charge": LSET ttype$ = MKI$(3)
LSET taprin$ = MKS$(0): LSET taint$ = MKS$(0): LSET taadv$ = MKS$(adamt)
LSET tacol$ = MKS$(0): LSET talate$ = MKS$(0): LSET tmrec$ = MKS$(acctrecord)
LSET TNEXTTR$ = MKS$(0): LSET tdisamt$ = MKS$(0)
LSET tbal$ = MKS$(newbalance)
PUT 2, LOF(2) / 96 + 1: Trecord = LOF(2) / 96

IF CVS(ftr$) = 0 THEN LSET ftr$ = MKS$(Trecord): LSET ltr$ = MKS$(Trecord): PUT 1, acctrecord: CLOSE 1: CLOSE 2: RETURN
 prev = CVS(ltr$): LSET ltr$ = MKS$(Trecord): PUT 1, acctrecord
 GET 2, prev: LSET TNEXTTR$ = MKS$(Trecord): PUT 2, prev: CLOSE 2: RETURN

8900 p = p + 1
     PRINT #10, TAB(24); "Tax Customer Advertising Listing": PRINT #10, "Report Date: "; DATE$; TAB(68); "Page - "; USING "####"; p
     PRINT #10, "":
8910 PRINT #10, "Acct #        Parcel #"; TAB(30); "Account Name"; TAB(55); "Prop Desc"; TAB(72); "Cur. Tax"
8920 PRINT #10, STRING$(79, "="): lc = 4: RETURN
9000 CLS : RUN "tbdelnot"

10000 REM payment register
 SCREN$ = "2.10.00": CLS :
 GOSUB 63500: COLOR 11
 LOCATE 1, 32: PRINT "Cash Receipts Recap"
10010 LOCATE 4, 1: COLOR 11
      PRINT "Enter Beginning Date (mmddyy) <ESC>-Menu ... "; : fl = -6: GOSUB 62000
      IF CF = 1 THEN 10
      IF LEN(in$) <> 6 THEN 10
      IF VAL(LEFT$(in$, 2)) < 1 OR VAL(LEFT$(in$, 2)) > 12 THEN 10010
      IF VAL(MID$(in$, 3, 2)) < 1 OR VAL(MID$(in$, 3, 2)) > 31 THEN 10010
      BDATE$ = in$
      begdate = VAL((RIGHT$(in$, 2) + LEFT$(in$, 2) + MID$(in$, 3, 2)))
10020 LOCATE 5, 1: COLOR 11
      PRINT "Enter Ending Date (mmddyy) <ESC>-Menu ...... "; : fl = -6: GOSUB 62000
      IF CF = 1 THEN 10
      IF LEN(in$) <> 6 THEN 10
      IF VAL(LEFT$(in$, 2)) < 1 OR VAL(LEFT$(in$, 2)) > 12 THEN 10020
      IF VAL(MID$(in$, 3, 2)) < 1 OR VAL(MID$(in$, 3, 2)) > 31 THEN 10020
      EDATE$ = in$
      enddate = VAL((RIGHT$(in$, 2) + LEFT$(in$, 2) + MID$(in$, 3, 2)))
10025 LOCATE 7, 1: COLOR 11
      PRINT "Print (D)etail or (S)ummary Total Only (ESC-Menu) .. "; : fl = 1
      GOSUB 62210: IF CF = 1 THEN 10
      IF NOT (in$ = "D" OR in$ = "S") THEN 10025
      reptype$ = in$

10027 COLOR 11: LOCATE 9, 1
      PRINT "Output Specification: Output to (S)creen or (P)rinter (S,P or ESC-ABORT) .. "; : fl = 1: GOSUB 62210
      IF CF = 1 THEN 10
      IF NOT (in$ = "S" OR in$ = "P") THEN 10027
      devspec$ = in$: LOCATE 9, 77: PRINT devspec$
      IF devspec$ = "S" THEN DEV$ = "SCRN:" ELSE DEV$ = "lpt1:"
      OPEN "O", 10, DEV$: WIDTH #10, 255
10030 GOSUB 61500: GOSUB 60000
      curyear = VAL(CURRENTYEAR$)
      IF devspec$ = "S" THEN CLS : GOSUB CRHEADING ELSE GOSUB CRHEADING
      FOR x = LOF(12) / 64 TO 1 STEP -1
      GET 12, x
      tdate = VAL((RIGHT$(rpdate$, 2) + LEFT$(rpdate$, 2) + MID$(rpdate$, 3, 2)))
      REM IF tdate < begdate THEN GOTO 10300
10035 IF tdate >= begdate AND tdate <= enddate THEN 10040 ELSE 10200
10040 IF reptype$ = "D" THEN
       GOTO 10050
       ELSE
       payment = CVS(rpbmamt$) + CVS(rpsdamt$) + CVS(rpint$) + CVS(rpadv$) + CVS(rpcol$)
       GOTO 10100
       END IF
10050 account = CVS(rprec$)
      IF account < 1 OR account > LOF(1) / 512 THEN
       acctname$ = ""
       ELSE
       GET 1, account
      k = INSTR(lnam$, "  "): IF k = 0 THEN nme$ = firstnam$ ELSE nme$ = LEFT$(lnam$, k) + ", " + firstnam$
      END IF
10051      IF devspec$ = "S" AND lc >= 18 THEN
       LOCATE 23, 1: PRINT "PRESS ANY KEY FOR NEXT PAGE! (ESC-END) .. "; : fl = 1
       GOSUB 62000: IF CF = 1 THEN CLOSE : GOTO 10
      CLS : GOSUB CRHEADING
      END IF
      IF devspec$ = "P" AND lc >= 55 THEN
        PRINT #10, CHR$(12);
       GOSUB CRHEADING
      END IF
       payment = CVS(rpbmamt$) + CVS(rpsdamt$) + CVS(rpint$) + CVS(rpadv$) + CVS(rpcol$)
      PRINT #10, LEFT$(rpdate$, 2); "/"; MID$(rpdate$, 3, 2); "/"; MID$(rpdate$, 5, 2); TAB(12); LEFT$(nme$, 25);
       PRINT #10, TAB(40); USING "######"; CVS(rprct$);
        PRINT #10, TAB(48); rptype$;
        PRINT #10, TAB(60); USING "$$#####,#.##"; payment
     PRINT #10, TAB(12); USING "#####.##"; CVS(rpbmamt$);
      PRINT #10, TAB(22); USING "#####.##"; CVS(rpsdamt$);
       PRINT #10, TAB(32); USING "#####.##"; CVS(rpint$);
        PRINT #10, TAB(42); USING "#####.##"; CVS(rpadv$);
         PRINT #10, TAB(52); USING "#####.##"; CVS(rpcol$);
         PRINT #10, TAB(62); USING "#####.##"; CVS(rpover$);
         PRINT #10, TAB(77); rpcty$
         PRINT #10, STRING$(79, "-"): lc = lc + 3
10100 IF rptype$ = "P" THEN ptype = 1 ELSE ptype = 2
      cty = 1: REM all in one county
      year = CVI(rpyear$)
      IF year = 0 THEN year = curyear: LSET rpyear$ = MKI$(year)
      tyear = curyear - year
      yrdesc(tyear) = CVI(rpyear$)

    
        totpay#(ptype, cty, tyear) = totpay#(ptype, cty, tyear) + (payment)
        totprin#(ptype, cty, tyear) = totprin#(ptype, cty, tyear) + CVS(rpbmamt$)
        totsd#(ptype, cty, tyear) = totsd#(ptype, cty, tyear) + CVS(rpsdamt$)
        totint#(ptype, cty, tyear) = totint#(ptype, cty, tyear) + CVS(rpint$)
        totadv#(ptype, cty, tyear) = totadv#(ptype, cty, tyear) + CVS(rpadv$)
        totcol#(ptype, cty, tyear) = totcol#(ptype, cty, tyear) + CVS(rpcol$)
        TOTOVER# = TOTOVER# + CVS(rpover$)
10200 NEXT x
10300 IF devspec$ = "S" THEN
       LOCATE 23, 1: PRINT "Press <ENTER> for Totals "; : fl = 1: GOSUB 62000
       CLS
       ELSE
       PRINT #10, CHR$(12);
      END IF
      FOR years = 0 TO 11
      IF yrdesc(years) <= 1 THEN 10301
       PRINT #10, "PAYMENTS FOR TAX YEAR: "; yrdesc(years); TAB(50); "PERIOD "; BDATE$; " - "; EDATE$
       PRINT #10, ""
      PRINT #10, "Total Payments for Period = "; USING "$$#######,#.##"; totpay#(1, 1, years)
      PRINT #10, "BROKEN DOWN AS FOLLOWS .... "
      PRINT #10, "Principal Tax ............. "; USING "$$#######,#.##"; totprin#(1, 1, years)
      PRINT #10, "Late List Amount .......... "; USING "$$#######,#.##"; totsd#(1, 1, years)
      PRINT #10, "Interest .................. "; USING "$$#######,#.##"; totint#(1, 1, years)
      PRINT #10, "Advertising ............... "; USING "$$#######,#.##"; totadv#(1, 1, years)
      PRINT #10, "Collection ................ "; USING "$$#######,#.##"; totcol#(1, 1, years)
      PRINT #10, ""
      IF devspec$ = "S" THEN LOCATE 23, 1: PRINT "Press ENTER for Next Page "; : fl = 1: GOSUB 62000: CLS
      PRINT #10, "RELEASES FOR TAX YEAR "; yrdesc(years): PRINT #10, ""
      PRINT #10, "Total Releases for Period = "; USING "$$#######,#.##"; totpay#(2, 1, years)
      PRINT #10, "BROKEN DOWN AS FOLLOWS .... "
      PRINT #10, "Principal Tax ............. "; USING "$$#######,#.##"; totprin#(2, 1, years)
      PRINT #10, "Late List Amount .......... "; USING "$$#######,#.##"; totsd#(2, 1, years)
      PRINT #10, "Interest .................. "; USING "$$#######,#.##"; totint#(2, 1, years)
      PRINT #10, "Advertising ............... "; USING "$$#######,#.##"; totadv#(2, 1, years)
      PRINT #10, "Collection ................ "; USING "$$#######,#.##"; totcol#(2, 1, years)
      PRINT #10, ""
      PRINT #10, "OVERPAYMENTS .............. "; USING "$$########,#.##"; TOTOVER#
      PRINT #10, "     Total Net A/R Change = "; USING "$$#######,#.##"; totpay#(1, 1, years) + totpay#(1, 2, years) + totpay#(2, 1, years) + totpay#(2, 2, years) + TOTOVER#
      TOTOVER# = 0: REM CLEAR OVERPAY AFTER FIRST YEAR PRINTS
      IF devspec$ = "S" THEN LOCATE 23, 1: PRINT "Press ENTER for Next Page "; : fl = 1: GOSUB 62000: CLS
      IF devspec$ = "P" THEN LPRINT CHR$(12);
10301 NEXT years
      CLOSE
IF devspec$ = "S" THEN
 LOCATE 23, 1: PRINT "THAT'S ALL : PRESS ANY KEY TO CONTINUE "; : fl = 1
 GOSUB 62000: CLOSE : GOTO 10
END IF
GOTO 10

CRHEADING:
PRINT #10, TAB(25); "Tax Billing System : Cash Receipt Recap"
PRINT #10, TAB(35); "Date: "; DATE$
PRINT #10, ""
PRINT #10, "Rcpt Date"; TAB(12); "Account Name"; TAB(40); "Recpt#"; TAB(47); "Type"; TAB(60); "Total Amount"
PRINT #10, TAB(12); "Tax Amt"; TAB(22); "Late Amt"; TAB(32); "Inter't"; TAB(42); "Advert"; TAB(52); "Collect"; TAB(62); "Over Pay"; TAB(74); "County"
PRINT #10, STRING$(79, "="): lc = 4: RETURN



20000 OPEN "tbacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 31: FIELD 4, 27 AS met$, 4 AS ARREC$
20010 IF LOF(4) / 22 = 0 THEN RETURN
20020 HI = LOF(4) / 22: LO = 1: me = HI: LENG = LEN(account$)
20030 FOR XX = 1 TO 16: GET 4, me
20040 IF account$ = LEFT$(met$, LENG) THEN NG = 1: record = CVS(ARREC$): CLOSE 4: RETURN
20050 IF account$ > LEFT$(met$, LENG) THEN LO = me ELSE HI = me
20060 me = INT((HI - LO) / 2) + LO
20070 NEXT XX: CLOSE 4
20080 CLOSE 4: NG = 0: RETURN
60000 CLOSE 1: OPEN "tbcust.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 512
60010 FIELD 1, 12 AS acct$, 25 AS firstnam$, 35 AS lnam$, 35 AS A1$, 35 AS A2$, 20 AS cy$, 2 AS st$, 10 AS zp$, 12 AS tp$
      FIELD 1, 186 AS dummy$, 39 AS pdesc$, 22 AS pmap$, 35 AS psub$, 4 AS plot$, 35 AS pcounty$, 4 AS psize$, 8 AS pland$, 8 AS mv$
      FIELD 1, 341 AS dummy1$, 8 AS mh$, 8 AS fm$, 8 AS ls$, 4 AS op$, 4 AS me$, 6 AS HOMEEXP$, 1 AS taxtype$, 3 AS taxtype1$
      FIELD 1, 381 AS dummy2$, 4 AS curlateamt$, 4 AS curtaxamt$, 4 AS curintamt$, 4 AS curadvamt$, 4 AS curcolamt$, 4 AS ctaxnotice$
      FIELD 1, 405 AS dummy3$, 4 AS ftr$, 4 AS ltr$, 11 AS ss$, 1 AS late$, 15 AS pin$, 68 AS blank$
60025 CLOSE 3: OPEN "tbcust1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #3 LEN = 256
      FOR tloop = 0 TO 9
      FIELD 3, (24 * tloop) AS dummy4$, 4 AS PTAXAMT$(tloop), 4 AS PintAMT$(tloop), 4 AS PADVAMT$(tloop), 4 AS PCOLAMT$(tloop), 4 AS PLATEAMT$(tloop), 4 AS ptaxnotice$(tloop): NEXT tloop
60030 RETURN
61000 CLOSE 2: OPEN "tbartran.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 96
61010 FIELD 2, 6 AS tdate$, 4 AS tamt$, 2 AS ttaxyr$, 20 AS tdesc$, 2 AS ttype$, 4 AS taprin$, 4 AS taint$, 4 AS taadv$, 4 AS tacol$, 4 AS talate$, 4 AS tmrec$, 4 AS TNEXTTR$, 4 AS tbal$, 4 AS tdisamt$, 26 AS tblank$
61020 RETURN
61500 CLOSE 12: OPEN "tbrptran.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #12 LEN = 64
      FIELD 12, 6 AS rpdate$, 4 AS rpbmamt$, 4 AS rpsdamt$, 4 AS rpint$, 4 AS rpadv$, 4 AS rpcol$, 4 AS rpover$, 1 AS rpcty$, 2 AS rpyear$, 4 AS rprct$, 4 AS rprec$, 1 AS rptype$
      RETURN

62000 COLOR 14: CF = 0: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
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
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN
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
               fl = 1: GOSUB 62000: IF CF = 1 THEN CLOSE : GOTO 10 ELSE GOSUB 62000
            
     END SELECT

33000 CLS
33010 LOCATE 13, 10
      PRINT "Press <ENTER> to continue or press <ESC> to stop printing"
33020 z$ = INKEY$: IF z$ = "" THEN 33020
33025 IF ASC(z$) = 27 THEN ok2continue = 0: GOTO 33050
33030 IF ASC(z$) = 13 THEN ok2continue = 1: GOTO 33050
33040 GOTO 33020
33050 CLS : LOCATE 13, 23: PRINT "Now Processing : Please be Patient!"
33055 LOCATE 23, 1: PRINT "Press <ESC> to stop printing . . . ."
33060 RETURN

