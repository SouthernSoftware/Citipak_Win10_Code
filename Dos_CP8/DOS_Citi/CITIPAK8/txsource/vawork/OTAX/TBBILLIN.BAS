10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbbilling
30 REM Tax Billing Menu
40 REM Last Revision Date:11-29-88
50 REM by Nick Picerno
60 CLS : CLEAR
   WIDTH LPRINT 132
   ON ERROR GOTO networkerror
   DEFDBL T
110 scren$ = "3.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500
130 LOCATE 5, 1: COLOR 11
200 PRINT TAB(30); "Tax Billing System"
210 PRINT TAB(30); " Billing/Interest"
220 PRINT : PRINT
230 PRINT TAB(25); "1. Attach Late List Penalty to Account"
240 PRINT
250 PRINT TAB(25); "2. Calculate and Print Pre-Billing Report"
260 PRINT
270 PRINT TAB(25); "3. Print and Post Tax Billing"
280 PRINT
290 PRINT TAB(25); "4. Print Billing Summary Report"
300 PRINT
    PRINT TAB(25); "5. Apply Interest to Accounts"
310 LOCATE 20, 1: COLOR 11: PRINT TAB(25); "Selection No. (ESC-Master Menu) : "; : fl = -1: GOSUB 62000
320 IF cf = 1 THEN CLOSE : RUN "tbmenu"
325 IF LEN(in$) = 0 THEN 310
330 IF VAL(in$) < 1 OR VAL(in$) > 5 THEN 310
340 ON VAL(in$) GOTO 1000, 2000, 3000, 4000, 5000
1000 CLS : scren$ = "3.01.00": GOSUB 63500: COLOR 11:
     LOCATE 1, 28: PRINT "Attach Late List Penalty":
 LOCATE 3, 1: COLOR 11
PRINT "Enter Customer Account Number (TAB-Last Name : ESC-Menu) .. "; : fl = 8
     GOSUB 62000:
     IF cf = 1 THEN
     CLOSE : GOTO 10
     END IF
     IF cf = 2 THEN
      GOSUB namesearch
      IF ggo = 1 THEN 1005
      IF ggo = 0 THEN LOCATE 10, 1: PRINT "Account Not Found : Press Any Key to Continue "; : fl = 1: GOSUB 62000: GOTO 1000
     END IF

account$ = in$
COLOR 11
PRINT "Enter Parcel Number ... "; : fl = -4: GOSUB 62000
     parcel$ = in$

 GOSUB 1750: IF NG = 0 THEN LOCATE 8, 1: PRINT "Customer Account Not Located, Press any key to try Again!"; : fl = 1: GOSUB 62000: CLOSE : GOTO 1000
 GOSUB 60000: GET 1, RECORD: GET 3, RECORD:  LOCATE 23, 1: PRINT STRING$(79, 32);
1005 GET 1, RECORD: GET 3, RECORD
FOR zz = 5 TO 15: LOCATE zz, 1: PRINT STRING$(79, 32): NEXT zz
 COLOR 11: LOCATE 8, 1: PRINT "Account of ...";
 K = INSTR(firstnam$, "  "): IF K = 0 THEN n$ = lnam$ ELSE n$ = LEFT$(firstnam$, K) + lnam$
 PRINT TAB(20); n$
 COLOR 12: PRINT : PRINT "STATUS: "; : IF late$ = "Y" THEN COLOR 12: PRINT "LATE LIST PENALTY ATTACHED" ELSE COLOR 10: PRINT "No Late List Penalty Attached"
1010 LOCATE 14, 1: COLOR 11: PRINT "Attach Late Listing to This Account (Y/N) .. "; : fl = 1: GOSUB 62210
 IF NOT (in$ = "Y" OR in$ = "N") THEN 1010
 IF in$ = "Y" THEN 1020 ELSE LSET late$ = "N": PUT 1, RECORD: CLOSE 4
 GOTO 1000
1020 LSET late$ = "Y": PUT 1, RECORD: CLOSE 4: GOTO 1000
1750 CLOSE 4
     OPEN "tbacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 26
     FIELD 4, 22 AS met$, 4 AS ARREC$
     IF LEN(account$) < 9 THEN account$ = account$ + STRING$(9 - LEN(account$), 32)
     IF LEN(parcel$) < 13 THEN parcel$ = parcel$ + STRING$(13 - LEN(parcel), 32)
     thisone$ = account$ + parcel$
     IF LEN(thisone$) < 22 THEN thisone$ = thisone$ + STRING$(22 - LEN(thisone$), 32)
1752 IF LOF(4) / 26 = 0 THEN RETURN
1755 HI = LOF(4) / 26: LO = 1: ME = HI
1760 FOR XX = 1 TO 16: GET 4, ME
1765 IF thisone$ = met$ THEN NG = 1: RECORD = CVS(ARREC$): CLOSE 4: RETURN
1770 IF thisone$ > met$ THEN LO = ME ELSE HI = ME
1775 ME = INT((HI - LO) / 2) + LO
1780 NEXT XX: CLOSE 4
1782 CLOSE 4: NG = 0: RETURN
1785 REM
namesearch:

LOCATE 3, 1:
COLOR 11: PRINT "Enter Last Name (Partial Accepted : ESC - End) .. "; : fl = 25
 GOSUB 62000: IF cf = 1 THEN ggo = 0: RETURN
 nn$ = in$
 nn = LEN(nn$)
 CLOSE 4: OPEN "R", 4, "tblname.dat", 24: FIELD 4, 20 AS lastnam$, 4 AS MREC$
 CLOSE 1: GOSUB 60000

FOR LL = 1 TO LOF(4) / 24: GET 4, LL
IF LEFT$(lastnam$, nn) = nn$ THEN GOTO N1
n3:
NEXT LL
ggo = 0: RETURN
N1:
GET 1, CVS(MREC$)
LOCATE 6, 1: COLOR 11
kk = INSTR(firstnam$, "  ")
IF kk = 0 THEN kk = 25
IF kk = 1 THEN nme$ = lnam$ ELSE nme$ = LEFT$(firstnam$, kk) + lnam$
COLOR 10
FOR lll = 6 TO 9: LOCATE lll, 1: PRINT STRING$(79, 32): NEXT lll
LOCATE 6, 1
PRINT "Full Account Name: "; nme$
PRINT "   Account Number: "; ACCT$
PRINT "    Parcel Number: "; PIN$
ACCTNUMBER$ = ACCT$
PRINT
n2:
LOCATE 9, 1: COLOR 11
PRINT "Correct Account (Y/N/Q) .. "; : fl = 1: GOSUB 62210
IF NOT (in$ = "Y" OR in$ = "N" OR in$ = "Q") THEN GOTO n2
IF in$ = "Q" THEN ggo = 0: RETURN
IF in$ = "Y" THEN ggo = 1: RECORD = CVS(MREC$): RETURN
IF in$ = "N" THEN GOTO n3
GOTO n2

2000 CLS : REALFLAG = 0: PERSFLAG = 0
     scren$ = "3.02.00": GOSUB 63500: COLOR 11: LOCATE 1, 30: PRINT "Print Pre-Billing Report": totreal# = 0: totpers# = 0: tottax# = 0
     LOCATE 3, 1: PRINT "Enter this Year's Real Estate Tax Rate (PRESS <ENTER> FOR PERSONAL) .. "
     COLOR 11: LOCATE 4, 1: PRINT "(Ex: .50 for 50 cents/$100 Assessed Val.) <ESC-Menu> ..... "; : fl = -5: GOSUB 62000
     IF cf = 1 THEN 10 ELSE TAXRATE = VAL(in$)
     IF TAXRATE > 0 THEN REALFLAG = 1: GOTO 2050
     COLOR 11
     LOCATE 3, 1: PRINT STRING$(79, 32)
     LOCATE 4, 1: PRINT STRING$(79, 32)
     LOCATE 3, 1: PRINT "Enter this Year's Personal Tax Rate .. ";
     COLOR 11: LOCATE 4, 1: PRINT "(Ex: .50 for 50 cents/$100 Assessed Val.) <ESC-Menu> ..... "; : fl = -5: GOSUB 62000
     IF cf = 1 THEN 10 ELSE TAXRATE = VAL(in$)
     PERSFLAG = 1
2050 COLOR 11: LOCATE 7, 1: PRINT "Enter Late List Percentage (if any) (<ESC>-Menu) ......... "; : fl = -4: GOSUB 62000
     IF cf = 1 THEN 10 ELSE late = VAL(in$): late = late / 100
2100 GOSUB 60000: GOSUB 2800: GOSUB 2900
     FOR x = 1 TO LOF(4) / 24
     GET 4, x: RECORD = CVS(MREC$)
     IF RECORD < 1 OR RECORD > LOF(1) / 512 THEN 2200
     GET 1, RECORD: GET 3, RECORD
     REALVALUE# = CVD(PLAND$) + CVD(PBLDG$)
     PERSVALUE# = CVD(pbus$) + CVD(pres$) + CVD(ppub$)
     REALVALUE# = INT((REALVALUE# * 100) + .5) / 100
     PERSVALUE# = INT((PERSVALUE# * 100) + .5) / 100
2110 IF REALFLAG = 1 THEN PERSVALUE# = 0
     IF PERSFLAG = 1 THEN REALVALUE# = 0
     IF REALVALUE# + PERSVALUE# = 0 THEN 2200
     IF lc >= 40 THEN LPRINT CHR$(12); : GOSUB 2900
     K = INSTR(firstnam$, "  "): IF K = 0 THEN K = 20
     ACCTNAME$ = LEFT$(firstnam$, K) + lnam$
2120 LPRINT LEFT$(ACCT$, 8); TAB(10); PIN$;
     LPRINT TAB(35); USING "#,###,###"; REALVALUE#;
     LPRINT TAB(47); USING "#,###,###"; PERSVALUE#;
     LPRINT TAB(59); USING "#,###,###"; (REALVALUE# + PERSVALUE#);
     RegTax# = ((REALVALUE# - CVS(SCEXP$)) + PERSVALUE#) / 100 * TAXRATE
     IF late$ = "Y" THEN
       LATETAX! = ((PERSVALUE# / 100) * TAXRATE) * late
        ELSE LATETAX! = 0
     END IF

     RegTax# = RegTax# * 100: RegTax# = INT(RegTax# + .5) / 100
      IF PERSFLAG = 1 THEN
       Prorate = CVS(useexp$)
       IF Prorate <= 0 OR Prorate > 12 THEN Prorate = 12
       RegTax# = RegTax# * (Prorate / 12)
      END IF
     LATETAX! = LATETAX! * 100: LATETAX! = INT(LATETAX! + .5) / 100
     TAXOWED! = RegTax# + LATETAX!
     TAXOWED! = TAXOWED! * 100: TAXOWED! = INT(TAXOWED! + .5) / 100
     IF TAXOWED! < 0 THEN TAXOWED! = 0
     LPRINT TAB(70); USING "$###,###.##"; TAXOWED!
     LPRINT LEFT$(PDESC$, 22);
     LPRINT TAB(30); USING "###,#.##"; LATETAX!;
     LPRINT TAB(40); USING "###.##"; CVS(SCEXP$);
     LPRINT TAB(65); USING "$####,#.##"; origtax#
     tottax# = tottax# + TAXOWED!
     totreal# = totreal# + REALVALUE#
     totpers# = totpers# + PERSVALUE#
     TOTREG# = TOTREG# + RegTax#
     totlate# = totlate# + LATETAX!
     TOTSRCIT# = TOTSRCIT# + CVS(SCEXP$)
     LPRINT firstnam$ + lnam$
     LPRINT STRING$(79, "-")
     lc = lc + 4

2200 NEXT x
2201 CLOSE :
     LPRINT "TOTALS : Total Taxes to Bill ... "; USING "$########,#.##"; tottax#
     LPRINT "       : Total Real Valuation .. "; USING "$########,#"; totreal#
     LPRINT "       : Total Pers. Valuation . "; USING "$########,#"; totpers#
     LPRINT "       : Total Sr Cit. Dis. .... "; USING "$########,#.##"; TOTSRCIT#
     LPRINT
     LPRINT "Breakdown:"
     LPRINT "       : Total Regular Tax ..... "; USING "$########,#.##"; TOTREG#
     LPRINT "       : Total Late List Pen. .. "; USING "$########,#.##"; totlate#
     LPRINT CHR$(12);
     GOTO 10
2800 OPEN "tblname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
     FIELD 4, 20 AS lastnam$, 4 AS MREC$: ENDOFFILE = LOF(4) / 24: RETURN
2850 OPEN "tbacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #3 LEN = 16: FIELD 3, 12 AS met$, 4 AS MREC$: ENDOFFILE = LOF(3) / 16: RETURN
2900 LPRINT "Property Tax Pre-Billing Report         :        For Audit Purposes Only  "
2910 LPRINT "Report Date: "; DATE$; TAB(50); "Gen. Tax Rate: "; USING "$##.##"; TAXRATE
     IF REALFLAG = 1 THEN LPRINT "REAL ESTATE TAXES" ELSE LPRINT "PERSONAL PROPERTY TAXES"
2920 LPRINT
2930 LPRINT "Account Number/Parcel #"; TAB(35); "Real Val."; TAB(47); "Pers. Val."; TAB(59); "Tot. Val."; TAB(71); "Calc. Bill"
     LPRINT "Property Desc "; TAB(30); "Late List"; TAB(41); "Disc't "; TAB(52); "Dis. Amount"; TAB(65); "Reg Tax Amt"
     LPRINT "Account Name"
     LPRINT STRING$(79, "="): lc = 4: RETURN

3000 CLS : scren$ = "3.03.00": GOSUB 63500: COLOR 11: LOCATE 1, 32: PRINT "Print Tax Bills"
     COLOR 31
     LOCATE 3, 1: PRINT "*** "; : COLOR 15: PRINT "YOU MUST PRINT THE REAL ESTATE BILLS FIRST EACH TAX SEASON "; : COLOR 31: PRINT "***"
     COLOR 15
     LOCATE 4, 1: PRINT "PRESS <ENTER> TO CONTINUE "; : fl = 1
     GOSUB 62000
     CLS : scren$ = "3.03.00": GOSUB 63500: COLOR 11: LOCATE 1, 32: PRINT "Print Tax Bills"
     CLOSE : firsttime$ = "Y": REALFLAG = 0: PERSFLAG = 0
     totreal# = 0
     totpers# = 0
     tottax# = 0
     totlate# = 0
     TOTSRCIT# = 0
     LOCATE 3, 1: PRINT "Enter this Year's Real Estate Tax Rate (PRESS <ENTER> FOR PERSONAL) .. "
     COLOR 11: LOCATE 4, 1: PRINT "(Ex: .50 for 50 cents/$100 Assessed Val.) <ESC-Menu> ..... "; : fl = -5: GOSUB 62000
     IF cf = 1 THEN 10 ELSE TAXRATE = VAL(in$)
     IF TAXRATE > 0 THEN REALFLAG = 1: GOTO 3050
     COLOR 11
     LOCATE 3, 1: PRINT STRING$(79, 32)
     LOCATE 4, 1: PRINT STRING$(79, 32)
     LOCATE 3, 1: PRINT "Enter this Year's Personal Tax Rate .. ";
     COLOR 11: LOCATE 4, 1: PRINT "(Ex: .50 for 50 cents/$100 Assessed Val.) <ESC-Menu> ..... "; : fl = -5: GOSUB 62000
     IF cf = 1 THEN 10 ELSE TAXRATE = VAL(in$)
     PERSFLAG = 1

       
3050 COLOR 11: LOCATE 7, 1: PRINT "Enter the Tax Year (ex. 1989) ........... "; : fl = -4: GOSUB 62000: IF cf = 1 THEN 10
     IF LEN(in$) <> 4 THEN 3050 ELSE year$ = in$
3017 COLOR 11: LOCATE 9, 1: PRINT "Enter the Due Date (mmddyy) ............. "; : fl = 6: GOSUB 62000
     IF cf = 1 THEN 10
     IF LEN(in$) <> 6 THEN 3017
     DUEDATE$ = LEFT$(in$, 2) + "/" + MID$(in$, 3, 2) + "/" + RIGHT$(in$, 2)

     
3010 COLOR 11: LOCATE 10, 1: PRINT "Print (A)ll or (I)ndivid. Tax Bill (ESC-Menu)... "; : fl = 1: GOSUB 62210: LOCATE 10, 50: PRINT in$: IF cf = 1 THEN 10
     IF NOT (in$ = "A" OR in$ = "I") THEN 3010 ELSE num$ = in$
3015 COLOR 11: LOCATE 11, 1: PRINT "Print (R)eprint or (F)ull Processing (ESC-Menu). "; : fl = 1: GOSUB 62210: LOCATE 11, 50: PRINT in$: IF cf = 1 THEN 10
     IF NOT (in$ = "R" OR in$ = "F") THEN 3015 ELSE billinfo$ = in$

3020 IF num$ = "I" THEN 3100
3030 REM all bills using last name to print
3035 GOSUB 2800
     GOSUB 60000
     GOSUB 61000
     IF LOF(4) / 24 = 0 THEN CLOSE : GOTO 10
     GOSUB taxalign
     IF billinfo$ = "F" AND num$ = "A" THEN
          GOSUB 10000: REM open tbsummary file & initialize
          OPEN "R", 14, "TBYEAR.DAT", 4: FIELD 14, 4 AS YR1$
          GET 14, 1
          LSET YR1$ = year$: PUT 14, 1: CLOSE 14

          ELSE
          IF billinfo$ = "F" AND num$ = "I" THEN GOSUB 10000
     END IF
     FOR Xy = 1 TO LOF(4) / 24
          GET 4, Xy: RECORD = CVS(MREC$)
          IF RECORD < 1 OR RECORD > LOF(1) / 512 THEN 3040
          GET 1, RECORD: GET 3, RECORD
          IF PERSFLAG = 1 THEN 3036
          IF billinfo$ = "F" THEN
               REM IND YEARS
              LSET ptaxamt$(9) = MKS$(CVS(ptaxamt$(9)) + CVS(ptaxamt$(8)))
              LSET pintamt$(9) = MKS$(CVS(pintamt$(9)) + CVS(pintamt$(8)))
              LSET PADVAMT$(9) = MKS$(CVS(PADVAMT$(9)) + CVS(PADVAMT$(8)))
              LSET PCOLAMT$(9) = MKS$(CVS(PCOLAMT$(9)) + CVS(PCOLAMT$(8)))
              LSET plateamt$(9) = MKS$(CVS(plateamt$(9)) + CVS(plateamt$(8)))
              LSET ptaxnotice$(9) = MKS$(CVS(ptaxnotice$(8)))
              
               FOR zz = 8 TO 1 STEP -1
               LSET ptaxamt$(zz) = MKS$(CVS(ptaxamt$(zz - 1)))
               LSET pintamt$(zz) = MKS$(CVS(pintamt$(zz - 1)))
               LSET PADVAMT$(zz) = MKS$(CVS(PADVAMT$(zz - 1)))
               LSET PCOLAMT$(zz) = MKS$(CVS(PCOLAMT$(zz - 1)))
               LSET plateamt$(zz) = MKS$(CVS(plateamt$(zz - 1)))
               LSET ptaxnotice$(zz) = MKS$(CVS(ptaxnotice$(zz - 1)))
               NEXT zz
               LSET ptaxamt$(0) = MKS$(CVS(curtaxamt$))
               LSET pintamt$(0) = MKS$(CVS(curintamt$))
               LSET PADVAMT$(0) = MKS$(CVS(curadvamt$))
               LSET PCOLAMT$(0) = MKS$(CVS(curcolamt$))
               LSET plateamt$(0) = MKS$(CVS(curlateamt$))
               LSET ptaxnotice$(0) = MKS$(CVS(ctaxnotice$))
                              
               LSET curlateamt$ = MKS$(0): LSET curtaxamt$ = MKS$(0)
               LSET curintamt$ = MKS$(0): LSET curcolamt$ = MKS$(0)
               LSET curadvamt$ = MKS$(0): LSET ctaxnotice$ = MKS$(0)

               REM  move curtax to past due tax
               PUT 1, RECORD: PUT 3, RECORD

          END IF
3036 IF billinfo$ = "F" THEN
           year = VAL(RIGHT$(year$, 2)): COUNT = year * 10000
           COUNT = COUNT + LOF(9) / 256 + 1
           ELSE
           COUNT = 0
          END IF
          GOSUB calcbill
          IF TAXOWED! = 0 THEN 3040
         
          IF REALFLAG = 1 THEN GOSUB PRINTREALBILL
          IF PERSFLAG = 1 THEN GOSUB PRINTPERSBILL

          IF billinfo$ = "F" THEN GOSUB postbill
3040 NEXT Xy
     LPRINT CHR$(12);
     CLOSE
     GOTO 10

3100 REM individual bill only
     FOR JJ = 3 TO 22: LOCATE JJ, 1: PRINT STRING$(79, 32): NEXT JJ
     COLOR 11: LOCATE 3, 1
     PRINT "Enter Customer Account Number (TAB-Last Name : ESC-Menu) .. "; : fl = 8
     GOSUB 62000:
     IF cf = 1 THEN
     CLOSE : GOTO 10
     END IF
     IF cf = 2 THEN
      GOSUB namesearch
      IF ggo = 1 THEN 3130
      IF ggo = 0 THEN LOCATE 10, 1: PRINT "Account Not Found : Press Any Key to Continue "; : fl = 1: GOSUB 62000: GOTO 3100
     END IF

account$ = in$
COLOR 11
PRINT "Enter Parcel Number ... "; : fl = -4: GOSUB 62000
     parcel$ = in$
     GOSUB 1750: IF NG = 0 THEN LOCATE 8, 1: PRINT "Customer Account Not Located, Press any key to try Again!"; : fl = 1: GOSUB 62000: CLOSE : GOTO 3100
      GOSUB 60000
3130 GET 1, RECORD: GET 3, RECORD: LOCATE 23, 1: PRINT STRING$(79, 32);
3135 GOSUB 61000
     IF firsttime$ = "Y" THEN GOSUB taxalign: firsttime$ = "N"
     IF billinfo$ = "F" THEN
          GOSUB 10000: REM open tbsummary file
          REM IND YEARS
              LSET ptaxamt$(9) = MKS$(CVS(ptaxamt$(9)) + CVS(ptaxamt$(8)))
              LSET pintamt$(9) = MKS$(CVS(pintamt$(9)) + CVS(pintamt$(8)))
              LSET PADVAMT$(9) = MKS$(CVS(PADVAMT$(9)) + CVS(PADVAMT$(8)))
              LSET PCOLAMT$(9) = MKS$(CVS(PCOLAMT$(9)) + CVS(PCOLAMT$(8)))
              LSET plateamt$(9) = MKS$(CVS(plateamt$(9)) + CVS(plateamt$(8)))
              LSET ptaxnotice$(9) = MKS$(CVS(ptaxnotice$(8)))
             
               FOR zz = 8 TO 1 STEP -1
               LSET ptaxamt$(zz) = MKS$(CVS(ptaxamt$(zz - 1)))
               LSET pintamt$(zz) = MKS$(CVS(pintamt$(zz - 1)))
               LSET PADVAMT$(zz) = MKS$(CVS(PADVAMT$(zz - 1)))
               LSET PCOLAMT$(zz) = MKS$(CVS(PCOLAMT$(zz - 1)))
               LSET plateamt$(zz) = MKS$(CVS(plateamt$(zz - 1)))
               LSET ptaxnotice$(zz) = MKS$(CVS(ptaxnotice$(zz - 1)))
               NEXT zz
               LSET ptaxamt$(0) = MKS$(CVS(curtaxamt$))
               LSET pintamt$(0) = MKS$(CVS(curintamt$))
               LSET PADVAMT$(0) = MKS$(CVS(curadvamt$))
               LSET PCOLAMT$(0) = MKS$(CVS(curcolamt$))
               LSET plateamt$(0) = MKS$(CVS(curlateamt$))
               LSET ptaxnotice$(0) = MKS$(CVS(ctaxnotice$))
                             
               LSET curlateamt$ = MKS$(0): LSET curtaxamt$ = MKS$(0)
               LSET curintamt$ = MKS$(0): LSET curcolamt$ = MKS$(0)
               LSET curadvamt$ = MKS$(0): LSET ctaxnotice$ = MKS$(0)

               PUT 1, RECORD: PUT 3, RECORD
     END IF
          IF billinfo$ = "F" THEN
           year = VAL(RIGHT$(year$, 2)): COUNT = year * 10000
           COUNT = COUNT + LOF(9) / 256 + 1
           ELSE
           COUNT = 0
          END IF

     GOSUB calcbill
     IF TAXOWED! = 0 THEN 3140
    
     IF REALFLAG = 1 THEN GOSUB PRINTREALBILL
     IF PERSFLAG = 1 THEN GOSUB PRINTPERSBILL

     IF billinfo$ = "F" THEN GOSUB postbill
3140 COLOR 11: LOCATE 20, 1: PRINT STRING$(79, 32): LOCATE 20, 1
     LOCATE 20, 1: PRINT "Any More Individual Bills to Print (Y/N) .. "; : fl = 1
     GOSUB 62210: IF NOT (in$ = "Y" OR in$ = "N") THEN 3140
     IF in$ = "Y" THEN 3100
     LPRINT CHR$(12);
     CLOSE
     GOTO 10

taxalign:
     LOCATE 20, 1: COLOR 11
     PRINT "Print Alignment (Y/N) or (ESC-ABORT PROCESS) .. "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN CLOSE : GOTO 10
     IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO taxalign
     IF in$ = "N" THEN RETURN
     REM FORM IS BLANK 21 PRINT LINES
     REM NAME AND ADDRESS SHOULD BE LINES
     LPRINT
     LPRINT TAB(4); "TOWN OF REMINGTON   P.O. BOX 276   REMINGTON, VA 22734"; TAB(60); "XXXXXX"
     LPRINT
     LPRINT
     LPRINT
     LPRINT TAB(4); STRING$(30, "X"); TAB(40); "XX.XX"; TAB(47); "XX,XXX,XXX"; TAB(59); "$X,XXX,XXX.XX"
     LPRINT TAB(4); STRING$(30, "X"); TAB(40); "XX.XX"; TAB(47); "XX,XXX,XXX"; TAB(59); "$X,XXX,XXX.XX"
     LPRINT
     LPRINT
     LPRINT
     LPRINT
     LPRINT TAB(59); "$X,XXX,XXX.XX"
     LPRINT
     LPRINT TAB(7); STRING$(30, "X")
     LPRINT TAB(7); STRING$(30, "X")
     LPRINT TAB(7); STRING$(30, "X")
     LPRINT TAB(7); STRING$(30, "X")
     LPRINT
     LPRINT
     LPRINT
     LPRINT
     GOTO taxalign
    
calcbill:
     REALVALUE# = CVD(PLAND$) + CVD(PBLDG$)
     PERSVALUE# = CVD(pbus$) + CVD(pres$) + CVD(ppub$)
    
     REALVALUE# = INT((REALVALUE# * 100) + .5) / 100
     PERSVALUE# = INT((PERSVALUE# * 100) + .5) / 100
     IF REALFLAG = 1 THEN PERSVALUE# = 0
     IF PERSFLAG = 1 THEN REALVALUE# = 0
     RegTax# = ((REALVALUE# - CVS(SCEXP$)) + PERSVALUE#) / 100 * TAXRATE
     IF RegTax# < 0 THEN RegTax# = 0

     RegTax# = RegTax# * 100: RegTax# = INT(RegTax# + .5) / 100

      IF PERSFLAG = 1 THEN
       Prorate = CVS(useexp$)
       IF Prorate <= 0 OR Prorate > 12 THEN Prorate = 12
       RegTax# = RegTax# * (Prorate / 12)
      END IF



     TAXOWED! = 0: REM INITIALIZE TAXOWED=0
     TAXOWED! = RegTax#
     TAXOWED! = TAXOWED! * 100: TAXOWED! = INT(TAXOWED! + .5) / 100
     RETURN



PRINTREALBILL:
        REM GET NAME INTO N$
        K = INSTR(firstnam$, "  ")
        IF K = 0 THEN K = 25
        IF K = 1 THEN n$ = lnam$ ELSE n$ = LEFT$(firstnam$, K) + lnam$
        LOT = INSTR(prop$, "LOT")
        IF LOT = 0 THEN GOTO CNT
        LOT$ = MID$(prop$, LOT, 9)
CNT: LPRINT TAB(12); "DUE ON OR BEFORE "; DUEDATE$
     LPRINT TAB(4); "TOWN OF REMINGTON   P.O. BOX 276   REMINGTON, VA 22734"; TAB(60); USING "######"; COUNT
     LPRINT
     LPRINT
     REM LINE 5
     LPRINT TAB(5); LEFT$(PDESC$, 30); TAB(41); USING "#.##"; TAXRATE
     LPRINT TAB(5); LEFT$(PLOT$, 30);
      IF CVD(PLAND$) - CVS(SCEXP$) > 0 THEN
       LPRINT TAB(41); "LAND"; TAB(47); USING "$#####,#.##"; CVD(PLAND$) - CVS(SCEXP$);
       LPRINT TAB(59); USING "$#####,#.##"; ((CVD(PLAND$) - CVS(SCEXP$)) / 100) * TAXRATE
       ELSE
       LPRINT
      END IF
     LPRINT TAB(5); LEFT$(PMAP$, 30);
      IF CVD(PBLDG$) > 0 THEN
      LPRINT TAB(41); "IMP"; TAB(47); USING "$#####,#.##"; CVD(PBLDG$);
      LPRINT TAB(59); USING "$#####,#.##"; (CVD(PBLDG$) / 100) * TAXRATE
      ELSE
      LPRINT
     END IF
     LPRINT TAB(5); ACCT$; " "; PIN$
     LPRINT
     LPRINT TAB(5); "THERE IS A 10% PENALTY IF PAID AFTER "; DUEDATE$
     REM LINE 11
     LPRINT
     LPRINT TAB(47); DUEDATE$; TAB(59); USING "$#####,#.##"; TAXOWED!
     LPRINT
     REM LINE 14

     LPRINT TAB(7); n$
     LPRINT TAB(7); A1$
     LPRINT TAB(7); A2$
     LPRINT TAB(7); cy$; " "; st$; " "; zp$
     LPRINT : LPRINT : LPRINT : LPRINT
     RETURN
          
                                
postbill:
  REM post to account balance first

  LSET curlateamt$ = MKS$(0)
  LSET curtaxamt$ = MKS$(TAXOWED!)
  LSET ctaxnotice$ = MKS$(COUNT)
  
  PUT 1, RECORD: PUT 3, RECORD
  GOSUB figurebalance
  acctbal# = curbal# + pastdue#

 REM post to tbsummary account
  LSET wacct$ = ACCT$: LSET wname$ = n$
  LSET wreal$ = MKD$(REALVALUE#)
  LSET wpers$ = MKD$(PERSVALUE#)
  LSET wsrcit$ = MKD$(exemption#)
  LSET wtaxgen$ = MKD$(RegTax#)
  LSET wtaxfire$ = MKD$(0)
  IF PERSFLAG = 1 THEN
   LSET wcartag$ = MKD$(1): REM SET TO 1 FOR PERSONAL TAX
   ELSE
   LSET wcartag$ = MKD$(2): REM SET TO 2 FOR REAL ESTATE TAX
  END IF
  LSET wbillnum$ = MKS$(COUNT)
  LSET wtaxowed$ = MKD$(TAXOWED!)
  LSET wlate$ = MKD$(LATETAX!)
  LSET wtaxyear$ = MKS$(TAXYEAR)
  PUT 9, LOF(9) / 256 + 1

 REM post transaction
  LSET TDATE$ = (LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 2))
  LSET TAMT$ = MKS$(TAXOWED!)
  LSET ttype$ = MKI$(1): LSET TDESC$ = "Tax Bill"
  LSET ttaxyr$ = MKI$(VAL(year$))
  LSET taprin$ = MKS$(TAXOWED! - LATETAX!)
  LSET taint$ = MKS$(0): LSET taadv$ = MKS$(0)
  LSET tacol$ = MKS$(0): LSET talate$ = MKS$(LATETAX!)
  LSET TMREC$ = MKS$(RECORD)
  LSET TNEXTTR$ = MKS$(0): LSET tbal$ = MKS$(acctbal#)
  LSET tdisamt$ = MKS$(0): LSET tblank$ = ""
  PUT 2, LOF(2) / 96 + 1
  Trecord = LOF(2) / 96

  IF CVS(ftr$) = 0 THEN LSET ftr$ = MKS$(Trecord): LSET ltr$ = MKS$(Trecord): PUT 1, RECORD: RETURN
  prev = CVS(ltr$): LSET ltr$ = MKS$(Trecord): PUT 1, RECORD
  GET 2, prev: LSET TNEXTTR$ = MKS$(Trecord): PUT 2, prev
  RETURN

figurebalance:
     curbal# = 0: pastdue# = 0
     curbal# = (CVS(curtaxamt$) + CVS(curintamt$) + CVS(curadvamt$) + CVS(curcolamt$) + CVS(curlateamt$))
     curbal# = INT((curbal# * 100) + .5) / 100
     FOR LL = 0 TO 9
     pastdue# = pastdue# + (CVS(ptaxamt$(LL)) + CVS(pintamt$(LL)) + CVS(PADVAMT$(LL)) + CVS(PCOLAMT$(LL)) + CVS(plateamt$(LL)))
     pastdue# = INT((pastdue# * 100) + .5) / 100
     NEXT LL
     RETURN


4000 CLS : scren$ = "3.04.00": GOSUB 63500: COLOR 11: LOCATE 1, 30: PRINT "Print Billing Summary"
4010 LOCATE 3, 1: PRINT "ENTER (P)ersonal or (R)eal Estate Tax Summary Report (P,R or ESC-Menu) .. "; : fl = 1: GOSUB 62210
4020 IF cf = 1 THEN 10
     IF NOT (in$ = "P" OR in$ = "R") THEN 4010
     typ$ = in$
4030 GOSUB 10000
     IF LOF(9) / 256 = 0 THEN CLOSE : GOTO 10
     GOSUB 4900
4040 FOR x = 1 TO LOF(9) / 256: GET 9, x
     IF CVS(wtaxyear$) = 1994 THEN 4100 ELSE 4041
4041 IF CVD(wcartag$) = 1 AND typ$ = "P" THEN 4042
     IF CVD(wcartag$) = 2 AND typ$ = "R" THEN 4042
     
     GOTO 4100
4042 IF lc >= 56 THEN LPRINT CHR$(12); : GOSUB 4900
     LPRINT wacct$; " "; wname$; TAB(40); USING "######,#"; CVD(wreal$);
     LPRINT TAB(52); USING "######,#"; CVD(wpers$);
     LPRINT TAB(65); USING "######"; CVS(wbillnum$);
     LPRINT TAB(72); USING "####,#.##"; CVD(wtaxowed$)
     LPRINT USING "$######"; CVD(wsrcit$);
     LPRINT TAB(30); USING "$####,#.##"; CVD(wtaxgen$);
     LPRINT TAB(45); USING "$####,#.##"; CVD(wlate$);
     
     tpers = tpers + CVD(wpers$): treal = treal + CVD(wreal$)
     tdisc = tdisc + CVD(wsrcit$): tgtax = tgtax + CVD(wtaxgen$)
     tftax = tftax + CVD(wtaxfire$): ttax = ttax + CVD(wtaxowed$)
     tlate = tlate + CVD(wlate$)
     tdogs = tdogs + CVS(wtaxyear$)
     tpers# = INT((tpers# * 100) + .5) / 100: treal# = INT((treal# * 100) + .5) / 100
     ttax = (ttax * 100): ttax = INT(ttax + .5) / 100
     tgtax = (tgtax * 100): tgtax = INT(tgtax + .5) / 100
     tlate = (tlate * 100): tlate = INT(tlate + .5) / 100
     LPRINT STRING$(79, "-"): lc = lc + 3
4100 NEXT x
     CLOSE : LPRINT CHR$(12)
     LPRINT "Tax Billing Summary Page "; TAB(60); "Date: "; DATE$
     LPRINT
     LPRINT "Total Real Valuation Billed For ............ "; USING "$$#######,.##"; treal
     LPRINT "Total Personal Valuation Billed For ........ "; USING "$$#######,.##"; tpers
     LPRINT "Total Discount Allowed ..................... "; USING "$$#######,.##"; tdisc
     LPRINT "Net Valuation Taxed ........................ "; USING "$$#######,.##"; treal + tpers - tdisc
     LPRINT
     LPRINT "Total Tax Billed For ....................... "; USING "$$#######,#.##"; ttax
     LPRINT
     LPRINT
     LPRINT "Breakdown:"
     LPRINT "Total General Tax Billed For ............... "; USING "$$#######,#.##"; tgtax
     LPRINT "Total Late Listing Penalties Billed For .... "; USING "$$#######,#.##"; tlate
     LPRINT CHR$(12); : GOTO 10

4900 LPRINT "Tax Billing Summary Report"; TAB(60); "Date: "; DATE$
     IF typ$ = "P" THEN LPRINT "Personal Property Taxes" ELSE LPRINT "Real Estate Taxes"
     LPRINT
     LPRINT "Customer Account Number/Name"; TAB(40); "Real Val"; TAB(52); " Pers Val."; TAB(64); "Rcpt#"; TAB(70); "Billed Amt"
     LPRINT "Discount Amt"; TAB(30); "Gen. Tax Amt"; TAB(45); "Late List Pen."
     LPRINT STRING$(79, "="): lc = 5: RETURN

5000 REM
5001 FEEFLAG = 0
5010 scren$ = "3.05.00": CLS : GOSUB 63500: COLOR 11
     LOCATE 1, 32: PRINT "Apply Penalty/Interest"
5020 COLOR 11: LOCATE 3, 1
     PRINT "*** For Penalty"
5030 PRINT "Enter Percentage to Charge on Current Tax (ESC-Menu) .. "; : fl = -4: GOSUB 62000
5040 IF cf = 1 THEN 10 ELSE pcurpercent = VAL(in$): LOCATE 4, 57: PRINT USING "#.##%"; pcurpercent
5045 COLOR 11: LOCATE 5, 1
     PRINT "Enter Percentage to Charge on Past Due Taxes (ESC-Menu) "; : fl = -4: GOSUB 62000
     IF cf = 1 THEN 10 ELSE ppastpercent = VAL(in$): LOCATE 5, 57: PRINT USING "#.##%"; ppastpercent
5046 COLOR 11: LOCATE 8, 1:
     PRINT "*** For Interest"
     PRINT "Enter Percentage to Charge on Current Tax (ESC-Menu) .. "; : fl = -4: GOSUB 62000
     IF cf = 1 THEN 10 ELSE curpercent = VAL(in$): LOCATE 9, 57: PRINT USING "#.##%"; curpercent
     COLOR 11: LOCATE 10, 1
     PRINT "Enter Percentage to Charge on Past Due Taxes (ESC-Menu) "; : fl = -4: GOSUB 62000
     IF cf = 1 THEN 10 ELSE pastpercent = VAL(in$): LOCATE 10, 57: PRINT USING "#.##%"; pastpercent
5047 COLOR 11: LOCATE 13, 1
     PRINT "Are the Above Percentages Correct (Y/N) .. "; : fl = 1
     GOSUB 62210: IF cf = 1 THEN 10
     IF NOT (in$ = "Y" OR in$ = "N") THEN 5047
     IF in$ = "N" THEN 5000
5060 OPEN "R", 2, "tbfee.dat", 128: CLOSE 2: KILL "tbfee.dat"
     GOSUB 60000
     OPEN "tbacct.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 26
     FIELD 4, 22 AS met$, 4 AS MREC$

5070 FOR x = 1 TO LOF(4) / 26: GET 4, x
5071 IF CVS(MREC$) = 0 THEN 5090 ELSE GET 1, CVS(MREC$): GET 3, CVS(MREC$)
5080 IF curpercent > 0 OR pcurpercent > 0 THEN GOSUB curtax
     IF pastpercent > 0 OR ppastpercen > 0 THEN GOSUB pasttax
     IF ctax# > 0 OR ptax# > 0 OR pctax# > 0 OR pptax# > 0 THEN GOSUB 5100
5090 NEXT x:
     GOSUB 5700
     CLOSE
     GOTO 10
5100 REM display account and query for charge
5102 m$ = ACCT$
5110 b = INSTR(firstnam$, "  "): IF b = 0 THEN b = 21
5120 IF b = 1 THEN nme$ = lnam$: GOTO 5140
5130 nme$ = LEFT$(firstnam$, b) + lnam$
5140 COLOR 11: LOCATE 15, 1: PRINT STRING$(53, 32)
      LOCATE 15, 1:
      PRINT "Processing the Account of ... "; : COLOR 15: PRINT nme$
5180 GOSUB figurebalance
     acctbal# = curbal# + pastdue#
     GOSUB 61000: NEXRECORD = LOF(2) / 96 + 1
5190 LSET tbal$ = MKS$(acctbal#)
     LSET TDATE$ = LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 2)
     LSET TAMT$ = MKS$(ctax# + ptax# + pctax# + pptax#)
     LSET ttype$ = MKI$(5): LSET TDESC$ = "Pen./Int."
     LSET ttaxyr$ = MKI$(0): LSET taprin$ = MKS$(0)
     LSET taint$ = MKS$(ctax# + ptax#)
     LSET taadv$ = MKS$(0): LSET tacol$ = MKS$(0)
     LSET talate$ = MKS$(pctax# + pptax#)
     LSET tdisamt$ = MKS$(0)
     LSET TMREC$ = MKS$(CVS(MREC$)): PUT 2, NEXRECORD
5200 IF CVS(ftr$) = 0 THEN LSET ftr$ = MKS$(NEXRECORD): LSET ltr$ = MKS$(NEXRECORD): PUT 1, CVS(MREC$): CLOSE 2: GOTO 5220
5210 prev = CVS(ltr$): LSET ltr$ = MKS$(NEXRECORD): PUT 1, CVS(MREC$): GET 2, prev: LSET TNEXTTR$ = MKS$(NEXRECORD): PUT 2, prev: CLOSE 2
5220 REM put in temp file for printout
5230 OPEN "R", 2, "tbfee.dat", 128
      FIELD 2, 15 AS FEEMETER$, 30 AS FEENAME$, 8 AS CFEEAMT$, 8 AS pfeeamt$, 8 AS cpenamt$, 8 AS ppenamt$
5240 LSET FEEMETER$ = m$: LSET FEENAME$ = nme$: LSET CFEEAMT$ = MKD$(ctax#)
     LSET pfeeamt$ = MKD$(ptax#): LSET cpenamt$ = MKD$(pctax#)
     LSET ppenamt$ = MKD$(pptax#)
5250 PUT 2, LOF(2) / 128 + 1: CLOSE 2: FEEFLAG = 1: RETURN

5700 REM open and print fee charged report
5705 IF FEEFLAG = 0 THEN RETURN
5710 LOCATE 18, 1: COLOR 11: PRINT STRING$(79, "-")
5720 LOCATE 19, 1: COLOR 11: PRINT "Get the Printer Ready for Printing, Press Any Key When Ready "; : fl = 1: GOSUB 62000
5730 OPEN "R", 2, "tbfee.dat", 128
     FIELD 2, 15 AS FEEMETER$, 30 AS FEENAME$, 8 AS CFEEAMT$, 8 AS pfeeamt$, 8 AS cpenamt$, 8 AS ppenamt$
5740 GOSUB 5900
5750 FOR x = 1 TO LOF(2) / 128: GET 2, x
5760 IF lc >= 56 THEN LPRINT CHR$(12); : GOSUB 5900
5770 LPRINT FEEMETER$; TAB(17); LEFT$(FEENAME$, 25); TAB(45); USING "####.##"; CVD(CFEEAMT$);
      LPRINT TAB(54); USING "####.##"; CVD(pfeeamt$);
       LPRINT TAB(63); USING "####.##"; CVD(cpenamt$);
        LPRINT TAB(72); USING "####.##"; CVD(ppenamt$)
     T# = T# + CVD(CFEEAMT$): lc = lc + 1
     t1# = t1# + CVD(pfeeamt$)
     t2# = t2# + CVD(cpenamt$)
     t3# = t3# + CVD(ppenamt$)
5780 NEXT x
5790 LPRINT STRING$(79, "-"): LPRINT "Total CURRENT Interest Charged .. "; TAB(49); USING "$$####,#.##"; T#
                              LPRINT "Total PAST Interest Charged ..... "; TAB(49); USING "$$####,#.##"; t1#
                              LPRINT "Total CURRENT Penalty Charged ... "; TAB(49); USING "$$####,#.##"; t2#
                              LPRINT "Total PAST Penalty Charged ...... "; TAB(49); USING "$$####,#.##"; t3#
                              t2# = 0: t3# = 0
                              t1# = 0: T# = 0: LPRINT CHR$(12); : RETURN
5900 LPRINT TAB(30); "Interest Charged Report": LPRINT "Date: "; DATE$: LPRINT
5910 LPRINT "Account # "; TAB(17); "Customer Name "; TAB(45); "Cur Int."; TAB(54); "Past Int."; TAB(64); "Cur Pen"; TAB(73); "Past Pen"
     LPRINT STRING$(79, "="): lc = 4: RETURN

curtax:
ctax# = 0: pctax# = 0
IF CVS(curtaxamt$) >= .01 THEN
 ctax# = (curpercent * CVS(curtaxamt$)) / 100
 pctax# = (pcurpercent * CVS(curtaxamt$)) / 100
 ELSE
 ctax# = 0
 cptax# = 0
 RETURN
END IF
ctax# = INT((ctax# * 100) + .5) / 100
pctax# = INT((pctax# * 100) + .5) / 100
LSET curintamt$ = MKS$(CVS(curintamt$) + ctax#)
LSET curintamt$ = MKS$(CVS(curintamt$) + pctax#)
PUT 1, CVS(MREC$)
RETURN

pasttax:
ptax# = 0: pptax# = 0
FOR lloop = 0 TO 9
INTEREST#(lloop) = 0
 IF CVS(ptaxamt$(lloop)) >= .01 THEN
 INTEREST#(lloop) = (pastpercent * CVS(ptaxamt$(lloop)) / 100)
 interest1#(lloop) = (ppastpercent * CVS(ptaxamt$(lloop)) / 100)
 ELSE
 GOTO past1
END IF
 INTEREST#(lloop) = INT((INTEREST#(lloop) * 100) + .5) / 100
 interest1#(lloop) = INT((intest1#(lloop) * 100) + .5) / 100

 LSET pintamt$(lloop) = MKS$(CVS(pintamt$(lloop)) + INTEREST#(lloop))
 LSET pintamt$(lloop) = MKS$(CVS(pintamt$(lloop)) + intest1#(lloop))
 ptax# = ptax# + INTEREST#(lloop)
 pptax# = pptax# + interest1#(lloop)
past1:
 NEXT lloop

 ptax# = INT((ptax# * 100) + .5) / 100
 pptax# = INT((pptax# * 100) + .5) / 100
 PUT 3, CVS(MREC$): RETURN


6000 GOTO 10

10000 CLOSE 9: OPEN "R", 9, "tbsum.dat", 256
      FIELD 9, 8 AS wacct$, 30 AS wname$, 8 AS wreal$, 8 AS wpers$, 8 AS wsrcit$, 8 AS wtaxgen$, 8 AS wtaxfire$, 8 AS wcartag$, 4 AS wbillnum$, 8 AS wtaxowed$, 8 AS wlate$, 4 AS wtaxyear$
      RETURN

60000 CLOSE 1: OPEN "tbcust.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #1 LEN = 512
60010 FIELD 1, 9 AS ACCT$, 4 AS nu$, 25 AS firstnam$, 35 AS lnam$, 35 AS A1$, 35 AS A2$, 20 AS cy$, 2 AS st$, 10 AS zp$, 12 AS tp$
      FIELD 1, 187 AS dummy$, 39 AS PDESC$, 22 AS PMAP$, 35 AS psub$, 4 AS NOTUSED$, 35 AS PCOUNTY$, 4 AS psize$, 8 AS PLAND$, 8 AS PBLDG$
      FIELD 1, 342 AS dummy1$, 8 AS pbus$, 8 AS pres$, 8 AS ppub$, 4 AS persexp$, 4 AS SCEXP$, 4 AS homeexp$, 4 AS useexp$
      FIELD 1, 382 AS dummy2$, 4 AS curlateamt$, 4 AS curtaxamt$, 4 AS curintamt$, 4 AS curadvamt$, 4 AS curcolamt$, 4 AS ctaxnotice$
      FIELD 1, 406 AS dummy3$, 4 AS ftr$, 4 AS ltr$, 11 AS ss$, 1 AS late$, 8 AS PLOT$, 13 AS PIN$, 62 AS blank$

60025 CLOSE 3: OPEN "tbcust1.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #3 LEN = 256
      FOR tloop = 0 TO 9
      FIELD 3, (24 * tloop) AS dummy4$, 4 AS ptaxamt$(tloop), 4 AS pintamt$(tloop), 4 AS PADVAMT$(tloop), 4 AS PCOLAMT$(tloop), 4 AS plateamt$(tloop), 4 AS ptaxnotice$(tloop): NEXT tloop
60030 RETURN
61000 CLOSE 2: OPEN "tbartran.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #2 LEN = 96
61010 FIELD 2, 6 AS TDATE$, 4 AS TAMT$, 2 AS ttaxyr$, 20 AS TDESC$, 2 AS ttype$, 4 AS taprin$, 4 AS taint$, 4 AS taadv$, 4 AS tacol$, 4 AS talate$, 4 AS TMREC$, 4 AS TNEXTTR$, 4 AS tbal$, 4 AS tdisamt$, 26 AS tblank$
61020 RETURN




62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO

62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
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
    

