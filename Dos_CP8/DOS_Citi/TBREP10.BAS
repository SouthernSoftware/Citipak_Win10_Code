10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbcustrep
30 REM Customer Report Menu
40 REM Last Revision Date: 8-24-87
50 REM by Nick Picerno
   DECLARE FUNCTION QPTRIM$ (a$)
   DECLARE FUNCTION num2date$ (Num%)
   DECLARE FUNCTION Date2Num% (Dat$)

60 CLS : CLEAR
   '$INCLUDE: 'lcTAX.BI'

    TBpath$ = "F:\"                    ' change to "F:\" Before compiling

     CLOSE : OPEN "R", 1, TBpath$ + "tbyear.dat", 4: FIELD 1, 4 AS year$: GET 1, 1
     CurrentYear$ = year$: CLOSE
     ON ERROR GOTO networkerror
     REDIM TBRECust(1) AS TBRECust
     REDIM TBPPCust(1) AS TBPPCust
     REDIM TBTrans(1) AS TBTrans


3000 REM
3010 SCREN$ = "2.03.00": CLS : GOSUB 63500
3020 COLOR 11: LOCATE 1, 32: PRINT "Daily Collection Journal";
3025 LOCATE 3, 1: COLOR 11
     PRINT "Enter the Starting Collections Date (mmddyyyy) (ESC-End) .. "; : fl = -8
     GOSUB 62000
     IF cf = 1 THEN RUN "tbtcustr"
     IF LEN(In$) <> 8 THEN 3025
     Cdate$ = In$
     Cdate% = Date2Num%(In$)
3030 LOCATE 4, 1: COLOR 11
     PRINT "Enter the Ending Collections Date (mmddyyyy) (ESC-End) .... "; : fl = -8
     GOSUB 62000
     IF cf = 1 THEN RUN "tbtcustr"
     IF LEN(In$) <> 8 THEN 3030
     Edate$ = In$
     Edate% = Date2Num%(In$)
     IF Edate% < Cdate% THEN
      PRINT CHR$(7);
      PRINT "Incorrect Date Order : Press <Enter> to Try Again!"
      fl = 1
      GOSUB 62000
      GOTO 3000
     END IF
     GOSUB 3900
     LOCATE 5, 1
     COLOR 15
     PRINT "Now Extracting Data : Please Wait!!!"
3075 CLOSE
     GOSUB 61000
     StartRecord! = LOF(TBTransFile) / LEN(TBTrans(1)) - 7500
     IF StartRecord! < 1 THEN StartRecord! = 1
     FOR Cnt! = StartRecord! TO LOF(TBTransFile) / LEN(TBTrans(1))
     GET TBTransFile, Cnt!, TBTrans(1)
     IF (TBTrans(1).TransDate >= Cdate% AND TBTrans(1).TransDate <= Edate%) AND (TBTrans(1).TrType = 2 OR TBTrans(1).TrType = 5) THEN
      IF LC >= 55 THEN LPRINT CHR$(12); : GOSUB 3900
      LPRINT TBTrans(1).TrCustomerRecord;
       LPRINT TAB(11); TBTrans(1).TrSource;
        LPRINT TAB(13); USING "$$####,#.##"; TBTrans(1).TransAmt;
         LPRINT TAB(30); USING "$$####,#.##"; TBTrans(1).TrTax;
          LPRINT TAB(42); USING "$$####,#.##"; TBTrans(1).TrInterest;
           LPRINT TAB(55); USING "$$####,#.##"; TBTrans(1).TrPenalty;
            IF TBTrans(1).TrType = 5 THEN LPRINT TAB(72); "YES" ELSE LPRINT TAB(73); "NO"
            LC = LC + 1
            IF TBTrans(1).TrType = 2 THEN
            TotalTrans# = TotalTrans# + TBTrans(1).TransAmt
            TotalTax# = TotalTax# + TBTrans(1).TrTax
            TotalPen# = TotalPen# + TBTrans(1).TrPenalty
            TotalInt# = TotalInt# + TBTrans(1).TrInterest
            ELSE
            TotalRTrans# = TotalRTrans# + TBTrans(1).TransAmt
            TotalRTax# = TotalRTax# + TBTrans(1).TrTax
            TotalRPen# = TotalRPen# + TBTrans(1).TrPenalty
            TotalRInt# = TotalRInt# + TBTrans(1).TrInterest
            END IF

     END IF

     NEXT Cnt!
     Page = Page + 1
     LPRINT CHR$(12);
     LPRINT
     LPRINT "Totals Page for Collections Start Date: "; LEFT$(Cdate$, 2); "/"; MID$(Cdate$, 3, 2); "/"; RIGHT$(Cdate$, 2); TAB(65); "Page #"; Page
     LPRINT "  Totals Page for Collections End Date: "; LEFT$(Edate$, 2); "/"; MID$(Edate$, 3, 2); "/"; RIGHT$(Edate$, 2)
     LPRINT
     LPRINT
     LPRINT "Taxes Paid:"
     LPRINT "Grand Total Amount Received: "; USING "$$######,#.##"; TotalTrans#
     LPRINT "Total Tax Amount ........... "; USING "$$######,#.##"; TotalTax#
     LPRINT "Total Penalty Amount ....... "; USING "$$######,#.##"; TotalPen#
     LPRINT "Total Interest Amount ...... "; USING "$$######,#.##"; TotalInt#
     LPRINT
     LPRINT
     LPRINT "Taxes Released:"
     LPRINT "Grand Total Amount Released: "; USING "$$######,#.##"; TotalRTrans#
     LPRINT "Total Tax Amount ........... "; USING "$$######,#.##"; TotalRTax#
     LPRINT "Total Penalty Amount ....... "; USING "$$######,#.##"; TotalRPen#
     LPRINT "Total Interest Amount ...... "; USING "$$######,#.##"; TotalRInt#
     LPRINT CHR$(12);
     CLOSE
     RUN "TBtCustR"
     
3900 Page = Page + 1
     LPRINT TAB(33); "Collections Journal"
     LPRINT
     LPRINT "Collection Start Date: "; LEFT$(Cdate$, 2); "/"; MID$(Cdate$, 3, 2); "/"; RIGHT$(Cdate$, 2); TAB(65); "Page #"; Page
     LPRINT "  Collection End Date: "; LEFT$(Edate$, 2); "/"; MID$(Edate$, 3, 2); "/"; RIGHT$(Edate$, 2)
     LPRINT
     LPRINT "Acct #"; TAB(10); "SRC"; TAB(15); "Total Paid"; TAB(30); "Tax Paid"; TAB(42); "Inter't Pd"; TAB(55); "Penalty Pd"; TAB(70); "Released"
     LPRINT STRING$(79, "-")
     LC = 6: RETURN


60000  IF taxtype$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBpath$ + "TBRTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSE
         TBFile = FREEFILE
         OPEN TBpath$ + "TBPTCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         END IF
        RETURN


61000 TBTransFile = FREEFILE
      OPEN TBpath$ + "TBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTransFile LEN = LEN(TBTrans(1))
61020 RETURN

62000 COLOR 14: cf = 0: In$ = "": RO = CSRLIN: PO = POS(Xy): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
      IF LEN(z$) < 2 THEN 62065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN cf = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN cf = 12: RETURN

62065 IF ASC(z$) = 0 THEN 62030
62070 IF LEN(z$) > 1 THEN cf = ASC(RIGHT$(z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
62080 IF z$ = " " THEN 62130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 62000
62090 IF z$ <> CHR$(8) THEN 62130 ELSE IF ZL = 0 THEN 62030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 62110 ELSE IF z$ = "," THEN 62120
62100 IF z$ = "." THEN ZD = 0: GOTO 62110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
62110 In$ = LEFT$(In$, LEN(In$) - 1)
62120 ZL = ZL - 1: GOTO 62030
62130 IF ABS(fl) = ZL THEN 62030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 62180
62140 IF z$ = "." AND ZD = 0 THEN ZD = 1: GOTO 62180
62150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF z$ < "0" OR z$ > "9" THEN 62030
62180 PRINT z$; : In$ = In$ + z$: ZL = ZL + 1
62190 IF ABS(fl) = 1 THEN 62050 ELSE 62030
62200 LOCATE RO, PO: PRINT In$: RETURN
62210 GOSUB 62000: IF LEN(In$) THEN In$ = CHR$(ASC(In$) AND 223)
62220 RETURN
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; DATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN


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

