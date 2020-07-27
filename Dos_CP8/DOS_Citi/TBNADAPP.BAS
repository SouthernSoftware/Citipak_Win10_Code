10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbccustrp
30 REM Customer Report Menu
40 REM Last Revision Date: 8-24-87
50 REM by Nick Picerno

 ' Must be compiled using the compile directory
   DEFINT A-Z
   DECLARE FUNCTION QPTrim$ (a$)
   DECLARE FUNCTION num2date$ (Num%)
   DECLARE FUNCTION Date2Num% (Dat$)
   DECLARE FUNCTION WaitKey% ()
   DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
   DECLARE SUB PrintTitle (Title$)
   DECLARE SUB PrintLoading ()
   DECLARE SUB ShowCursor ()
   DECLARE SUB HideCursor ()

   
'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'Qscr.BI'
'$INCLUDE: 'SetCnf.BI'
'$INCLUDE: 'lcTAX.BI'
   
60 CLS :
   tbpath$ = "F:\"                    ' change to "F:\" Before compiling
   CLOSE
   OPEN "R", 1, tbpath$ + "tbyear.dat", 4: FIELD 1, 4 AS Year$: GET 1, 1
   CurrentYear$ = Year$: CLOSE 1
   ON ERROR GOTO networkerror


     REDIM TransLink!(500)

     REDIM TBPPCust(1) AS TBPPCust
     REDIM TBPPCustBal(1) AS TBPPBal
     REDIM TBTrans(1) AS TBTrans
     REDIM TBPage(1) AS TBPage1         'Page 1 is Personal Property
     REDIM TBVeh(1) AS PPVehType
     REDIM NADA(1) AS NADA
     
     WIDTH LPRINT 236
     
100 REM begin program
110 scren$ = "2.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500
1000
1100 scren$ = "2.01.01": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 32: PRINT "NADA FILE PREP"
     LOCATE 3, 1
     taxtype$ = "P"
     GOSUB 60000
1150 IF LOF(1) / 512 = 0 THEN CLOSE : RUN "tbccustr"
     
     OPEN tbpath$ + "NADA" + RIGHT$(DATE$, 4) + ".TXT" FOR RANDOM ACCESS READ WRITE SHARED AS #10 LEN = 322
     CLOSE 10
     df$ = "del " + tbpath$ + "NADA" + RIGHT$(DATE$, 4) + ".TXT"
     SHELL df$
     OPEN tbpath$ + "NADA" + RIGHT$(DATE$, 4) + ".TXT" FOR OUTPUT AS #10
     
1160 OPEN tbpath$ + "TBPPname.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #4 LEN = 24
        FIELD 4, 20 AS LastNam$, 4 AS Mrec$: EndofFile = LOF(4) / 24
     
1165 FOR District = 1 TO 10                     'MAIN LOOP
     LOCATE 3, 1: COLOR 11
     PRINT "NOW SPOOLING OUT DISTRICT # "; District
     IF District >= 2 AND District <= 8 THEN 1231       'No Districts 2 thru 8 in Personal Prop
1170 FOR X = 1 TO EndofFile
     GET 4, X
     GET 1, CVS(Mrec$), TBPPCust(1)
     Account! = CVS(Mrec$)
     
     
1191 IF TBPPCust(1).Deleted = "Y" THEN 1230
     IF VAL(TBPPCust(1).District) <> District THEN 1230
     
     
      VehRecord! = TBPPCust(1).FirstVeh
      WHILE VehRecord! <> 0
        GET TBVFile, VehRecord!, TBVeh(1)

        IF LEN(QPTrim$(TBVeh(1).VehVin)) > 0 AND TBVeh(1).VehValue = 0 AND TBVeh(1).VehTyp <> "D" THEN

         yr$ = RIGHT$(TBVeh(1).VehYear, 2)
         NADA(1).Blank = " "
         NADA(1).NCIC = "     "
         NADA(1).Year = yr$
         NADA(1).VIN = TBVeh(1).VehVin
         NADA(1).State = "  "
         NADA(1).VehRecord = LTRIM$(STR$(VehRecord!))
         NADA(1).CusRecord = LTRIM$(STR$(Account!))
         NADA(1).NotUsed = STRING$(59, " ")
         NADA(1).NADAFILLER = STRING$(222, " ")
         NADA(1).ValPeriod = "01" + RIGHT$(DATE$, 2)
         NADA(1).NadaFiller1 = STRING$(378, " ")

         NadaLine$ = NADA(1).Blank + NADA(1).NCIC + NADA(1).Year + NADA(1).VIN + NADA(1).State + NADA(1).VehRecord + NADA(1).CusRecord + NADA(1).NotUsed + NADA(1).NADAFILLER + NADA(1).ValPeriod + NADA(1).NadaFiller1
         NADA = NADA + 1
         PRINT #10, NadaLine$
        END IF
        VehRecord! = TBVeh(1).VehNext
      WEND

1230 NEXT X
1231 NEXT District
1240 CLOSE
     LOCATE 5, 1
     PRINT "NADA FILE CREATED"
     PRINT "File Name: NADA" + RIGHT$(DATE$, 4) + ".TXT"
     SLEEP 10
1250  RUN "tbccm"


60000  IF taxtype$ = "R" THEN
         TBFILE = FREEFILE
         OPEN tbpath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFILE LEN = LEN(TBRECust(1))
        ELSE
         TBFILE = FREEFILE
         OPEN tbpath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFILE LEN = LEN(TBPPCust(1))
         CLOSE TBVFile
         TBVFile = FREEFILE
         OPEN tbpath$ + "TBVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))

         END IF
        RETURN

60025 TBBalFile = FREEFILE
      IF taxtype$ = "R" THEN
      OPEN tbpath$ + "TBREBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN tbpath$ + "TBPPBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
      END IF
60030 RETURN

61000 TBTransFile = FREEFILE
      OPEN tbpath$ + "TBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTransFile LEN = LEN(TBTrans(1))
61020 RETURN

62000 COLOR 14: cf = 0: in$ = "": RO = CSRLIN: PO = POS(XY): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
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
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 61 THEN cf = 13: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 62 THEN cf = 14: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 63 THEN cf = 15: RETURN
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
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SYSDATE$; TAB(70); scren$: PRINT STRING$(79, "_"): RETURN


FigureBalance:
     CurBal# = 0: PastDue# = 0
     CurBal# = TBPPCust(1).CurYrPersPropTax + TBPPCust(1).CurYrFarmEquipTax + TBPPCust(1).CurYrMachToolsTax + TBPPCust(1).CurYrMerchCapTax + TBPPCust(1).CurYrMobileHomeTax + TBPPCust(1).CurYrLateFeeTax + TBPPCust(1).CurYrIntAmount + TBPPCust(1). _
CurYrPenAmount
     CurBal# = INT((CurBal# * 100) + .5) / 100
     FOR LL = 1 TO 10
     PastDue# = PastDue# + TBPPCustBal(1).PerPropTaxAmt(LL) + TBPPCustBal(1).FarmEquipAmt(LL) + TBPPCustBal(1).MachToolsAmt(LL) + TBPPCustBal(1).MerchCapAmt(LL) + TBPPCustBal(1).MobileHomeAmt(LL) + TBPPCustBal(1).LateFeeAmt(LL) + TBPPCustBal(1). _
IntAmt(LL) + TBPPCustBal(1).PenAmt(LL)
     PastDue# = INT((PastDue# * 100) + .5) / 100
     NEXT LL
     RETURN

Deleted:
    LOCATE 23, 1: COLOR 15
    PRINT "This Account Has Been Deleted! Press <ENTER> to Continue"; :
    fl = 1
    GOSUB 62000
    LOCATE 23, 1: PRINT STRING$(79, 32)
    RETURN

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

