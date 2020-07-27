10 REM (C)1987 SOUTHERN SOFTWARE SYSTEMS         by Nick Picerno
20 REM tbbilling
30 REM Tax Billing Menu
40 REM Last Revision Date:11-29-88
50 REM by Nick Picerno
60 CLS : CLEAR
   '$INCLUDE: 'lcTAX.BI'
   TBPath$ = "F:\"                    ' change to "F:\" Before compiling

    REDIM TBPPCust(1) AS TBPPCust

    
4000 CLS : scren$ = "3.04.00"
     GOSUB 63500: COLOR 11
     LOCATE 3, 1: PRINT "Now Setting Up New Boat and Mobile Home Detail Record"
     GOSUB 60000
      EndofFile = LOF(TBFile) / LEN(TBPPCust(1))
      
     FOR x = 1 TO EndofFile
        GET 1, x, TBPPCust(1)
        TBPPCust(1).MobileHome = 0
        TBPPCust(1).Boats = 0
        TBPPCust(1).BoatFirstTrans = 0
        TBPPCust(1).BoatLastTrans = 0
        TBPPCust(1).MHFirstTrans = 0
        TBPPCust(1).MHLastTrans = 0
        PUT 1, x, TBPPCust(1)
        NEXT x
       CLOSE
       END


60000    TBFile = FREEFILE
         OPEN TBPath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         RETURN

60025 tbbalfile = FREEFILE
      IF TaxType$ = "R" THEN
      OPEN TBPath$ + "TBRTBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #tbbalfile LEN = LEN(tbrecustbal(1))
      ELSE
      OPEN TBPath$ + "TBPTBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #tbbalfile LEN = LEN(tbppcustbal(1))
      END IF
60030 RETURN

61000 TBTranFile = FREEFILE
      OPEN TBPath$ + "TBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBTranFile LEN = LEN(TBTrans(1))
61020 RETURN




62000 COLOR 14: CF = 0: in$ = "": RO = CSRLIN: PO = POS(x): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN CF = 1: RETURN
      IF z$ = CHR$(9) THEN CF = 2: RETURN
      IF LEN(z$) < 2 THEN 62065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN CF = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN CF = 12: RETURN
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
63500 LOCATE 1, 1: COLOR 10: PRINT "Sys. Date: "; SYSDATE$; TAB(70); scren$: PRINT STRING$(79, "_"): RETURN
    

