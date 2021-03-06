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
   tbpath$ = ""                    ' change to "F:\" Before compiling

   CLOSE
   OPEN "R", 1, tbpath$ + "tbyear.dat", 4: FIELD 1, 4 AS Year$: GET 1, 1
   CurrentYear$ = Year$: CLOSE 1
'   ON ERROR GOTO networkerror


     REDIM TransLink!(500)

     REDIM TBPPCust(1) AS TBPPCust
     REDIM TBPPCustBal(1) AS TBPPBal
     REDIM TBTrans(1) AS TBTrans
     REDIM TBPage(1) AS TBPage1         'Page 1 is Personal Property
     REDIM TBVeh(1) AS PPVehType
     REDIM Nada(1) AS Nada
     
     WIDTH LPRINT 236
     
100 REM begin program
110 scren$ = "2.00.00": SYSDATE$ = DATE$
120 CLS : GOSUB 63500
1000
1100 scren$ = "2.01.01": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 32: PRINT "NADA FILE PREP"
     LOCATE 3, 1
     COLOR 12
     PRINT "WARNING!!!  Running this function will change the values"
     PRINT "Are you sure you want to Continue (Y/N) .. "; : fl = 1: GOSUB 62210
     IF in$ = "Y" THEN 1101
     RUN "tbccm"

1101 scren$ = "2.01.01": CLS : GOSUB 63500
     COLOR 11: LOCATE 1, 32: PRINT "NADA FILE PREP"
     LOCATE 3, 1


     PRINT "Do you wish to clear the Personal Property Values "
     PRINT "that exist in each customers file before updating"
     PRINT "with the NADA Priced Values"
     PRINT
     PRINT "Answer (Y/N) or (ESC-Abort Process) : "; : fl = 1: GOSUB 62210
     IF cf = 1 THEN CLOSE : RUN "TBCCM"
     IF NOT (in$ = "Y" OR in$ = "N") THEN 1100
     IF in$ = "Y" THEN
      GOSUB initvalues
     END IF
     Taxtype$ = "P"
     GOSUB 60000
    'CLEAR SCREEN AREA
     FOR LL = 3 TO 23
     LOCATE LL, 1
     PRINT STRING$(79, " ")
     NEXT LL
     LOCATE 3, 1

     
1150 IF LOF(1) / 512 = 0 THEN CLOSE : RUN "tbccustr"
      OPEN tbpath$ + "NADAPR02.TXT" FOR INPUT AS #10
     
MAINBODY:

     LINE INPUT #10, a$
     TCNT& = TCNT& + 1
     IF EOF(10) THEN GOTO ENDPROCESS
     CarVin$ = MID$(a$, 9, 17)
     CarVin$ = RTRIM$(CarVin$)
     VinL = LEN(CarVin$)
     Cardesc$ = MID$(a$, 134, 26)
     CarValue$ = MID$(a$, 111, 6)
     LOCATE 5, 1:
     COLOR 11
     PRINT "Now Looking For "; : COLOR 15: PRINT Cardesc$
     COLOR 11
     PRINT "          VIN # "; : COLOR 15: PRINT CarVin$

     FOR Cnt& = 1 TO LOF(TBVFile) / LEN(TBVeh(1))
     GET TBVFile, Cnt&, TBVeh(1)
     IF LEFT$(TBVeh(1).VehVin, VinL) = LEFT$(CarVin$, VinL) THEN
      TBVeh(1).VehValue = VAL(CarValue$)
      FCNT& = FCNT& + 1
      PUT TBVFile, Cnt&, TBVeh(1)
      EXIT FOR
     END IF

     

1230 NEXT Cnt&
     LOCATE 10, 1: PRINT "VEHICLES PROCESSED: "; TCNT&
'     LOCATE 11, 1: PRINT "  VEHICLES UPDATED: "; FCNT&
'     LOCATE 12, 1: PRINT "      % OF SUCCESS: "; USING "###.##"; FCNT& / TCNT& * 100

     GOTO MAINBODY
ENDPROCESS:
      CLOSE
     LOCATE 14, 1
     PRINT "ALL DONE!"
     SLEEP 15
1250  RUN "tbccm"

initvalues:
Taxtype$ = "P"
GOSUB 60000
     LOCATE 3, 1
     FOR LL = 3 TO 23: LOCATE LL, 1: PRINT STRING$(79, 32): NEXT LL
     LOCATE 3, 1
     COLOR 11
     PRINT "Now Clearing Old Car Values : Please Wait!!!"
     FOR Cnt& = 1 TO LOF(TBVFile) / LEN(TBVeh(1))
     GET TBVFile, Cnt&, TBVeh(1)
      TBVeh(1).VehValue = 0
     PUT TBVFile, Cnt&, TBVeh(1)
     NEXT Cnt&
     CLOSE
          LOCATE 3, 1
          PRINT STRING$(79, 32)
     RETURN
     



60000  IF Taxtype$ = "R" THEN
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
      IF Taxtype$ = "R" THEN
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

