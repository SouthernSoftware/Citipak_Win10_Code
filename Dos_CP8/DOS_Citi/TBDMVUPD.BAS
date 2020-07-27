
'$INCLUDE: 'LCTax.Bi'

10 CLS
'   ON ERROR GOTO networkerror
           
  TBPath$ = "F:\"                    ' change to "F:\" Before compiling
'   TBPath$ = "C:\qb45\lc\"

   REDIM TBRECust(1) AS TBRECust
   REDIM TBRECustBal(1) AS TBREBAL
   REDIM TBPPCust(1) AS TBPPCust
   REDIM TBPPCustBal(1) AS TBPPBal
   REDIM TBVeh(1) AS PPVehType
   DIM VREC!(250)

10000 SCREN$ = "1.10.00": CLS : GOSUB 63500
10010 LOCATE 1, 32: COLOR 11: PRINT "Update Weight Using DMV Data"; : LOCATE 3, 1
     taxtype$ = "P"
     LOCATE 3, 1
     PRINT "This function requires a Password to Process"
     PRINT
     PRINT "Enter the Password Now (ESC-Cancel) .. "; : fl = 6
     GOSUB 62000: IF cf = 1 THEN 10

     Password$ = UCASE$(IN$)
     IF Password$ <> "SOSOFT" THEN 10000

     TBVFile = FREEFILE
     OPEN TBPath$ + "TBVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))

     OPEN "I", 10, TBPath$ + "dmv2003.txt"

GetDMVData:
   LINE INPUT #10, a$

   IF EOF(10) THEN
    CLOSE
    END
   END IF
   COLOR 11: LOCATE 5, 1
   PRINT "DMV Record                                                            "
   PRINT "VIN    : "; MID$(a$, 1246, 22)
   VIN$ = RTRIM$(MID$(a$, 1246, 22))
   VIN$ = LTRIM$(VIN$)
   PRINT "Veh Wt : "; MID$(a$, 1676, 6)
   GOSUB UpdateTheAccount
   GOTO GetDMVData
   
UpdateTheAccount:
     FOR VCnt! = 1 TO (LOF(TBVFile) \ LEN(TBVeh(1)))
       GET TBVFile, VCnt!, TBVeh(1)
      IF RTRIM$(TBVeh(1).VehVin) = VIN$ THEN
       TBVeh(1).VehWght = VAL(MID$(a$, 1676, 6))
       PUT TBVFile, VCnt!, TBVeh(1)
       RETURN
      END IF
     NEXT VCnt!
     RETURN

60000  IF taxtype$ = "R" THEN
         CLOSE TBFile
         TBFile = FREEFILE
         OPEN TBPath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSEIF taxtype$ = "P" THEN
         CLOSE TBFile
         TBFile = FREEFILE
         OPEN TBPath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         CLOSE TBVFile
         TBVFile = FREEFILE
         OPEN TBPath$ + "TBVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))
        ELSEIF taxtype$ = "S" THEN
         CLOSE TBFile
         TBFile = FREEFILE
         OPEN TBPath$ + "TBSPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
        END IF
60025 CLOSE TBBalFile
      TBBalFile = FREEFILE
      IF taxtype$ = "R" THEN
      OPEN TBPath$ + "TBREBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBRECustBal(1))
      ELSE
      OPEN TBPath$ + "TBPPBal.dat" FOR RANDOM ACCESS READ WRITE SHARED AS #TBBalFile LEN = LEN(TBPPCustBal(1))
      END IF
60030 RETURN
        

62000 cf = 0: COLOR 14: IN$ = "": RO = CSRLIN: PO = POS(X): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(fl), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
      IF LEN(z$) < 2 THEN 62065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN cf = 11: RETURN
62065 IF ASC(z$) = 0 THEN 62030
62070 IF LEN(z$) > 1 THEN cf = ASC(RIGHT$(z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
62080 IF z$ = " " THEN 62130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 62000
62090 IF z$ <> CHR$(8) THEN 62130 ELSE IF ZL = 0 THEN 62030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 62110 ELSE IF z$ = "," THEN 62120
62100 IF z$ = "." THEN ZD = 0: GOTO 62110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
62110 IN$ = LEFT$(IN$, LEN(IN$) - 1)
62120 ZL = ZL - 1: GOTO 62030
62130 IF ABS(fl) = ZL THEN 62030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 62180
62140 IF z$ = "." AND ZD = 0 THEN ZD = 1: GOTO 62180
62150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF z$ < "0" OR z$ > "9" THEN 62030
62180 PRINT z$; : IN$ = IN$ + z$: ZL = ZL + 1
62190 IF ABS(fl) = 1 THEN 62050 ELSE 62030
62200 LOCATE RO, PO: PRINT IN$: RETURN
62210 GOSUB 62000: IF LEN(IN$) THEN IN$ = CHR$(ASC(IN$) AND 223)
62220 RETURN
63000 RETURN
63500 COLOR 10: LOCATE 1, 1: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN




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
          CASE IS = 6
               RESUME NEXT

          CASE ELSE
               CLS
               LOCATE 15, 1: COLOR 28: PRINT "Error Message: "; : COLOR 15: PRINT "Error "; Number; " Has Occured in Line "; ERL
               PRINT "Write Down Message and Hit <ESC> Key to exit."
               fl = 1: GOSUB 62000: IF cf = 1 THEN CLOSE : GOTO 10 ELSE GOSUB 62000
            
     END SELECT

