'$INCLUDE: 'LCTax.Bi'
10 CLS
   CLEAR
   
           
  'TBPath$ = "F:\"                    ' change to "F:\" Before compiling
   TBPath$ = "C:\qb45\lc\"

   REDIM TBRECust(1) AS TBRECust
   REDIM TBRECustBal(1) AS TBREBAL
   REDIM TBPPCust(1) AS TBPPCust
   REDIM TBPPCustBal(1) AS TBPPBal
   REDIM TBTrans(1) AS TBTrans
   REDIM TBVeh(1) AS PPVehType
   TaxType$ = "P"
   GOSUB 60000
   FOR Cnt! = 1 TO LOF(TBFile) \ LEN(TBPPCust(1))
   GET TBFile, Cnt!, TBPPCust(1)
   LOCATE 5, 1: PRINT "Now Initializing New Fields For Personal Property"
   PRINT "Record #"; Cnt!
   TBPPCust(1).FirstVeh = 0
   TBPPCust(1).LastVeh = 0
   TBPPCust(1).Lname2 = ""
   TBPPCust(1).Fname2 = ""
   TBPPCust(1).HeavyEquip = 0
   TBPPCust(1).Boats = 0
   TBPPCust(1).BusFurn = 0
   PUT TBFile, Cnt!, TBPPCust(1)
   NEXT Cnt!
   CLOSE
   END

60000  IF TaxType$ = "R" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TBRECUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBRECust(1))
        ELSEIF TaxType$ = "P" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TBPPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
         TBVFile = FREEFILE
         OPEN TBPath$ + "TBVEH.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBVFile LEN = LEN(TBVeh(1))

        ELSEIF TaxType$ = "S" THEN
         TBFile = FREEFILE
         OPEN TBPath$ + "TBSPCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TBFile LEN = LEN(TBPPCust(1))
        END IF
         RETURN

62000 cf = 0: COLOR 14: in$ = "": RO = CSRLIN: PO = POS(X): ZD = 0: ZS = 0: ZL = 0: IF fl = ZD THEN fl = 1
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
63000 RETURN
63500 COLOR 10: LOCATE 1, 1: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN


