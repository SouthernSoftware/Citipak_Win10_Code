10 REM new hand held program
20 REM Town of Granite Quarry, NC
     DECLARE FUNCTION Date2Num% (DateString$)
    '$INCLUDE: 'ubsetup.bi'
     REDIM UBPC3000RDRec(1) AS UBPC3000ReadRecType

     CLS
     PRINT "PC 3000 H/H Meter"
     PRINT " Reading System"
     PRINT
     PRINT "(C)1997 Southern"
     PRINT "    Software "
     PRINT
1030 PRINT "Enter Route # "; : FL = -2: GOSUB 62000: IF cf = 1 THEN END
1035 Book = VAL(in$)
1040 CLS : GOSUB 60000
     lastrecord = LOF(1) / UBPC3000RdRecLen
     CURRECORD = 1
     
1059 IF scancode = 1 THEN CLS : LOCATE 4, 4: PRINT "Scanning [ ]": char = 0
1060 IF CURRECORD < 1 THEN CURRECORD = 1: backflag = 0
     IF CURRECORD > lastrecord THEN
      GOTO 1080
      ELSE
      X = CURRECORD
      GET 1, X, UBPC3000RDRec(1)
      IF scancode = 1 THEN
       scanrec = CURRECORD
       GOSUB scanning
       IF trreadcode$ = "Y" THEN CURRECORD = CURRECORD + 1: GOTO 1060
      END IF
     END IF
     IF UBPC3000RDRec(1).Book <> Book THEN 1072
1070 scancode = 0
     GOSUB 1100: GOSUB 1300
     IF (cf = 1) AND (CURRECORD <> lastrecord) THEN GOTO 1060
1072 IF backflag = 1 THEN CURRECORD = CURRECORD - 1: GOTO 1060
     CURRECORD = CURRECORD + 1: GOTO 1060
1080 CLOSE : CLS
     PRINT CHR$(7);
     PRINT "NO MORE METERS!!!"
     PRINT CHR$(7); CHR$(7); CHR$(7)
     END

1100 CLS : IF UBPC3000RDRec(1).ReadFlag = "Y" THEN PRINT "*";  ELSE PRINT " ";
1110 PRINT UBPC3000RDRec(1).CustName
1115 PRINT " "; UBPC3000RDRec(1).ServAddress
1120 PRINT " Mtr# "; UBPC3000RDRec(1).MeterID;
     IF UBPC3000RDRec(1).MeterType = "E" THEN PRINT " *Electric"
     IF UBPC3000RDRec(1).MeterType = "D" THEN PRINT " *Demand"
     IF UBPC3000RDRec(1).MeterType = "W" THEN PRINT " *Water"
     IF UBPC3000RDRec(1).MeterType = "C" THEN PRINT " *Wtr/Swr"
     IF UBPC3000RDRec(1).MeterType = "S" THEN PRINT " *Sewer"
     IF UBPC3000RDRec(1).MeterType = "G" THEN PRINT " *Gas"
     PRINT " Previous.";
     IF UBPC3000RDRec(1).ReadFlag = "Y" THEN
       PRINT USING "########"; UBPC3000RDRec(1).PastRead
     ELSE
       PRINT USING "########"; UBPC3000RDRec(1).CurRead
     END IF
     IF UBPC3000RDRec(1).ReadFlag = "Y" THEN
      LOCATE 8, 1
      PRINT " Last Rdg."; USING "########"; UBPC3000RDRec(1).CurRead
     END IF
     IF ASC(LEFT$(UBPC3000RDRec(1).Note1, 1)) > 32 OR ASC(LEFT$(UBPC3000RDRec(1).Note2, 1)) > 32 OR ASC(LEFT$(UBPC3000RDRec(1).Note3, 1)) > 32 THEN PRINT CHR$(7);
1130 LOCATE 5, 1: PRINT " Current..";
     FL = -8: GOSUB 62000: IF cf = 1 THEN RETURN
     IF cf = 21 THEN
       backflag = 0
       CURRECORD = CURRECORD + 1: IF CURRECORD > lastrecord THEN CURRECORD = lastrecord
       GOTO 1060
     END IF
     IF cf = 20 THEN
       backflag = 1
       CURRECORD = CURRECORD - 1: IF CURRECORD < 1 THEN CURRECORD = 1
       GOTO 1060
     END IF
     IF cf = 31 THEN GOSUB displaypage: GOTO 1100
     IF cf = 30 THEN
       GOSUB findmeter
       IF scancode = 1 THEN CURRECORD = 1: GOTO 1059
       GOTO 1100
     END IF
1131 LOCATE 5, 11: CurRead# = VAL(in$): PRINT USING "########"; CurRead#
1140 LOCATE 6, 10: PRINT "---------"
     IF UBPC3000RDRec(1).ReadFlag = "Y" THEN
         IF CurRead# < UBPC3000RDRec(1).CurRead THEN
         GOSUB 1750
         ELSE
         calcusage# = CurRead# - UBPC3000RDRec(1).PastRead
         END IF
       ELSE
         IF CurRead# < UBPC3000RDRec(1).CurRead THEN
         GOSUB 1755
         ELSE
         calcusage# = CurRead# - UBPC3000RDRec(1).CurRead
         END IF
       END IF
     
1150 PRINT " Usage:   "; USING "########"; calcusage#

     IF calcusage# < UBPC3000RDRec(1).LowRead THEN
       LOCATE 7, 20: PRINT "L"
       PRINT CHR$(7) + CHR$(7)
     END IF
     IF calcusage# > UBPC3000RDRec(1).HighRead THEN
       LOCATE 7, 20: PRINT "H"
       PRINT CHR$(7) + CHR$(7)
     END IF



1152 LOCATE 8, 1: PRINT " ESC-NO : Enter-OK."; : FL = 1: GOSUB 62000: IF cf = 1 THEN GOSUB 1300: GOTO 1100
1170 IF UBPC3000RDRec(1).ReadFlag <> "Y" THEN UBPC3000RDRec(1).PastRead = UBPC3000RDRec(1).CurRead
     UBPC3000RDRec(1).CurRead = CurRead#
     UBPC3000RDRec(1).ReadTime = LEFT$(TIME$, 5)
     UBPC3000RDRec(1).ReadDate = Date2Num%(DATE$)
     UBPC3000RDRec(1).ReadFlag = "Y"
     PUT 1, CURRECORD, UBPC3000RDRec(1)
     backflag = 0
     RETURN

1300 LOCATE 5, 1: PRINT STRING$(62, 32): LOCATE 7, 26: PRINT STRING$(15, 32); : LOCATE 8, 26: PRINT STRING$(15, 32): LOCATE 10, 26: PRINT STRING$(15, 32): LOCATE 12, 1: PRINT STRING$(45, 32)
1310 RETURN


1750 REM check for meter triping
 J = LEN(STR$(UBPC3000RDRec(1).PastRead)) - 1
 MAX = 10 ^ J
 calcusage# = MAX - UBPC3000RDRec(1).PastRead + CurRead#
 RETURN

1755 REM check for meter triping
 J = LEN(STR$(UBPC3000RDRec(1).CurRead)) - 1
 MAX = 10 ^ J
 calcusage# = MAX - UBPC3000RDRec(1).CurRead + CurRead#
 RETURN

findmeter:
CLS
upper = 1
PRINT "SCAN  for un-read"
PRINT "or FIRST for first"
PRINT "or LAST  for last"
PRINT "or (ESC-RETURN) "
PRINT "==>"; : FL = 15: GOSUB 62000
upper = 0
in$ = UCASE$(in$)
IF (VAL(in$) = 0) AND (in$ = "SCAN") THEN scancode = 1: RETURN
IF (VAL(in$) = 0) AND (in$ = "FIRST") THEN CURRECORD = 1: GET 1, CURRECORD: RETURN
IF (VAL(in$) = 0) AND (in$ = "LAST") THEN CURRECORD = lastrecord: GET 1, CURRECORD: RETURN
scancode = 0
IF cf = 1 THEN RETURN
CLS : LOCATE 4, 4: PRINT "Scanning [ ]": char = 0
FOR X = 1 TO LOF(1) / UBPC3000RdRecLen
   GET 1, X, UBPC3000RDRec(1)
  scanrec = X
  GOSUB scanning
  IF (LTRIM$(RTRIM$(in$)) = LTRIM$(RTRIM$(trmeter$))) THEN CURRECORD = X: RETURN
NEXT X
GET 1, CURRECORD
LOCATE 8, 1: PRINT " METER NOT LOCATED"
PRINT " Press <ENTER>"; : FL = 1: GOSUB 62000: RETURN
    

displaypage:
IF cf = 31 THEN
 CLS
 LOCATE 1, 1: PRINT STRING$(20, CHR$(42)): LOCATE 1, 8: PRINT "Notes";
 LOCATE 2, 1: PRINT STRING$(20, 42)
 LOCATE 3, 1
 PRINT UBPC3000RDRec(1).Note1
 PRINT UBPC3000RDRec(1).Note2
 PRINT UBPC3000RDRec(1).Note3
 PRINT STRING$(20, 42)
dp1: LOCATE 8, 1: PRINT "New Notes (Y/N) [ ]"; : LOCATE 8, 18: FL = 1: GOSUB 62210
 IF cf = 30 THEN RETURN
 IF NOT (in$ = "Y" OR in$ = "N") THEN GOTO dp1
 IF in$ = "N" THEN RETURN
 IF in$ = "Y" THEN FOR i = 3 TO 5: LOCATE i, 1: PRINT STRING$(20, 32): NEXT i
dp2: LOCATE 8, 1: PRINT STRING$(20, 32)
  LOCATE 3, 1: FL = 20: GOSUB 62000: n1$ = in$: IF cf = 1 THEN RETURN
  LOCATE 4, 1: FL = 20: GOSUB 62000: n2$ = in$: IF cf = 1 THEN RETURN
  LOCATE 5, 1: FL = 20: GOSUB 62000: n3$ = in$: IF cf = 1 THEN RETURN
dp3: LOCATE 8, 1: PRINT "Save Notes (Y/N) [ ]"; : LOCATE 8, 19: FL = 1: GOSUB 62210
  IF cf = 1 THEN noteflag = 0: RETURN
  IF in$ = "Y" THEN
dp4: LOCATE 8, 1: PRINT "Perm/Temp (P/T) [ ] "; : LOCATE 8, 18: FL = 1: GOSUB 62210
  IF cf = 1 THEN noteflag = 0: RETURN
  in$ = UCASE$(in$)
  IF (in$ <> "P") AND (in$ <> "T") THEN GOTO dp4
   noteflag = 1
   UBPC3000RDRec(1).Note1 = n1$
   UBPC3000RDRec(1).Note2 = n2$
   UBPC3000RDRec(1).Note3 = n3$
   UBPC3000RDRec(1).NoteStatus = in$
   PUT 1, CURRECORD, UBPC3000RDRec(1)
   RETURN
  END IF

  IF in$ = "N" THEN GOTO dp2
  GOTO dp3
END IF
RETURN: REM must go back to first page

scanning:
  LOCATE 4, 14
  IF (scanrec MOD 10 = 0) THEN char = (char MOD 4) + 1
  IF char = 1 THEN PRINT "\"
  IF char = 2 THEN PRINT "|"
  IF char = 3 THEN PRINT "/"
  IF char = 4 THEN PRINT "-"
RETURN


60000 UBPC3000RdRecLen = LEN(UBPC3000RDRec(1))
      UBPC3000RdFile = FREEFILE
      OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS #1 LEN = UBPC3000RdRecLen
      RETURN

62000 cf = 0: in$ = "": RO = CSRLIN: PO = POS(X): ZD = 0: ZS = 0: ZL = 0: IF FL = ZD THEN FL = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(FL), 95); : LOCATE RO, PO
62030 z$ = INKEY$: IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(FL) - ZL, 32);
62050 GOSUB 62200: RETURN
62060 IF z$ = CHR$(27) OR z$ = CHR$(21) THEN cf = 1: RETURN
      IF z$ = CHR$(9) THEN cf = 2: RETURN
      IF LEN(z$) <> 2 THEN 62065
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 59 THEN cf = 11: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 60 THEN cf = 12: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 61 THEN cf = 13: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 72 THEN cf = 20: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 80 THEN cf = 21: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 75 THEN cf = 30: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 76 THEN cf = 30: RETURN
      IF ASC(LEFT$(z$, 1)) = 0 AND ASC(MID$(z$, 2, 1)) = 77 THEN cf = 31: RETURN
62065 IF ASC(z$) = 0 THEN 62030
62070 IF LEN(z$) > 1 THEN cf = ASC(RIGHT$(z$, 1)) - 57: IF cf > 0 THEN RETURN ELSE cf = cf + 27: RETURN
62080 IF z$ = " " THEN 62130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 62000
62090 IF z$ <> CHR$(8) THEN 62130 ELSE IF ZL = 0 THEN 62030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF FL > 0 THEN 62110 ELSE IF z$ = "," THEN 62120
62100 IF z$ = "." THEN ZD = 0: GOTO 62110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
62110 in$ = LEFT$(in$, LEN(in$) - 1)
62120 ZL = ZL - 1: GOTO 62030
62130 IF ABS(FL) = ZL THEN 62030 ELSE IF FL > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 62180
62140 IF z$ = "." AND ZD = 0 THEN ZD = 1: GOTO 62180
62150 IF z$ = "," THEN PRINT ","; : ZL = ZL + 1: GOTO 62190
62160 IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170 IF z$ < "0" OR z$ > "9" THEN 62030
62180 IF upper = 1 THEN z$ = UCASE$(z$)
      PRINT z$; : in$ = in$ + z$: ZL = ZL + 1
62190 IF ABS(FL) = 1 THEN 62050 ELSE 62030
62200 LOCATE RO, PO: PRINT in$: RETURN
62210 GOSUB 62000: IF LEN(in$) THEN in$ = CHR$(ASC(in$) AND 223)
62220 RETURN
63000 OPEN "R", 3, "wbtown.dat": FIELD 3, 40 AS TOWNNAME$: GET 3, 1: TOWN$ = TOWNNAME$: CLOSE 3: RETURN
63500 LOCATE 1, 1: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN



