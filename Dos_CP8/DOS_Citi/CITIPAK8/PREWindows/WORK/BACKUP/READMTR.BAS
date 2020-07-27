' New hand held program for Version 8.2
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION QPTrim$ (Text$)

'$INCLUDE: 'ubsetup.bi'

  STACK 10000

  REDIM UBPC3000RDRec(1) AS UBPC3000ReadRecType
  UBPC3000RdRecLen = LEN(UBPC3000RDRec(1))
'  SCREEN 1
'  WIDTH 80
  CLS
  PRINT "PC 3000 H/H Meter"
  PRINT " Reading System"
  PRINT
  PRINT "(C)1997 Southern"
  PRINT "    Software "
  PRINT
  PRINT "Enter Route # ";
  FL = -2
  GOSUB 62000
  IF cf = 1 THEN END
  Book = VAL(in$)
  CLS
  GOSUB OpenFile
  lastrecord = LOF(1) / UBPC3000RdRecLen
  GOSUB CloseFile
  CURRECORD = 1
  
1059
  IF scancode = 1 THEN
    CLS
    LOCATE 4, 4
    PRINT "Scanning [ ]"
  END IF
1060
  IF CURRECORD < 1 THEN
    CURRECORD = 1
    backflag = 0
  END IF
  IF CURRECORD > lastrecord THEN
    GOTO 1080
  ELSE
    X = CURRECORD
    GOSUB OpenFile
    GET 1, X, UBPC3000RDRec(1)
    GOSUB CloseFile
    IF scancode = 1 THEN
      'scanrec = CURRECORD
      GOSUB scanning
      IF trreadcode$ = "Y" THEN
        CURRECORD = CURRECORD + 1
        GOTO 1060
      END IF
    END IF
  END IF
  IF UBPC3000RDRec(1).Book <> Book THEN 1072
  scancode = 0
  GOSUB 1100
  GOSUB 1300
  IF (cf = 1) AND (CURRECORD <> lastrecord) THEN GOTO 1060

1072
  IF backflag = 1 THEN
    CURRECORD = CURRECORD - 1
  ELSE
    CURRECORD = CURRECORD + 1
  END IF
  GOTO 1060

1080
  CLOSE
  CLS
  PRINT CHR$(7);
  PRINT "NO MORE METERS!!!"
  PRINT CHR$(7); CHR$(7); CHR$(7)
  END
  
1100
  CLS
  IF UBPC3000RDRec(1).ReadFlag = "Y" THEN
    PRINT "*";
  ELSE
    PRINT " ";
  END IF
  PRINT UBPC3000RDRec(1).CustName
  PRINT " "; UBPC3000RDRec(1).ServAddress
  PRINT " Mtr# "; UBPC3000RDRec(1).MeterID;
  SELECT CASE UBPC3000RDRec(1).MeterType
  CASE "E"
    PRINT " *Electric"
  CASE "D"
    PRINT " *Demand"
  CASE "W"
    PRINT " *Water"
  CASE "C"
    PRINT " *Wtr/Swr"
  CASE "S"
    PRINT " *Sewer"
  CASE "G"
    PRINT " *Gas"
  END SELECT
  PRINT " Previous.";

  IF UBPC3000RDRec(1).ReadFlag = "Y" THEN
    PRINT USING "##########"; UBPC3000RDRec(1).PastRead
  ELSE
    PRINT USING "##########"; UBPC3000RDRec(1).CurRead
  END IF
  IF UBPC3000RDRec(1).ReadFlag = "Y" THEN
    LOCATE 8, 1
    PRINT " Last Rdg."; USING "##########"; UBPC3000RDRec(1).CurRead
  END IF
  IF ASC(LEFT$(UBPC3000RDRec(1).Note1, 1)) > 32 OR ASC(LEFT$(UBPC3000RDRec(1).Note2, 1)) > 32 OR ASC(LEFT$(UBPC3000RDRec(1).Note3, 1)) > 32 THEN PRINT CHR$(7);
  LOCATE 5, 1
  PRINT " Current..";
  FL = -10
  GOSUB 62000

  IF cf = 1 THEN RETURN

  IF cf = 21 THEN
    backflag = 0
    CURRECORD = CURRECORD + 1
    IF CURRECORD > lastrecord THEN CURRECORD = lastrecord
    GOTO 1060
  END IF
  IF cf = 20 THEN
    backflag = 1
    CURRECORD = CURRECORD - 1
    IF CURRECORD < 1 THEN CURRECORD = 1
    GOTO 1060
  END IF
  IF cf = 31 THEN
    GOSUB displaypage
    GOTO 1100
  END IF
  IF cf = 30 THEN
    GOSUB findmeter
    IF scancode = 1 THEN
      CURRECORD = 1
      GOTO 1059
    END IF
    GOTO 1100
  END IF

  LOCATE 5, 11
  CurRead# = VAL(in$)
  PRINT USING "##########"; CurRead#
  LOCATE 6, 10
  PRINT "---------"
  IF UBPC3000RDRec(1).ReadFlag = "Y" THEN
    IF CurRead# < UBPC3000RDRec(1).CurRead THEN
      GOSUB ChkMeterTrip
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
  
1150
  PRINT " Usage:   "; USING "##########"; calcusage#
  IF calcusage# < UBPC3000RDRec(1).LowRead THEN
    LOCATE 7, 20
    PRINT "L"
    PRINT CHR$(7) + CHR$(7)
  END IF

  IF calcusage# > UBPC3000RDRec(1).HighRead THEN
    LOCATE 7, 20
    PRINT "H"
    PRINT CHR$(7) + CHR$(7)
  END IF

1152
  LOCATE 8, 1
  PRINT " ESC-NO : Enter-OK.";
  FL = 1
  GOSUB 62000
  IF cf = 1 THEN
    GOSUB 1300
    GOTO 1100
  END IF

  IF UBPC3000RDRec(1).ReadFlag <> "Y" THEN
    UBPC3000RDRec(1).PastRead = UBPC3000RDRec(1).CurRead
  END IF
  UBPC3000RDRec(1).CurRead = CurRead#
  UBPC3000RDRec(1).ReadTime = LEFT$(TIME$, 5)
  UBPC3000RDRec(1).ReadDate = Date2Num%(DATE$)
  UBPC3000RDRec(1).ReadFlag = "Y"
  GOSUB OpenFile
  PUT 1, CURRECORD, UBPC3000RDRec(1)
  GOSUB CloseFile
  backflag = 0
RETURN
  
1300
     LOCATE 5, 1
     PRINT STRING$(62, 32)
     LOCATE 7, 26
     PRINT STRING$(15, 32)
     LOCATE 8, 26
     PRINT STRING$(15, 32)
     LOCATE 10, 26
     PRINT STRING$(15, 32)
     LOCATE 12, 1
     PRINT STRING$(45, 32)
RETURN
  
  
ChkMeterTrip:
  'check for meter triping
  J = LEN(STR$(UBPC3000RDRec(1).PastRead)) - 1
  MAX = 10 ^ J
  calcusage# = MAX - UBPC3000RDRec(1).PastRead + CurRead#
RETURN
  
1755          REM check for meter triping
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
  PRINT "==>";
  FL = 15
  GOSUB 62000
  upper = 0
  in$ = UCASE$(in$)

  IF (VAL(in$) = 0) AND (in$ = "SCAN") THEN
    scancode = 1
    RETURN
  END IF
  IF (VAL(in$) = 0) AND (in$ = "FIRST") THEN
    CURRECORD = 1
    GOSUB OpenFile
    GET 1, CURRECORD, UBPC3000RDRec(1)
    GOSUB CloseFile
    RETURN
  END IF
  IF (VAL(in$) = 0) AND (in$ = "LAST") THEN
    CURRECORD = lastrecord
    GOSUB OpenFile
    GET 1, CURRECORD, UBPC3000RDRec(1)
    GOSUB CloseFile
    RETURN
  END IF
  scancode = 0

  IF cf = 1 THEN RETURN
  CLS
  LOCATE 4, 4
  WhatMeter$ = LTRIM$(RTRIM$(in$))
  PRINT "Scanning [ ]"
  'Char = 0
  GOSUB OpenFile
  FOR X = 1 TO LOF(1) / UBPC3000RdRecLen
    GET 1, X, UBPC3000RDRec(1)
    GOSUB scanning
    trmeter$ = QPTrim$(UBPC3000RDRec(1).MeterID)
    IF WhatMeter$ = trmeter$ THEN
      CURRECORD = X
      GOSUB CloseFile
      RETURN
    END IF
  NEXT X
  GET 1, CURRECORD, UBPC3000RDRec(1)
  GOSUB CloseFile
  LOCATE 7, 1
  PRINT " METER: "; WhatMeter$
  PRINT " NOT LOCATED"
  PRINT " Press <ENTER>";
  FL = 1
  GOSUB 62000
RETURN

displaypage:
  IF cf = 31 THEN
    CLS
    LOCATE 1, 1
    PRINT STRING$(20, CHR$(42))
    LOCATE 1, 8
    PRINT "Notes";
    LOCATE 2, 1
    PRINT STRING$(20, 42)
    LOCATE 3, 1
    PRINT UBPC3000RDRec(1).Note1
    PRINT UBPC3000RDRec(1).Note2
    PRINT UBPC3000RDRec(1).Note3
    PRINT STRING$(20, 42)
dp1: LOCATE 8, 1
    PRINT "New Notes (Y/N) [ ]";
    LOCATE 8, 18
    FL = 1
    GOSUB 62210
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
      GOSUB OpenFile
      PUT 1, CURRECORD, UBPC3000RDRec(1)
      GOSUB CloseFile
      RETURN
    END IF
    IF in$ = "N" THEN GOTO dp2
    GOTO dp3
  END IF
  RETURN
  'must go back to first page
  
scanning:
  LOCATE 4, 14
  Char = Char + 1
  IF Char > 4 THEN Char = 1
  PRINT MID$("\|/-", Char, 1);
  'IF Char = 4 THEN Char = 0
  RETURN
  
  
OpenFile:
  OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS #1 LEN = UBPC3000RdRecLen
RETURN

CloseFile:
  CLOSE #1
RETURN

62000 cf = 0: in$ = "": RO = CSRLIN: PO = POS(X): ZD = 0: ZS = 0: ZL = 0: IF FL = ZD THEN FL = 1
62010 IF INKEY$ <> "" THEN 62010
62020 PRINT STRING$(ABS(FL), 95); : LOCATE RO, PO
62030 z$ = INKEY$
      IF z$ = "" THEN 62030
62040 IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(FL) - ZL, 32);
62050 GOSUB 62200
      RETURN
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
' OPEN "R", 3, "wbtown.dat": FIELD 3, 40 AS TOWNNAME$: GET 3, 1: TOWN$ = TOWNNAME$: CLOSE 3: RETURN
'63500 LOCATE 1, 1: PRINT "Sys. Date: "; SYSDATE$; TAB(70); SCREN$: PRINT STRING$(79, "_"): RETURN
  
  
  

