DECLARE SUB BackLightToggle ()
DECLARE SUB WInput (Edit$, GoodKey$, Row%, Col%, ExitCode%)
DECLARE SUB SetFont ()
' New hand held program for Version 8.2w
DEFINT A-Z
DECLARE SUB CLocate (BYVAL Row%, BYVAL Col%)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB QPSound (Frequency%, Duration%)
DECLARE FUNCTION GetPower ()
DECLARE FUNCTION BiosKey% ()
  '$INCLUDE: 'ubsetup.bi'

CONST False = 0, True = NOT False

GoodKey1$ = " abcdefghijklmnopqrstuvwxyz"
GoodKey2$ = UCASE$(GoodKey1$)
GoodKey3$ = "1234567890"
  
  STACK 10000

  CALL SetFont
  
  Ast$ = STRING$(20, "*")
  Blank$ = STRING$(20, " ")
  Fmt1$ = "########"
  REDIM UBPC3000RDRec(1) AS UBPC3000ReadRecType
  UBPC3000RdRecLen = LEN(UBPC3000RDRec(1))

top:
  CLS
  PRINT "ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿"          '1
  PRINT "³       HUSKY HandHeld       ³"          '2
  PRINT "³    Meter Reading System    ³"          '3
  PRINT "³ (C)1997 Southern Software  ³"          '4
  PRINT "³                            ³"          '5
  PRINT "³ Enter Route#               ³"          '6
  PRINT "³                            ³"          '7
  PRINT "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ";         '8

  fl = 2
  Route$ = SPACE$(ABS(fl))
  WInput Route$, GoodKey3$, 6, 10, ExitCode

' LOCATE 6, 16
'  fl = -2
'  GOSUB 62000
'  'IF cf = 1 THEN END
'
 ' IF Cf = 1 AND LEN(In$) = 0 THEN
 '   CLS
'    PRINT
'    PRINT "Program Exited."
'    END
'  ELSEIF Cf > 1 THEN
'    GOTO top
'  ELSEIF Cf = 0 AND LEN(In$) = 0 THEN
'
'  END IF

  CurBook = VAL(In$)
  CLS
  PRINT
  PRINT "Searching for Route: "; CurBook
  CLocate 4, 4
  PRINT "Scanning [ ]"
  GOSUB Scanning
  GOSUB OpenFile
  LastRecord = LOF(1) / UBPC3000RdRecLen
  GOSUB CloseFile
  CurRecord = 1
  
1059
  IF ScanCode = 1 OR BackFlag = 1 THEN
    CLS
    CLocate 4, 4
    PRINT "Scanning [ ]"
  END IF

1060

  IF CurRecord < 1 THEN
    CurRecord = 1
    BackFlag = 0
  END IF

  IF CurRecord > LastRecord THEN
    BackFlag = 0
    GOTO 1080
  ELSE

    X = CurRecord
    GOSUB OpenFile
    GET 1, X, UBPC3000RDRec(1)
      
    IF ScanCode = 1 OR BackFlag = 1 THEN
      GOSUB Scanning
    END IF
    GOSUB CloseFile
    IF ScanCode = 1 THEN
      IF TrReadCode$ = "Y" THEN
        CurRecord = CurRecord + 1
        GOTO 1060
      END IF
    END IF
  END IF
  IF UBPC3000RDRec(1).Book <> CurBook THEN 1072
  ScanCode = 0
  BackFlag = 0
  GOSUB 1100
  IF (Cf = 1) AND (CurRecord <> LastRecord) THEN
    GOTO 1060
  END IF
  
1072
  IF BackFlag = 1 THEN
    CurRecord = CurRecord - 1
    'ScanCode = 1
  ELSE
    CurRecord = CurRecord + 1
  END IF
  GOTO 1060
  
1080
  'CLOSE
  IF BeenHere THEN
    CLS
    QPSound 1000, 1
    PRINT "NO MORE METERS!!!"
    QPSound 1000, 3
    END
  ELSE
    CLS
    'QPSound 1000, 1
    PRINT "ROUTE: "; CurBook
    PRINT
    PRINT "NOT FOUND"
    PRINT
    PRINT
    PRINT "Press any Key. . .";
    'QPSound 1000, 3
    fl = 1
    GOSUB 62000
    GOTO top
  END IF

  
1100
  CLS
  BeenHere = -1
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
    PRINT USING Fmt1$; UBPC3000RDRec(1).PastRead
  ELSE
    PRINT USING Fmt1$; UBPC3000RDRec(1).CurRead
  END IF
  IF UBPC3000RDRec(1).ReadFlag = "Y" THEN
    CLocate 8, 1
    PRINT " Last Rdg."; USING Fmt1$; UBPC3000RDRec(1).CurRead
  END IF
  IF ASC(LEFT$(UBPC3000RDRec(1).Note1, 1)) > 32 OR ASC(LEFT$(UBPC3000RDRec(1).Note2, 1)) > 32 OR ASC(LEFT$(UBPC3000RDRec(1).Note3, 1)) > 32 THEN
    QPSound 1000, 1
  END IF
  CLocate 5, 1
  PRINT " Current..";
  fl = -8

GetInput:

  GOSUB 62000
  
  IF Cf = 1 THEN RETURN
  
  IF RestartFlag THEN
    CLS
    RUN
  END IF
  
  IF Cf = 21 THEN
    BackFlag = 0
    CurRecord = CurRecord + 1
    IF CurRecord > LastRecord THEN
      CurRecord = LastRecord
      GOTO 1080
    ELSE
      GOTO 1060
    END IF
  END IF

  IF Cf = 20 THEN
    CurRecord = CurRecord - 1
    IF CurRecord < 1 THEN CurRecord = 1
    IF BackFlag = 0 THEN
      BackFlag = 1
      GOTO 1059
    ELSE
      GOTO 1060
    END IF
  END IF

  IF Cf = 31 THEN
    GOSUB displaypage
    GOTO 1100
  END IF

  IF Cf = 30 THEN
    GOSUB findmeter
    IF ScanCode = 1 THEN
      CurRecord = 1
      GOTO 1059
    END IF
    GOTO 1100
  END IF

  SELECT CASE Cf
  CASE 11, 12, 13
    GOTO GetInput
  END SELECT
  
  CLocate 5, 11
  CurRead# = VAL(In$)
  PRINT USING Fmt1$; CurRead#
  CLocate 6, 10
  PRINT "---------"
  IF UBPC3000RDRec(1).ReadFlag = "Y" THEN
    IF CurRead# < UBPC3000RDRec(1).CurRead THEN
      GOSUB ChkMeterTrip
    ELSE
      CalcUsage# = CurRead# - UBPC3000RDRec(1).PastRead
    END IF
  ELSE
    IF CurRead# < UBPC3000RDRec(1).CurRead THEN
      GOSUB 1755
    ELSE
      CalcUsage# = CurRead# - UBPC3000RDRec(1).CurRead
    END IF
  END IF
  
1150
  PRINT " Usage:   "; USING Fmt1$; CalcUsage#
  IF CalcUsage# < UBPC3000RDRec(1).LowRead THEN
    CLocate 7, 20
    PRINT "L"
    QPSound 1000, 2
  END IF
  
  IF CalcUsage# > UBPC3000RDRec(1).HighRead THEN
    CLocate 7, 20
    PRINT "H"
    QPSound 1000, 2
  END IF
  
1152
  CLocate 8, 1
  PRINT " ESC-NO : Enter-OK.";
  fl = 1
  GOSUB 62000
  IF Cf = 1 THEN
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
  PUT 1, CurRecord, UBPC3000RDRec(1)
  GOSUB CloseFile
  BackFlag = 0
RETURN
  
ChkMeterTrip:
  'check for meter triping
  J = LEN(STR$(UBPC3000RDRec(1).PastRead)) - 1
  Max# = 10 ^ J
  CalcUsage# = Max# - UBPC3000RDRec(1).PastRead + CurRead#
  RETURN
  
1755          REM check for meter triping
  J = LEN(STR$(UBPC3000RDRec(1).CurRead)) - 1
  Max# = 10 ^ J
  CalcUsage# = Max# - UBPC3000RDRec(1).CurRead + CurRead#
  RETURN
  
findmeter:
  CLS
  PRINT "Battery Power: "; GetPower + 1; "%"
  PRINT "Enter Meter Number"
  PRINT "SCAN  for un-read"
  PRINT "FIRST for first"
  PRINT "LAST  for last"
  PRINT "or (ESC-RETURN) "
  PRINT "==>";
  fl = 15
  GOSUB 62000
  
  IF Cf = 1 OR LEN(In$) = 0 THEN
    RestartFlag = 0
    RETURN
  END IF
  
  In$ = UCASE$(In$)
  
  IF (VAL(In$) = 0) AND (In$ = "SCAN") THEN
    ScanCode = 1
    RETURN
  END IF

  IF (VAL(In$) = 0) AND (In$ = "FIRST") THEN
    CurRecord = 1
    GOSUB OpenFile
    TNumOfRecs& = LOF(1) \ UBPC3000RdRecLen
    FOR CurRecord = 1 TO TNumOfRecs&
      GET 1, CurRecord, UBPC3000RDRec(1)
      IF UBPC3000RDRec(1).Book = CurBook THEN
        IF UBPC3000RDRec(1).ReadFlag <> "Y" THEN
          EXIT FOR
        END IF
      END IF
    NEXT
    GOSUB CloseFile
    RETURN
  END IF

  IF (VAL(In$) = 0) AND (In$ = "LAST") THEN
    'CurRecord = LastRecord
    Last = CurRecord
    GOSUB OpenFile
    TNumOfRecs& = LOF(1) \ UBPC3000RdRecLen
    FOR CurRecord = 1 TO TNumOfRecs&
      GET 1, CurRecord, UBPC3000RDRec(1)
      IF UBPC3000RDRec(1).Book = CurBook THEN
        IF UBPC3000RDRec(1).ReadFlag <> "Y" THEN
          Last = CurRecord
        END IF
      END IF
    NEXT
    CurRecord = Last
    GET 1, CurRecord, UBPC3000RDRec(1)
    GOSUB CloseFile
    RETURN
  END IF
  ScanCode = 0
  
  IF Cf = 1 THEN RETURN
  CLS
  CLocate 4, 4
  WhatMeter$ = QPTrim$(In$)
  PRINT "Scanning [ ]"
  GOSUB OpenFile
  FOR X = 1 TO LastRecord
    GET 1, X, UBPC3000RDRec(1)
    GOSUB Scanning
    trmeter$ = QPTrim$(UBPC3000RDRec(1).MeterID)
    IF WhatMeter$ = trmeter$ THEN
      CurRecord = X
      GOSUB CloseFile
      RETURN
    END IF
  NEXT X
  GET 1, CurRecord, UBPC3000RDRec(1)
  GOSUB CloseFile
  CLocate 7, 1
  PRINT " METER: "; WhatMeter$
  PRINT " NOT Located"
  PRINT " Press <ENTER>";
  fl = 1
  GOSUB 62000
  RETURN
  
displaypage:
  IF Cf = 31 THEN
    CLS
    CLocate 1, 1
    PRINT Ast$
    CLocate 1, 8
    PRINT "Notes"
    PRINT Ast$
    PRINT UBPC3000RDRec(1).Note1
    PRINT UBPC3000RDRec(1).Note2
    PRINT UBPC3000RDRec(1).Note3
    PRINT Ast$
    
dp1:
    CLocate 8, 1
    PRINT "New Notes (Y/N) [ ]";
    CLocate 8, 18
    fl = 1
    GOSUB 62210
    IF Cf = 30 THEN RETURN
    IF NOT (In$ = "Y" OR In$ = "N") THEN GOTO dp1
    
    IF In$ = "N" THEN RETURN
    IF In$ = "Y" THEN
      FOR i = 3 TO 5
        CLocate i, 1
        PRINT Blank$
      NEXT i
    END IF
dp2:
    CLocate 8, 1
    PRINT Blank$
    
    CLocate 3, 1
    fl = 20
    GOSUB 62000
    n1$ = In$
    IF Cf = 1 THEN RETURN
    
    CLocate 4, 1
    fl = 20
    GOSUB 62000
    n2$ = In$
    IF Cf = 1 THEN RETURN
    
    CLocate 5, 1
    fl = 20
    GOSUB 62000
    n3$ = In$
    IF Cf = 1 THEN RETURN
    
dp3:
    CLocate 8, 1
    PRINT "Save Notes (Y/N) [ ]";
    CLocate 8, 19
    fl = 1
    GOSUB 62210
    IF Cf = 1 THEN noteflag = 0: RETURN
    IF In$ = "Y" THEN
dp4:
      CLocate 8, 1
      PRINT "Perm/Temp (P/T) [ ] ";
      CLocate 8, 18
      fl = 1
      GOSUB 62210
      IF Cf = 1 THEN noteflag = 0: RETURN
      In$ = UCASE$(In$)
      IF (In$ <> "P") AND (In$ <> "T") THEN GOTO dp4
      noteflag = 1
      UBPC3000RDRec(1).Note1 = n1$
      UBPC3000RDRec(1).Note2 = n2$
      UBPC3000RDRec(1).Note3 = n3$
      UBPC3000RDRec(1).NoteStatus = In$
      GOSUB OpenFile
      PUT 1, CurRecord, UBPC3000RDRec(1)
      GOSUB CloseFile
      RETURN
    END IF
    IF In$ = "N" THEN GOTO dp2
    GOTO dp3
  END IF
  RETURN
  'must go back to first page
  
Scanning:
  CLocate 4, 14
  Char = Char + 1
  IF Char > 4 THEN Char = 1
  PRINT MID$("\|/-", Char, 1);
RETURN
  
  
OpenFile:
  OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS #1 LEN = UBPC3000RdRecLen
RETURN
  
CloseFile:
  CLOSE #1
RETURN
  
62000
  Cf = 0
  In$ = ""
  RO = CSRLIN
  PO = POS(X)
  ZD = 0
  ZS = 0
  ZL = 0
  IF fl = ZD THEN fl = 1
62010
  IF INKEY$ <> "" THEN 62010
62020
  PRINT STRING$(ABS(fl), 95);
  CLocate RO, PO
62030
  z$ = INKEY$
  IF LEN(z$) = 0 THEN 62030

62040
  IF z$ <> CHR$(13) THEN 62060 ELSE PRINT STRING$(ABS(fl) - ZL, 32);
62050
  GOSUB 62200
  RETURN
62060
  IF z$ = CHR$(27) OR z$ = CHR$(21) THEN Cf = 1: RETURN
  IF z$ = CHR$(9) THEN Cf = 2: RETURN
  IF LEN(z$) <> 2 THEN 62065
  
  IF LEN(z$) = 2 THEN

    SELECT CASE ASC(MID$(z$, 2, 1))
    'CASE 104
    '  RestartFlag = -1
    CASE 59
      Cf = 11
    CASE 60
      Cf = 12
      CALL BackLightToggle
    CASE 61
      Cf = 13
    CASE 62  'f4
      RestartFlag = -1
    CASE 72
      Cf = 20
    CASE 80
      Cf = 21
    CASE 75
      Cf = 30
    CASE 76
      Cf = 30
    CASE 77
      Cf = 31
    END SELECT

    IF Cf > 0 OR RestartFlag THEN
      RETURN
    END IF
  END IF
  
62065
  IF ASC(z$) = 0 THEN 62030
62070
  IF LEN(z$) > 1 THEN
    Cf = ASC(RIGHT$(z$, 1)) - 57:
    IF Cf > 0 THEN
      RETURN
    ELSE
      Cf = Cf + 27
      RETURN
    END IF
  END IF
  
62080
  IF z$ = " " THEN 62130 ELSE IF z$ = CHR$(29) THEN PRINT STRING$(ZL, 29); : GOTO 62000
62090
  IF z$ <> CHR$(8) THEN 62130 ELSE IF ZL = 0 THEN 62030 ELSE PRINT CHR$(29); : PRINT "_"; : PRINT CHR$(29); : IF fl > 0 THEN 62110 ELSE IF z$ = "," THEN 62120
62100
  IF z$ = "." THEN ZD = 0: GOTO 62110 ELSE IF z$ = "+" OR z$ = "-" THEN ZS = 0
62110
  In$ = LEFT$(In$, LEN(In$) - 1)
62120
  ZL = ZL - 1: GOTO 62030
62130
  IF ABS(fl) = ZL THEN 62030 ELSE IF fl > 0 THEN IF z$ >= " " AND z$ <= "z" THEN 62180
62140
  IF z$ = "." AND ZD = 0 THEN ZD = 1: GOTO 62180
62150
  IF z$ = "," THEN
    PRINT ",";
    ZL = ZL + 1
    GOTO 62190
  END IF
  
62160
  IF (z$ = "-" OR z$ = "+") AND ZS = 0 AND ZL = 0 THEN ZS = 1: GOTO 62180
62170
  IF z$ < "0" OR z$ > "9" THEN 62030
62180
  z$ = UCASE$(z$)
  PRINT z$;
  In$ = In$ + z$
  ZL = ZL + 1
62190
  IF ABS(fl) = 1 THEN 62050 ELSE 62030
62200
  CLocate RO, PO
  PRINT In$
  RETURN
  
62210 GOSUB 62000
  IF LEN(In$) THEN In$ = CHR$(ASC(In$) AND 223)
62220 RETURN

SUB CLocate (BYVAL Row%, BYVAL Col%)
  LOCATE Row, Col
END SUB

