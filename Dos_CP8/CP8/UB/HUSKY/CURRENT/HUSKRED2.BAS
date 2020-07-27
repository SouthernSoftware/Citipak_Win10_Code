DEFINT A-Z   ' New hand held program for Version 8.2w
DECLARE SUB SetFont ()
DECLARE SUB WInput (Edit$, GoodKey$, Row%, Col%, ExitCode%)
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
  
  CALL SetFont
  
  Ast$ = STRING$(20, "*")
  Blank$ = STRING$(20, " ")
  Fmt2$ = "        "
  REDIM UBPC3000RDRec(1) AS UBPC3000ReadRecType
  UBPC3000RdRecLen = LEN(UBPC3000RDRec(1))

Top:
  CLS
  PRINT "ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿"          '1
  PRINT "³       HUSKY HandHeld       ³"          '2
  PRINT "³ Meter Reading System v2.1  ³"          '3
  PRINT "³ (C)1998 Southern Software  ³"          '4
  PRINT "³                            ³"          '5
  PRINT "³ Enter Route#               ³"          '6
  PRINT "³                            ³"          '7
  PRINT "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ";         '8

  Route$ = "  "
  WInput Route$, GoodKey3$, 6, 16, ExitCode
  IF LEN(QPTrim$(Route$)) = 0 OR ExitCode = -27 THEN
    CLS
    CLocate 3, 1
    PRINT "Program Exited."
    END
  END IF

  CurBook = VAL(Route$)

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

  IF (CF = -27) AND (CurRecord <> LastRecord) THEN
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
  IF BeenHere THEN
    CLS
    QPSound 1000, 1
    PRINT "NO MORE METERS!!!"
    QPSound 1000, 3
    PRINT
    PRINT
    PRINT
    PRINT "Press any Key...";
    zz = BiosKey
    RUN
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
    zz = BiosKey
    GOTO Top
  END IF

  
1100
  CLS
  BeenHere = True
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
  PRINT " Previous:";
  IF UBPC3000RDRec(1).ReadFlag = "Y" THEN
    RSET Fmt2$ = QPTrim$(STR$(UBPC3000RDRec(1).PastRead))
    'PRINT USING Fmt1$; UBPC3000RDRec(1).PastRead
    PRINT Fmt2$;
  ELSE
    RSET Fmt2$ = QPTrim$(STR$(UBPC3000RDRec(1).CurRead))
    'PRINT USING Fmt1$; UBPC3000RDRec(1).CurRead
    PRINT Fmt2$;
  END IF
  IF UBPC3000RDRec(1).ReadFlag = "Y" THEN
    CLocate 8, 1
    RSET Fmt2$ = QPTrim$(STR$(UBPC3000RDRec(1).CurRead))
    PRINT " Last Rdg."; Fmt2$;
    'PRINT " Last Rdg."; USING Fmt1$; UBPC3000RDRec(1).CurRead
  END IF
  IF ASC(LEFT$(UBPC3000RDRec(1).Note1, 1)) > 32 OR ASC(LEFT$(UBPC3000RDRec(1).Note2, 1)) > 32 OR ASC(LEFT$(UBPC3000RDRec(1).Note3, 1)) > 32 THEN
    CLocate 6, 2
    PRINT "Note!";
    QPSound 1000, 1
  END IF
  CLocate 5, 1
  PRINT " Current..";

GetInput:

  CF = 0
  fl = 8

  in$ = SPACE$(fl)

  IF NewFlag THEN
    NewFlag = False
  ELSE
    t$ = in$
    LSET in$ = t$
  END IF

  WInput in$, GoodKey3$, 5, 11, ExitCode

  CF = ExitCode

  SELECT CASE ExitCode
  CASE 62
    RestartFlag = True
  END SELECT

  IF CF = -27 THEN RETURN

  IF RestartFlag THEN
    CLS
    RUN
  END IF

  IF CF = 80 THEN  'DownArrow
    NewFlag = True
    BackFlag = 0
    CurRecord = CurRecord + 1
    IF CurRecord > LastRecord THEN
      CurRecord = LastRecord
      GOTO 1080
    ELSE
      GOTO 1060
    END IF
  END IF

  IF CF = 72 THEN             'UpArrow
    NewFlag = True
    CurRecord = CurRecord - 1
    IF CurRecord < 1 THEN CurRecord = 1
    IF BackFlag = 0 THEN
      BackFlag = 1
      GOTO 1059
    ELSE
      GOTO 1060
    END IF
  END IF

  IF CF = 60 THEN              'f2key
    GOSUB displaypage
    GOTO 1100
  END IF

  IF CF = 59 THEN              'f1
    GOSUB FindMeter
    IF ScanCode = 1 THEN
      CurRecord = 1
      GOTO 1059
    END IF
    GOTO 1100
  END IF

  SELECT CASE CF
  CASE 11, 12
    GOTO GetInput
  END SELECT
  
  CLocate 5, 11
  CurRead# = VAL(in$)
  RSET Fmt2$ = QPTrim$(STR$(CurRead#))
  'PRINT USING Fmt1$; CurRead#
  PRINT Fmt2$;

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
  RSET Fmt2$ = QPTrim$(STR$(CalcUsage#))
  PRINT " Usage:   "; Fmt2$;
  'PRINT " Usage:   "; USING Fmt1$; CalcUsage#

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

  kk$ = ""
  WInput kk$, "", 8, 21, ExitCode

  CF = ExitCode
  IF CF = -27 THEN
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
  
FindMeter:
  CLS
  PRINT "Battery Power: "; GetPower + 1; "%"
  PRINT "Enter Meter Number"
  PRINT
  PRINT "F1 = First un-read"
  PRINT "F2 = Last un-read"
  PRINT "or (ESC-RETURN) "
  PRINT "==>";

  CF = 0
  fl = 15
  MNum$ = SPACE$(fl)

DO
  NewFlag = True

  WInput MNum$, GoodKey2$ + GoodKey3$, 7, 6, ExitCode

  CF = ExitCode

  IF CF = -27 OR LEN(MNum$) = 0 THEN
    RestartFlag = 0
    NewFlag = False
    EXIT DO
  END IF

  MNum$ = UCASE$(MNum$)
  
  'IF ((VAL(MNum$) = 0) AND (MNum$ = "SCAN")) THEN
  '  ScanCode = 1
  '  EXIT DO
  'END IF

  IF CF = 59 THEN
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
    EXIT DO
  END IF

  IF CF = 60 THEN
    Last = CurRecord
    GOSUB OpenFile
    TNumOfRecs& = LOF(1) \ UBPC3000RdRecLen
    FOR CurRecord = TNumOfRecs& TO 1 STEP -1
      GET 1, CurRecord, UBPC3000RDRec(1)
      IF UBPC3000RDRec(1).Book = CurBook THEN
        IF UBPC3000RDRec(1).ReadFlag <> "Y" THEN
          EXIT FOR
        END IF
      END IF
    NEXT
    GET 1, CurRecord, UBPC3000RDRec(1)
    GOSUB CloseFile
    EXIT DO
  END IF
  ScanCode = 0
  IF CF = -27 THEN
    EXIT DO
  END IF

  IF CF >= 60 AND CF <= 70 THEN
    GOTO FindMeter
  END IF

FindFirst:
  CLS
  CLocate 4, 4
  WhatMeter$ = QPTrim$(MNum$)
  PRINT "Scanning [ ]"
  GOSUB OpenFile
  FOR X = 1 TO LastRecord
    GET 1, X, UBPC3000RDRec(1)
    GOSUB Scanning
    trmeter$ = QPTrim$(UBPC3000RDRec(1).MeterID)
    IF WhatMeter$ = trmeter$ THEN
      IF UBPC3000RDRec(1).ReadFlag <> "Y" THEN
        CurRecord = X
        GOSUB CloseFile
        EXIT DO
      END IF
    END IF
  NEXT X
  GET 1, CurRecord, UBPC3000RDRec(1)
  GOSUB CloseFile
  CLocate 7, 1
  PRINT " METER: " + WhatMeter$
  PRINT " NOT Located"
  PRINT " Press <ENTER>"
  zz = BiosKey
  GOTO FindMeter

LOOP

RETURN

displaypage:
  IF CF = 60 THEN
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
    PRINT "New Notes (Y/N) [ ]"     ';
    YN$ = " "
    WInput YN$, "YNyn", 8, 18, ExitCode

    CF = ExitCode

    IF CF = 76 THEN RETURN

    'IF NOT (YN$ = "Y" OR YN$ = "N") THEN GOTO dp1
    IF YN$ = "N" OR CF = -1 OR CF = -27 THEN RETURN

dp2:

    CLocate 8, 1
    PRINT Blank$
    fl = 20

    Note1$ = LEFT$(QPTrim$(UBPC3000RDRec(1).Note1) + SPACE$(fl), 20)
    Note2$ = LEFT$(QPTrim$(UBPC3000RDRec(1).Note2) + SPACE$(fl), 20)
    Note3$ = LEFT$(QPTrim$(UBPC3000RDRec(1).Note3) + SPACE$(fl), 20)

Line1:

    WInput Note1$, GoodKey2$ + GoodKey3$, 3, 1, ExitCode
    CF = ExitCode
    IF CF = -27 THEN RETURN
    LSET UBPC3000RDRec(1).Note1 = QPTrim$(Note1$)
    IF CF = 72 THEN
      GOTO Line1
    END IF
    IF CF = 80 THEN
      GOTO Line2
    END IF

Line2:

    WInput Note2$, GoodKey2$ + GoodKey3$, 4, 1, ExitCode
    CF = ExitCode
    IF CF = -27 THEN RETURN
    LSET UBPC3000RDRec(1).Note2 = QPTrim$(Note2$)
    IF CF = 72 THEN
      GOTO Line1
    END IF
    IF CF = 80 THEN
      GOTO Line3
    END IF


Line3:

    WInput Note3$, GoodKey2$ + GoodKey3$, 5, 1, ExitCode

    CF = ExitCode
    IF CF = -27 THEN RETURN
    LSET UBPC3000RDRec(1).Note3 = QPTrim$(Note3$)
    IF CF = -1 THEN
      GOTO dp3
    END IF
    IF CF = 72 THEN
      GOTO Line2
    END IF
    IF CF = 13 THEN
      GOTO dp3
    ELSE
      GOTO Line3
    END IF

dp3:

    CLocate 8, 1
    PRINT "Save Notes (Y/N) [ ]"
    YN$ = " "
    WInput YN$, "YNyn", 8, 19, ExitCode
    CF = ExitCode
    IF CF = -27 THEN noteflag = 0: RETURN

    IF YN$ = "Y" THEN

dp4:
      CLocate 8, 1
      PRINT "Perm/Temp (P/T) [ ] "      ';
      CLocate 8, 18

      PT$ = " "
      WInput PT$, "PTpt", 8, 19, ExitCode
      CF = ExitCode
      IF CF = -27 THEN
        noteflag = 0
        RETURN
      END IF
      PT$ = UCASE$(PT$)
      IF (PT$ <> "P") AND (PT$ <> "T") THEN GOTO dp4
      noteflag = 1
      UBPC3000RDRec(1).NoteStatus = PT$
      GOSUB OpenFile
      PUT 1, CurRecord, UBPC3000RDRec(1)
      GOSUB CloseFile
      RETURN
    ELSEIF YN$ = "N" THEN
      RETURN
    END IF
    IF PT$ = "N" THEN GOTO dp2
    GOTO dp3
  END IF
  RETURN
  'must go back to first page

Scanning:
  CLocate 4, 14
  Char = Char + 1
  IF Char > 4 THEN Char = 1
  PRINT MID$("\|/-", Char, 1)
RETURN


OpenFile:
  OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS #1 LEN = UBPC3000RdRecLen
RETURN

CloseFile:
  CLOSE #1
RETURN

SUB CLocate (BYVAL Row%, BYVAL Col%)
  LOCATE Row, Col
END SUB

