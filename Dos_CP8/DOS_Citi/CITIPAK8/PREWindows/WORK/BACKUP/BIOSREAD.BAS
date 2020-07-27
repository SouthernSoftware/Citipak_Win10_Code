' New hand held program for Version 8.2w
DEFINT A-Z
DECLARE SUB WCursorSetSize (BYVAL Top%, BYVAL Bot%)
DECLARE SUB WInput (Edit$, GoodKey$, Row%, Col%, ExitCode%)
DECLARE SUB WLocate (BYVAL Row%, BYVAL Col%)
DECLARE SUB WCls ()
DECLARE SUB WPrint (Text$)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB QPSound (Frequency%, Duration%)
DECLARE FUNCTION BiosKey% ()
DECLARE FUNCTION FUsing$ (Number$, Fmt$)
  '$INCLUDE: 'ubsetup.bi'
CONST False = 0, True = NOT False

GoodKey1$ = " abcdefghijklmnopqrstuvwxyz"
GoodKey2$ = UCASE$(GoodKey1$)
GoodKey3$ = "1234567890"

  STACK 10000
  Ast$ = STRING$(20, "*")
  Blank$ = STRING$(20, " ")
  Fmt1$ = "########"
  REDIM UBPC3000RDRec(1) AS UBPC3000ReadRecType
  UBPC3000RdRecLen = LEN(UBPC3000RDRec(1))
  CL$ = CHR$(10) + CHR$(13)

  REDIM MText$(1 TO 9)
    MText$(1) = "*********************"
    MText$(2) = "*       Meter       *"
    MText$(3) = "*   Reading System  *"
    MText$(4) = "* Southern Software *"
    MText$(5) = "*                   *"
    MText$(6) = "* Route#            *"
    MText$(7) = "*                   *"
    MText$(8) = "*********************"

  REDIM F1Text$(1 TO 6)
    F1Text$(1) = "Enter Meter Number"
    F1Text$(2) = "" 'SCAN  for un-read"
    F1Text$(3) = "F1 = First un-read"
    F1Text$(4) = "F2 = Last un-read"
    F1Text$(5) = "or (ESC-RETURN) "
    F1Text$(6) = "==>"

Top:
  WCls
  FOR zz = 1 TO 8
    WLocate zz, 1
    WPrint MText$(zz)
  NEXT

  fl = 2
  Route$ = SPACE$(ABS(fl))
  WInput Route$, GoodKey3$, 6, 10, ExitCode

  IF LEN(QPTrim$(Route$)) = 0 OR ExitCode = -27 THEN
    WCls
    WLocate 3, 1
    WPrint "Program Exited." + CL$
    END
  END IF

  CurBook = VAL(Route$)

  WCls
  WLocate 2, 1
  WPrint "Searching for:" + STR$(CurBook)
  WLocate 4, 1
  WPrint "Scanning [ ]"
  GOSUB Scanning
  GOSUB OpenFile
  LastRecord = LOF(1) / UBPC3000RdRecLen
  GOSUB CloseFile
  CurRecord = 1
  NewFlag = True

1059
  IF ScanCode = 1 OR BackFlag = 1 THEN
    WCls
    WLocate 4, 1
    WPrint "Scanning [ ]"
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
    WCls
    WPrint "NO MORE METERS!!!"
    QPSound 1000, 1
    QPSound 1000, 3
    WLocate 5, 1
    WPrint "PRESS ANY KEY. . ."
    zz = BiosKey
    RUN
  ELSE
    WCls
    WPrint "ROUTE:" + STR$(CurBook)
    WLocate 3, 1
    WPrint "NOT FOUND"
    WLocate 6, 1
    WPrint "Press any Key. . ."   ';
    zz = BiosKey
    GOTO Top
  END IF

1100
  WCls
  BeenHere = -1
  IF UBPC3000RDRec(1).ReadFlag = "Y" THEN
    WPrint "*"    ';
  ELSE
    WPrint " "     ';
  END IF
  WPrint UBPC3000RDRec(1).CustName
  WLocate 2, 1
  WPrint " " + UBPC3000RDRec(1).ServAddress
  WLocate 3, 1
  WPrint " Mtr# " + UBPC3000RDRec(1).MeterID
  SELECT CASE UBPC3000RDRec(1).MeterType
  CASE "E"
    WPrint " *Elec"
  CASE "D"
    WPrint " *Dmnd"
  CASE "W"
    WPrint " *Wat"
  CASE "C"
    WPrint " *W/S"
  CASE "S"
    WPrint " *Sew"
  CASE "G"
    WPrint " *Gas"
  END SELECT
  WLocate 4, 1
  WPrint " Previous."
  IF UBPC3000RDRec(1).ReadFlag = "Y" THEN
    WPrint FUsing$(STR$(UBPC3000RDRec(1).PastRead), Fmt1$)
  ELSE
    WPrint FUsing$(STR$(UBPC3000RDRec(1).CurRead), Fmt1$)
  END IF

  IF UBPC3000RDRec(1).ReadFlag = "Y" THEN
    WLocate 8, 1
    WPrint " Last Rdg."   '; USING Fmt1$; UBPC3000RDRec(1).CurRead
    WPrint FUsing$(STR$(UBPC3000RDRec(1).CurRead), Fmt1$)
    in$ = QPTrim$(STR$(UBPC3000RDRec(1).CurRead))
    NewFlag = 0
  END IF
  IF ASC(LEFT$(UBPC3000RDRec(1).Note1, 1)) > 32 OR ASC(LEFT$(UBPC3000RDRec(1).Note2, 1)) > 32 OR ASC(LEFT$(UBPC3000RDRec(1).Note3, 1)) > 32 THEN
    WLocate 6, 1
    WPrint "*Note"
    QPSound 1000, 1
  END IF
  WLocate 5, 1
  WPrint " Current.."
  WLocate 5, 11
 
GetInput:
  CF = 0
  fl = 8

  IF NewFlag THEN
    NewFlag = False
    in$ = SPACE$(fl)
  ELSE
    t$ = in$
    in$ = SPACE$(fl)
    LSET in$ = t$
  END IF

  WInput in$, GoodKey3$, 5, 11, ExitCode
  
  CF = ExitCode

  SELECT CASE ExitCode
  CASE 62
    RestartFlag = -1
  END SELECT

  IF CF = -27 THEN RETURN
  
  IF RestartFlag THEN
    WCls
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

  IF CF = 59 THEN
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
  
  WLocate 5, 11
  CurRead# = VAL(in$)
  WPrint FUsing$(STR$(CurRead#), Fmt1$)
  WLocate 6, 10
  WPrint "---------"

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

  WLocate 7, 1
  WPrint " Usage:   " + FUsing$(STR$(CalcUsage#), Fmt1$)

  IF CalcUsage# < UBPC3000RDRec(1).LowRead THEN
    WLocate 7, 20
    WPrint "L"
    QPSound 1000, 2
  END IF
  
  IF CalcUsage# > UBPC3000RDRec(1).HighRead THEN
    WLocate 7, 20
    WPrint "H"
    QPSound 1000, 2
  END IF
  
1152
  WLocate 8, 1
  WPrint " ESC=NO Enter=OK."
  fl = 0
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
  NewFlag = True
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
  WCls
  FOR zz = 1 TO 6
    WLocate zz, 1
    WPrint F1Text$(zz)
  NEXT

  CF = 0
  fl = 15
  MNum$ = SPACE$(ABS(fl))

DO
  NewFlag = True

  WInput MNum$, GoodKey2$ + GoodKey3$, 6, 6, ExitCode
  
  CF = ExitCode
  
  IF CF = -27 OR LEN(MNum$) = 0 THEN
    RestartFlag = 0
    NewFlag = False
    EXIT DO
  END IF
  
  MNum$ = UCASE$(MNum$)
  

  IF ((VAL(MNum$) = 0) AND (MNum$ = "SCAN")) THEN
    ScanCode = 1
    'GOTO FindFirst
    EXIT DO
  END IF

  IF (VAL(MNum$) = 0) AND (MNum$ = "FIRST") OR CF = 59 THEN
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

  IF (VAL(MNum$) = 0) AND (MNum$ = "LAST") OR CF = 60 THEN
    'CurRecord = LastRecord
    Last = CurRecord
    GOSUB OpenFile
    TNumOfRecs& = LOF(1) \ UBPC3000RdRecLen
    FOR CurRecord = TNumOfRecs& TO 1 STEP -1
      GET 1, CurRecord, UBPC3000RDRec(1)
      IF UBPC3000RDRec(1).Book = CurBook THEN
        IF UBPC3000RDRec(1).ReadFlag <> "Y" THEN
          EXIT FOR
          'Last = CurRecord
        END IF
      END IF
    NEXT
    'CurRecord = Last
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
  WCls
  WLocate 4, 4
  WhatMeter$ = QPTrim$(MNum$)
  WPrint "Scanning [ ]"
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
  WLocate 7, 1
  WPrint " METER: " + WhatMeter$ + CL$
  WPrint " NOT Located" + CL$
  WPrint " Press <ENTER>" + CL$
  zz = BiosKey
  GOTO FindMeter

LOOP

RETURN
  
displaypage:
  IF CF = 60 THEN
    WCls
    WLocate 1, 1
    WPrint Ast$
    WLocate 1, 8
    WPrint "Notes" + CL$
    WPrint Ast$ + CL$
    WPrint UBPC3000RDRec(1).Note1 + CL$
    WPrint UBPC3000RDRec(1).Note2 + CL$
    WPrint UBPC3000RDRec(1).Note3 + CL$
    WPrint Ast$ + CL$
    
dp1:
    WLocate 8, 1
    WPrint "New Notes (Y/N) [ ]"     ';
    fl = 1
    YN$ = SPACE$(ABS(fl))
    WInput YN$, "YNyn", 8, 18, ExitCode

    CF = ExitCode

    IF CF = 76 THEN RETURN

    'IF NOT (YN$ = "Y" OR YN$ = "N") THEN GOTO dp1
    IF YN$ = "N" OR CF = -1 OR CF = -27 THEN RETURN

dp2:

    WLocate 8, 1
    WPrint Blank$
    fl = 20
    Note1$ = SPACE$(fl)
    Note2$ = Note1$
    Note3$ = Note1$
    LSET Note1$ = UBPC3000RDRec(1).Note1
    LSET Note2$ = UBPC3000RDRec(1).Note2
    LSET Note3$ = UBPC3000RDRec(1).Note3

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

    WLocate 8, 1
    WPrint "Save Notes (Y/N) [ ]"
    fl = 1
    YN$ = SPACE$(ABS(fl))
    WInput YN$, "YNyn", 8, 19, ExitCode
    CF = ExitCode
    IF CF = -27 THEN noteflag = 0: RETURN

    IF YN$ = "Y" THEN

dp4:
      WLocate 8, 1
      WPrint "Perm/Temp (P/T) [ ] "      ';
      WLocate 8, 18

      fl = 1
      PT$ = SPACE$(fl)
      WInput PT$, "PTpt", 8, 19, ExitCode
      CF = ExitCode

      IF CF = -27 THEN noteflag = 0: RETURN
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
  WLocate 4, 11
  Char = Char + 1
  IF Char > 4 THEN Char = 1
  WPrint MID$("\|/-", Char, 1)
RETURN
  
  
OpenFile:
  OPEN "UBCUSTTR.DAT" FOR RANDOM SHARED AS #1 LEN = UBPC3000RdRecLen
RETURN
  
CloseFile:
  CLOSE #1
RETURN
  

