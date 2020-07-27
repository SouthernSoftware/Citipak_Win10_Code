DECLARE SUB ShowEditScrn ()
DECLARE SUB WaitForAKey ()
DEFINT A-Z
DECLARE SUB BackLightToggle ()
DECLARE SUB WInput (Edit$, GoodKey$, Row%, Col%, KeyCode%)
DECLARE SUB SetFont ()
DECLARE SUB CLocate (BYVAL Row%, BYVAL Col%)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB QPSound (Frequency%, Duration%)
DECLARE FUNCTION GetPower ()
DECLARE FUNCTION BiosKey% ()

'$INCLUDE: 'ubsetup.bi'

  CONST False = 0, True = NOT False

  DIM SHARED UBPC3000RDRec(1) AS UBPC3000ReadRecType
  UBPC3000RdRecLen = LEN(UBPC3000RDRec(1))

  GoodKey1$ = " abcdefghijklmnopqrstuvwxyz"
  GoodKey2$ = UCASE$(GoodKey1$)
  GoodKey3$ = "1234567890"


  Ast$ = STRING$(20, "*")
  Blank$ = STRING$(20, " ")
  Fmt1$ = "########"
  
  SetFont

Top:
  GOSUB GetWhatRoute
  GOSUB SearchForRoute
  IF FoundIt THEN
    ShowScrn = True
    GOTO EditLoop
  ELSE
    GOSUB NoRoute
    GOTO Top
  END IF


EditLoop:
  IF ShowScrn THEN
    ShowScrn = False
    ShowEditScrn
  END IF
  GOSUB GetInput

GOTO EditLoop


ExitProg:
END

GetInput:

  KeyCode = 0
  fl = 8

  IF NewFlag THEN
    NewFlag = False
    in$ = SPACE$(fl)
  ELSE
    t$ = in$
    in$ = SPACE$(fl)
    LSET in$ = t$
  END IF

  WInput in$, GoodKey3$, 5, 11, KeyCode

  IF KeyCode = -27 THEN RETURN

  SELECT CASE KeyCode
  CASE -27

  CASE 62                     'F4
    GOTO Top
  CASE 80                     'DownArrow
    BackFlag = False
    CurRecord = CurRecord + 1
    GOSUB GetReadRec

    IF CurRecord > LastRecord THEN
      CurRecord = LastRecord
      'GOTO LastInRoute
    ELSE
      ShowScrn = True
      'GOSUB
    END IF

  CASE 72                  'UpArrow
    BackFlag = True
    'NewFlag = True
    CurRecord = CurRecord - 1
    IF CurRecord < 1 THEN
      CurRecord = 1
    ELSE
      ShowScrn = True
      GOSUB GetReadRec
    END IF
  CASE 60                   'f2key
    'GOSUB displaypage
    'GOTO 1100
  'END IF

  CASE 59                   'F1
    'GOSUB FindMeter
    IF ScanCode = 1 THEN
      CurRecord = 1
      'GOTO 1059
    END IF
    'GOTO 1100
  CASE 11, 12
  '  GOTO GetInput

  END SELECT

RETURN



SearchForRoute:

  FoundIt = False
  CLS
  PRINT
  PRINT "Searching for Route: "; CurBook
  LOCATE 4, 4
  PRINT "Scanning [ ]"
  GOSUB ShowScanProgress
  GOSUB OpenFile
  LastRecord = LOF(1) / UBPC3000RdRecLen

  FOR CurRecord = 1 TO LastRecord
    GET #1, CurRecord, UBPC3000RDRec(1)
    GOSUB ShowScanProgress
    IF UBPC3000RDRec(1).Book = CurBook THEN
      FoundIt = True
      EXIT FOR
    END IF
  NEXT
  GOSUB CloseFile
  'IF FoundIt THEN

RETURN



  CurRecord = 1


ShowScanScrn:
  LOCATE 4, 4
  PRINT "Scanning [ ]"
RETURN


ShowScanProgress:
  LOCATE 4, 14
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


GetWhatRoute:
  CLS
  PRINT "ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿"          '1
  PRINT "³       HUSKY HandHeld       ³"          '2
  PRINT "³    Meter Reading System    ³"          '3
  PRINT "³ (C)1997 Southern Software  ³"          '4
  PRINT "³                            ³"          '5
  PRINT "³ Enter Route#               ³"          '6
  PRINT "³                            ³"          '7
  PRINT "ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ";         '8

  Route$ = SPACE$(2)
  WInput Route$, GoodKey3$, 6, 16, KeyCode
  IF LEN(QPTrim$(Route$)) = 0 OR KeyCode = -27 THEN
    CLS
    LOCATE 3, 1
    PRINT "Program Exited."
    END
  END IF

  CurBook = VAL(Route$)

  GOTO ExitGetWhatRoute

ExitGetWhatRoute:

RETURN

GetReadRec:
  GOSUB OpenFile
  GET #1, CurRecord, UBPC3000RDRec(1)
  GOSUB CloseFile
RETURN

NoRoute:
  CLS
  'QPSound 1000, 1
  PRINT "ROUTE: "; CurBook
  PRINT
  PRINT "NOT FOUND"
  PRINT
  PRINT
  PRINT "Press any Key. . .";
  'QPSound 1000, 3
  WaitForAKey
RETURN

SUB ShowEditScrn

  SHARED Fmt1$

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
    LOCATE 6, 2
    PRINT "NOTE!";
    QPSound 1000, 1
  END IF
  CLocate 5, 1
  PRINT " Current..";
  fl = -8

END SUB

SUB WaitForAKey
  zz = BiosKey
END SUB

