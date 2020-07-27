DEFINT A-Z
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE SUB SelectRate2Edit (RecNo%)
DECLARE SUB AddEditRateTbl (EdFlag%)
DECLARE SUB LoadRateTblRec (Form$(), TblNum%)
DECLARE SUB SaveRateTblRec (Form$(), RecNo%)
DECLARE SUB LoadMeterRec (Form$())
DECLARE SUB AddNewMeter (Cust$, EdFlag)
DECLARE SUB EditMeter ()
DECLARE SUB SaveMeterRec (Form$())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB RestScrn (Array%())
DECLARE SUB AddEditCustomer (EdFlag%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB LoadCustRec (Form$())
DECLARE SUB SaveCustRec (Form$())
DECLARE SUB AddNewCustomer ()
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB SaveSetUpRec (Form$())
DECLARE SUB LoadSetUpRec (Form$())
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB EditCust ()
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB StuffBuf (ky$)

'$INCLUDE: 'DefCnf.BI'

DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)

DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf AS Config)
DECLARE SUB WaitForAction ()

'DECLARE SUB VendorMenu ()
DECLARE SUB HideCursor ()

DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)

'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
''$INCLUDE: 'ubfiles.BI'
''$INCLUDE: 'init.BI'

'$INCLUDE: 'PageInfo.BI'                  'Form Page info

'$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
''$INCLUDE: 'ScrAux.bi'

'$INCLUDE: 'SetCnf.bi'
'$INCLUDE: 'ubcust.bi'
'$INCLUDE: 'ubrate.BI'

CONST False = 0, True = NOT False
 
 STACK 8000

   '--Dim the choice array to the number of menu items
   REDIM MChoice$(1 TO 4)

   MChoice$(1) = "Add a New Rate Table "
   MChoice$(2) = "Edit an Existing Rate Table  "
   MChoice$(3) = "Exit to OS"

   MaxLen = 0     'Set menu width to zero
   BoxBot = 20    'limit the box length to go no lower than line 20
   Action = 0     '0 means stay in the menu until they select something
   Choice = 1     'Pre-load choice to highlight

   '--Find max menu width
   FOR Cnt = 1 TO UBOUND(MChoice$)
     TLen = LEN(MChoice$(Cnt))
     IF TLen > MaxLen THEN
       MaxLen = TLen
     END IF
   NEXT

   '--Center Menu within Screen
   Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
   Col = ((80 - MaxLen) \ 2) - 2
   Help$ = "Customer Maintenance Menu"
   
   DO

      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      LibFile2Scrn "UB.QSL", "MENUBAK", MonoCode, -1, ErrorCode

      TitleBox 3, Col, MaxLen + 3, "Rate Table Menu ", Cnf
      TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

      PrintTitle user$
      PrintHelp Help$

      ShowCursor

      VertMenu MChoice$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf

      IF ky$ = CHR$(27) THEN EXIT DO 'choice = 0

      SELECT CASE Choice
          CASE 1
            AddEditRateTbl False
          CASE 2
            AddEditRateTbl True
          CASE 3
             CLS
             END
      END SELECT
   LOOP

   RUN "ubmisc"

SUB AddEditRateTbl (EdFlag)
  
  REDIM ScrnArray(0)

  LibName$ = "UB"
  ScrnName$ = "UBRATE"
  NumScrns = 1
  
  '--define the multi-choice fields
  SHARED Choice$()
  REDIM Choice$(0 TO 4, 0)
  
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT

  IF EdFlag THEN
    SelectRate2Edit RecNo
    IF RecNo > 0 THEN
      LoadRateTblRec Form$(), RecNo
      UnPackBuffer 0, 0, Form$(), Fld()
    ELSE
      GOTO ExitEdit
    END IF
  ELSE

    RecNo = GetNumRateRecs + 1
  END IF

  '--Set screen number to one and display screen
  Scr = 1
  LibFile2Scrn LibName$, ScrnName$, MonoCode, Attribute%, ErrCode

  ShowCursor

  Action = 1

  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
      CASE -68 'F10Key
        SaveScrn ScrnArray()
        SaveRateTblRec Form$(), RecNo
        ExitFlag = True
        RestScrn ScrnArray()

    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
        CASE 21        'Look for the f10 or esc button
          SELECT CASE Frm(1).MCol
            CASE 43 TO 54    '--Save Button
              PressButton -68, 21, 43, 54
            CASE 62 TO 72    '--cancel button
              PressButton 27, 21, 57, 70
          END SELECT
      END SELECT 'row
    END IF

    '--Check screen page
  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag

  IF Frm(1).KeyCode = 27 THEN GOTO ExitEdit

ExitEdit:
  HideCursor
EXIT SUB

END SUB

FUNCTION GetNumRateRecs
  REDIM UBRateTblRec(1) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))
  GetNumRateRecs = FileSize("UBRATE.DAT") \ UBRateTblRecLen
  ERASE UBRateTblRec
END FUNCTION

SUB LoadRateTblRec (Form$(), TblNum)

  REDIM UBRateTblRec(1) AS UBRateTblRecType

  UBRateTblRecLen = LEN(UBRateTblRec(1))

  Form$(0, 0) = SPACE$(UBRateTblRecLen)
'
  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM AS UBFile LEN = UBRateTblRecLen
  GET UBFile, TblNum, UBRateTblRec(1)
  CLOSE

  BCopy VARSEG(UBRateTblRec(1)), VARPTR(UBRateTblRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), UBRateTblRecLen, 0

  ERASE UBRateTblRec

END SUB

SUB SaveRateTblRec (Form$(), RecNo)

  REDIM UBRateTblRec(1) AS UBRateTblRecType

  UBRateTblRecLen = LEN(UBRateTblRec(1))

  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBRateTblRec(1)), VARPTR(UBRateTblRec(1)), UBRateTblRecLen, 0

  UBFile = FREEFILE

  OPEN "UBRATE.DAT" FOR RANDOM AS UBFile LEN = UBRateTblRecLen
  PUT UBFile, RecNo, UBRateTblRec(1)
  CLOSE

  LibFile2Scrn "UB", "UPDATEOK", MonoCode, Attribute%, ErrCode
  WaitForAction

  ERASE UBRateTblRec

END SUB

SUB SelectRate2Edit (RecNo)

  REDIM UBRateTblRec(1) AS UBRateTblRecType

  UBRateTblRecLen = LEN(UBRateTblRec(1))

  NumOfRateRecs = GetNumRateRecs

  REDIM MChoice$(1 TO NumOfRateRecs)
  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM AS UBFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRateRecs
    GET UBFile, Cnt, UBRateTblRec(1)
    MChoice$(Cnt) = SPACE$(50)
    LSET MChoice$(Cnt) = UBRateTblRec(1).RateCode
    MID$(MChoice$(Cnt), 6) = UBRateTblRec(1).RateDesc
  NEXT
  CLOSE

   MaxLen = 50     'Set menu width to zero
   BoxBot = 17    'limit the box length to go no lower than line 20
   Action = 0     '0 means stay in the menu until they select something
   Choice = 1     'Pre-load choice to highlight

   TText$ = SPACE$(MaxLen + 4)
   LSET TText$ = " Code  Description"

   '--Center Menu within Screen
   Row = 8
   Col = 15

   DO

      '--Set upper left corner of menu, turn off the cursor
      LOCATE Row, Col, 0
      LibFile2Scrn "UB.QSL", "MENUBAK", MonoCode, -1, ErrorCode
      ShowCursor
      QPrintRC TText$, Row - 1, Col, 112
      VertMenu MChoice$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf
      IF ky$ = CHR$(27) THEN
        RecNo = 0
        ExitFlag = True
      ELSE
        RecNo = Choice
        ExitFlag = True
      END IF

   LOOP UNTIL ExitFlag

END SUB

