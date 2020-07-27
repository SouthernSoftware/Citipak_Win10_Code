DEFINT A-Z
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB GetCursor (X, Y, Button)
DECLARE FUNCTION ASCII% (Strng$)
DECLARE FUNCTION CheckUserAbort% ()
DECLARE FUNCTION DOSError% ()
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FLInput$ (Handle%, Buff$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION LibNumberOfFields (LibName$, FrmName$)
DECLARE FUNCTION PRNReady% (LPTNumber%)
DECLARE FUNCTION Peek1% (Segment%, OffSet%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION WhichError% ()
DECLARE SUB AskUserCheckPrinter (OkFlag%)
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB AskUserFixPrinter (Retry%)
'DECLARE SUB BLPrint (LPTNo%, Work$, ErrCount%)
DECLARE SUB BlockClear ()
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, X%, Y%)
DECLARE SUB CursorOff ()
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB EditForm (Form$(), Fld() AS ANY, Frm AS ANY, Cnf AS ANY, Action)
DECLARE SUB FClose (Handle%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FileView (FileName$, Ky, Action, FVI AS ANY, SEG Array)
DECLARE SUB LibGetFldDef (LibName$, FrmName$, StartEl, Fld() AS ANY, Form$(), ErrCode)
DECLARE SUB Poke1 (Segment%, OffSet%, ByteVal%)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB WaitForAction ()
DECLARE SUB RestScrn (ScrnArray%())
DECLARE SUB SaveScrn (ScrnArray%())
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)

'$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.'COMMON SHARED /ConfigSet/ Cnf AS Config, MonoCode
'$INCLUDE: 'FieldInf.bi'                '        "        field information
'$INCLUDE: 'FormEdit.bi'                '        "        form editing info.
'$INCLUDE: 'PRFiles.bi'
''$INCLUDE: 'SetCnf.bi'                  'Defines TYPE for monitor/color info.'COMMON SHARED /ConfigSet/ Cnf AS Config, MonoCode

TYPE FVInfo
    Colr        AS INTEGER      'Text color (white on black default)
    ULRow       AS INTEGER      'these four describe the window's corners
    ULCol       AS INTEGER
    LRRow       AS INTEGER
    LRCol       AS INTEGER
    HorizOffset AS INTEGER      'left margin within the window (see below)
    LoPtr       AS INTEGER      'used internally -- do not change!
    FileHandle  AS INTEGER      'internal (the DOS file handle if you care)
    EndOfFile   AS INTEGER      'used internally -- do not change!
    LineNumber  AS LONG         'line number of top line displayed, READ ONLY
    TabStop     AS INTEGER      'Tab stop setting (see below)
    FileSeek    AS LONG         'force read anywhere in the file (see below)
    FileOffset  AS LONG         'used internally -- do not change!
END TYPE

CONST False = 0
CONST True = NOT False

SUB AskUserCheckPrinter (OkFlag)

  ActMRow = 14
  BlockClear
  DisplayMiscScrn CheckReady      'make sure printer is ready message

  DO

    Get.Moose.OR.Key Ky$, MooseButton, MRow, MCol

    IF MooseButton THEN
      MRow = (MRow \ 8) + 1     'Convert MCol, MRow to Row and Col.
      MCol = (MCol \ 8) + 1
      IF MRow = ActMRow THEN
        SELECT CASE MCol
        CASE 28 TO 36           'print it
          PressButton ASCII("P"), ActMRow, 28, 36
        CASE 42 TO 53           'Cancel printing
          PressButton EscKey, ActMRow, 42, 53
        END SELECT
      END IF
    END IF
  
    IF LEN(Ky$) THEN
   
      Choice = INSTR(1, "Pp" + CHR$(27), Ky$)
      SELECT CASE Choice
        CASE 1, 2
          OkFlag = True
          EXIT DO
        CASE 3
          BlockClear
          DisplayMiscScrn Canceled      'Ask the user are they sure
          WaitForAction
          OkFlag = False
          EXIT DO
      END SELECT
    END IF
  LOOP
END SUB

SUB AskUserFixPrinter (Retry)
  ActMRow = 14
  Retry = False
  BlockClear

  DisplayMiscScrn NotReady      'ask user to fix printer message

  DO

    Get.Moose.OR.Key Ky$, MooseButton, MRow, MCol
   
    IF MooseButton THEN
      MRow = (MRow \ 8) + 1     'Convert MCol, MRow to Row and Col.
      MCol = (MCol \ 8) + 1

      IF MRow = ActMRow THEN
        SELECT CASE MCol
        CASE 41 TO 52           'Retry
          PressButton ASCII("R"), ActMRow, 41, 52
        CASE 28 TO 39           'Cancel
          PressButton EscKey, ActMRow, 28, 39
        END SELECT
      END IF
    END IF
   
    IF LEN(Ky$) THEN
      Choice = INSTR(1, "R" + CHR$(13) + CHR$(27), UCASE$(Ky$))

      SELECT CASE Choice
        CASE 1, 2
          Retry = True
          BlockClear
          DisplayMiscScrn Printing       'printing msg
         
        CASE 3
          BlockClear
          DisplayMiscScrn Canceled        '
          WaitForAction
          EXIT DO
      END SELECT
    END IF
  LOOP UNTIL Retry

END SUB

SUB BlockClear
  LibFile2Scrn "CMNDLG", "BAKCLEAR", MonoCode, -1, ErrorCode
END SUB

FUNCTION CheckUserAbort%
 
  ActMRow = 13
 
  Ky$ = INKEY$              'ButNo,DnFlg,NoPresses,col,row
  ButtonPress 1, 0, MooseButton, MCol, MRow

  IF LEN(Ky$) OR MooseButton THEN
   
    IF MooseButton THEN
      MRow = (MRow \ 8) + 1     'Convert MCol, MRow to Row and Col.
      MCol = (MCol \ 8) + 1
      IF MRow = ActMRow THEN
        SELECT CASE MCol
        CASE 35 TO 46           'cancel button clicked
          PressButton EscKey, ActMRow, 35, 46
        END SELECT
      END IF
    END IF
   
    IF LEN(Ky$) THEN
      Choice = INSTR(1, CHR$(27), Ky$)
      IF Choice THEN
        ActMRow = 14
        BlockClear
        DisplayMiscScrn AskCancel
        DO
          DO
            Ky$ = INKEY$              'ButNo,DnFlg,NoPresses,col,row
            ButtonPress 1, 0, MooseButton, MCol, MRow
          LOOP UNTIL LEN(Ky$) OR MooseButton
         
          IF MooseButton THEN
            MRow = (MRow \ 8) + 1     'Convert MCol, MRow to Row and Col.
            MCol = (MCol \ 8) + 1
            IF MRow = ActMRow THEN
              SELECT CASE MCol
                CASE 34 TO 39           'yes cancel
                  PressButton ASCII("Y"), ActMRow, 34, 39
                CASE 41 TO 46           'no continue
                  PressButton ASCII("N"), ActMRow, 41, 46
              END SELECT
            END IF
          END IF
 
          IF LEN(Ky$) THEN
            Choice = INSTR(1, "NnYy", Ky$)
            SELECT CASE Choice
              CASE 1, 2              'No continue
                BlockClear
                DisplayMiscScrn Printing
                CheckUserAbort = False
                EXIT DO
              CASE 3, 4           'Esc cancel
                BlockClear
                DisplayMiscScrn Canceled
                WaitForAction
                CheckUserAbort = True
                EXIT DO
            END SELECT
          END IF
        LOOP
      END IF
    END IF
  END IF
END FUNCTION

SUB DisplayMiscScrn (ScrnName$) STATIC

  LibFile2Scrn "CMNDLG", ScrnName$, MonoCode, 0, 0

END SUB

SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)

    DO
      Ky$ = INKEY$              'ButNo,DnFlg,NoPresses,col,row
      'ButtonPress 1, 0, MooseButton, MCol, MRow
      ButtonPress 1, N, MooseButton, MCol, MRow  ' ----- Check Mouse activity
      IF N AND 2 THEN                'if it was the right moose button and
        DO                           'if they are still holding it down then
          GetCursor 0, 0, NewStatus  'wait till they let go of the button
        LOOP WHILE NewStatus
        ButtonPress 0, 0, 0, 0, 0    'this has the effect of clearing all
        Ky$ = CHR$(EscKey)           'moose buttons.
      END IF
    LOOP UNTIL LEN(Ky$) OR MooseButton

END SUB

SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
  
  Cols = ButtonRCol - ButtonLCol + 1
  REDIM Btn(Cols)
  MScrnSave ButtonRow, ButtonLCol, ButtonRow, ButtonRCol, Btn(1)

  MPaintBox ButtonRow, ButtonLCol, ButtonRow, ButtonRCol, 15
  
'  Pause 2
  DO
    GetCursor 0, 0, i
  LOOP WHILE i
  ButtonPress 0, 0, 0, 0, 0

  MScrnRest ButtonRow, ButtonLCol, ButtonRow, ButtonRCol, Btn(1)
  
  IF KeyCode < 0 THEN
    StuffBuf CHR$(0) + CHR$(ABS(KeyCode))
  ELSEIF KeyCode > 0 THEN
    StuffBuf CHR$(KeyCode)
  END IF

  ERASE Btn

END SUB

SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)

  REDIM TempScrn(1)

  IF EntryPoint = 2 THEN
    EntryPoint = 0
    GOTO ShowFileOnScreen
  ELSEIF EntryPoint = 4 THEN
    EntryPoint = 0
    GOTO PrintFile
  ELSEIF EntryPoint = 5 THEN
    EntryPoint = 1
    GOTO PrintFile
  END IF

  ViewWhere = False
  ExitFlag = False

  REDIM Frm(1) AS FormInfo

  FormName$ = "VIEWPRNT"
  NumFlds = LibNumberOfFields("CMNDLG", FormName$)
  REDIM Form$(NumFlds, 2)             'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo     'DIM the field information array
  StartEl = 0                         'Load first form at array start
  LibGetFldDef "CMNDLG", FormName$, StartEl, Fld(), Form$(), ErrCode

  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1

  '----- Setup TYPE for setting and reading form editing information.

  Frm(1).FldNo = 1                 'Start editing on field #1
  Frm(1).InsStat = False           'Set insert state (True = Insert on)
  Frm(1).StartEl = 0               'Set form starting element to 0 and
  TitleRow = 8
  TitleCol = 40 - (LEN(RptTitle$) \ 2) + 1
  BlockClear

  DisplayMiscScrn FormName$

  QPrintRC RptTitle$, TitleRow, TitleCol, 126

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
      CASE EscKey
        ExitFlag = True
      CASE SKey, SKey + 32
        ViewWhere = 2
      CASE PKey, PKey + 32
        ViewWhere = 1
    END SELECT

  LOOP UNTIL ViewWhere OR ExitFlag        'proper key not set

  LOCATE , , 0
  ERASE Form$, Fld

  IF ExitFlag THEN EXIT SUB

  SELECT CASE ViewWhere
    CASE 1               'Print the file
      GOTO PrintFile
    CASE 2               'Show it on the screen
      GOTO ShowFileOnScreen
  END SELECT

PrintFile:

  RetCode = 0

  IF Exist(FileName$) THEN
    FileLen& = FileSize(FileName$)
  ELSE
    RetCode = 1
  END IF
  IF FileLen& = 0 THEN
    RetCode = 2
  END IF
  IF RetCode THEN
    GOTO AbortExit
  END IF

  IF EntryPoint = 1 THEN
    GOTO EntryPoint1
  END IF

  AskUserCheckPrinter Ok

  IF NOT Ok THEN
    RetCode = -1
    GOTO AbortExit
  END IF

EntryPoint1:

  BufSize = 160                     'enough to read up to 138 chars
  Buffer$ = SPACE$(BufSize)         '  (or use whatever is appropriate)

  IF LPTPort <= 0 THEN LPTPort = 1

'**** 03/30/11  Kayren and crew have new windows 7 machines.
  LPTPort = 2

  BlockClear
  DisplayMiscScrn "PRINTING"
  Device$ = "LPT" + QPTrim$(STR$(LPTPort)) + ":"
  'Device$ = "SCRN:"

'**** 10/27/95 Changed to use BASIC device I/O
  ON LOCAL ERROR RESUME NEXT         'if there is an error just keep going

  WIDTH Device$, 255
  OPEN Device$ FOR OUTPUT AS #1

  FOpenS FileName$, Handle           'open for QuickPak Professional Binary

  DO
    ToPrint$ = FLInput$(Handle, Buffer$) 'get a line of input
    WHILE WhichError% = 83               'a "Buffer too small" error means
      ToPrint$ = ToPrint$ + FLInput$(Handle, Buffer$)  'we have to go back
    WEND                                               'and get some more
    IF DOSError% THEN EXIT DO            'exit if error (probably end of file)
    PRINT #1, ToPrint$
    IF ERR <> 0 THEN
      ERR = 0
      AskUserFixPrinter Retry
      IF NOT Retry THEN GOTO AbortExit
      IF ErrCode > 0 THEN
        ToPrint$ = RIGHT$(ToPrint$, LEN(ToPrint$) - ErrCode)
      END IF
    END IF
    IF CheckUserAbort THEN
      RetCode = -1
      GOTO AbortExit
    END IF
  LOOP
  FClose Handle
  CLOSE #1
  BlockClear
  DisplayMiscScrn PrintDone
  WaitForAction
  GOTO OverAbort

AbortExit:
  FClose Handle
  CLOSE #1

OverAbort:
  ON LOCAL ERROR GOTO 0
  EXIT SUB

ShowFileOnScreen:

  DIM FVI0 AS FVInfo            'create the TYPE variable

  REDIM Array(1 TO 16384)         'set up the MANDATORY 32K buffer

  FVI0.Colr = 15     'select a display color
  FVI0.ULRow = 2                 'define the window corners
  FVI0.ULCol = 1
  FVI0.LRRow = 24
  FVI0.LRCol = 80
  FVI0.TabStop = 11
  TitleRow = 1
  TitleCol = 40 - (LEN(RptTitle$) \ 2) + 1

  BlockClear

  DisplayMiscScrn FileViewWindow
  QPrintRC RptTitle$, TitleRow, TitleCol, 123
  Ky = 0                                'no keystrokes pending

  Action = 1                            'tell FileView to assume full control
  ActMRow = 25

  DO
    FileView FileName$, Ky, Action, FVI0, Array(1)
    '*** 09-22-94
    'Added Right moose button as the ESC key
    'for mouse handling consistance
    ButtonPress 1, N, MooseButton, MCol, MRow  ' ----- Check Mouse activity
    IF N AND 2 THEN                'if it was the right moose button and
      DO                           'if they are still holding it down then
        GetCursor 0, 0, NewStatus  'wait till they let go of the button
      LOOP WHILE NewStatus
      ButtonPress 0, 0, 0, 0, 0    'this has the effect of clearing all
      Ky = EscKey                  'moose buttons.
    END IF

    IF MooseButton THEN
      MRow = (MRow \ 8) + 1     'Convert MCol, MRow to Row and Col.
      MCol = (MCol \ 8) + 1
      IF MRow = ActMRow THEN
        SELECT CASE MCol
          CASE 2 TO 12           'F1 Help
            PressButton F1Key, ActMRow, 2, 12
          CASE 12 TO 23           'Escape  Menu
            PressButton EscKey, ActMRow, 12, 23
          CASE 23 TO 34           'F5 Print
            PressButton F5KEY, ActMRow, 23, 34
            DO
            LOOP WHILE LEN(INKEY$)        'clear key buffer
            Ky = EscKey
            PrintTheRpt = True
          CASE 44 TO 51           'Home
            PressButton HomeKey, ActMRow, 44, 51
          CASE 51 TO 57           'End
            PressButton EndKey, ActMRow, 51, 57
          CASE 57 TO 61           'Tab
            PressButton TabKey, ActMRow, 57, 61
          CASE 61 TO 65           'Shift Tab
            PressButton ShftTabKey, ActMRow, 61, 65
          CASE 65 TO 72           'Page Down
            PressButton PgUpKey, ActMRow, 65, 72
          CASE 72 TO 79           'Page Up
            PressButton PgDnKey, ActMRow, 72, 79
        END SELECT
      END IF
    END IF

    SELECT CASE Ky
      CASE F1Key
        SaveScrn TempScrn()
        DisplayMiscScrn "FVIEWHLP"
        WaitForAction
        RestScrn TempScrn()
      CASE F5KEY
        Ky = EscKey
        PrintTheRpt = True
      CASE ShftTabKey
        FVI0.HorizOffset = FVI0.HorizOffset - FVI0.TabStop
        IF FVI0.HorizOffset < 0 THEN FVI0.HorizOffset = 0
        Action = 2
      CASE TabKey
        FVI0.HorizOffset = FVI0.HorizOffset + FVI0.TabStop
        IF FVI0.HorizOffset >= 150 THEN FVI0.HorizOffset = 11
        IF FVI0.HorizOffset > 132 THEN FVI0.HorizOffset = 132
        Action = 2
    END SELECT
  LOOP UNTIL Ky = EscKey

  Action = 5 'tell fileview to close and clean itself up.

  FileView FileName$, Ky, Action, FVI0, Array(1)

  ERASE Array

  IF PrintTheRpt THEN
    EntryPoint = 0
    GOTO PrintFile
  END IF


END SUB

SUB RestScrn (ScrnArray())
  MScrnRest 1, 1, 25, 80, ScrnArray(0)
  'HideCursor
  'ExplodBox 1, 1, 25, 80, ScrnArray(), 0
  'ShowCursor
END SUB

SUB SaveScrn (ScrnArray())
 REDIM ScrnArray(2000)
 MScrnSave 1, 1, 25, 80, ScrnArray(0)
END SUB

SUB WaitForAction STATIC
  DO
    ButtonPress 1, Status, 0, 0, 0
'    call QPrintRC(STR$(FRE("")), 25, 1, 31)
  LOOP UNTIL LEN(INKEY$) OR Status      'then loop while status
  IF Status THEN
    DO
      GetCursor 0, 0, Status
    LOOP WHILE Status
    ButtonPress 0, 0, 0, 0, 0
  END IF
END SUB

