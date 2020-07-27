  DEFINT A-Z
'$INCLUDE: 'DefCnf.BI'
DECLARE FUNCTION ubGetAcctTitle$ (RecordNumber%)
DECLARE FUNCTION ubAcctLookUp% (Acct$)
DECLARE FUNCTION ubFindAcct% (AcctNum$)
DECLARE FUNCTION ubFmtAcct$ (AN$, FundLen%, AcctLen%, DetLen%)
DECLARE FUNCTION ubStripAcct$ (AcctNum$)
DECLARE FUNCTION QPTrim$ (x$)
DECLARE SUB ubGetAcctStruct (FundLen%, AcctLen%, DetLen%)
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB RestScrn (ScrnArray())
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB AcctPickList (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB ubOpenAcctFile (AcctFileNum%, NumAccts%)
DECLARE SUB ubOpenAcctIdx (AcctIdxFileNum%, NumIdxRecs%)
DECLARE SUB GetCursor (x, y, Button)
DECLARE SUB APrintT0 (SEG ArrayElement AS ANY, ElSize, NumEls, MidStart, NumChars, Colr)
DECLARE SUB ButtonPress (Button, Status, Count, x, y)
DECLARE SUB HideCursor ()
DECLARE SUB MQPrint (x$, Colr)
DECLARE SUB MPaintBox (ULRow, ULCol, LRRow, LRCol, Colr)
DECLARE SUB MScrnSave (ULRow, ULCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnRest (ULRow, ULCol, LRRow, LRCol, SEG Address)
DECLARE SUB MouseTrap (ULRow, ULCol, LRRow, LRCol)
DECLARE SUB GetVMode (Mode, Page, PageSize, Rows, Columns)
DECLARE SUB QPrintRC (x$, Row, Col, Colr)
DECLARE SUB ShowCursor ()
DECLARE FUNCTION DOSVer% ()
DECLARE FUNCTION FindLast% (BYVAL Address, Size)
DECLARE FUNCTION MinInt% (Val1, Val2)
DECLARE FUNCTION PDQTimer& ()
DECLARE FUNCTION Null% (Strng$)
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'GL.BI'
  
  TYPE FLen
    V AS STRING * 42
  END TYPE

'****************************************************************************
' Description: Used to display a pick and account
'              From the GL accounts.
'
' Parameters:
'
'   Choice$() -
'     The "Choice$" Array holds the menu Choices.
'
'   Choice -
'     Current Menu Choice.  A value on entry will cause that choice to be
'     displayed & hilighted.
'
'   MaxLen -
'     Length of the incoming Choice strings
'
'   BoxBot -
'     Row number of the bottom of the menu box.
'
'   Ky$ -
'     This string will contain the character or key code that was used to
'     exit the routine.
'
'   Action -
'     Action is a flag which can be used to control the routine in a
'     multi tasking (polled) mode.  The Action commands are as follows:
'       0 - Routine will take control until a choice is selected or the
'           [Esc] key is pressed.  (Single Tasking Mode)
'       1 - Used on initial entry for multi-tasking mode.  Saves the
'           background screen, displays the Menu.  Resets itself to
'           Action = 3.
'       2 - Redisplay the menu contents only.  Reset Action to 3
'       3 - On entry means just check the keyboard and handle keys used by
'           the routine and exit back to the caller.  On exit means nothing
'           was selected.
'       4 - User made a selection or pressed Escape.
'           routine.
'       5 - Restore the screen, erase the screen buffer used by the routine.
'
' Note: Requires a structure named - "Cnf" defined as "Config".  This
'       structure is included in "DefCnf.Bi"
'
' External Routines:  MScrnSave, MScrnRest, MPaintBox, MQPrint, Mouse
'************************************************************************
SUB AcctPickList (Items() AS FLen, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS Config)
  
  IF Action = 4 THEN Action = 1 'See if they forgot to reset
  IF Action < 2 THEN REDIM ScrBuf(1)            'Temp Dim the Screen buffer
  IF Action = 5 AND UBOUND(ScrBuf) > 1 THEN     'Restore the screen if action=5
    MScrnRest Top - 1, LeftMargin, Bottom + 2, RightMargin + 2, ScrBuf(0)
    ERASE ScrBuf
    EXIT SUB
  END IF
  
  IF Action < 2 THEN
    GetVMode 0, 0, 0, ScRows, 0
    MouseTrap 1, 1, ScRows, 80
    NPixLines = 8
    IF ScRows > 43 AND DOSVer% >= 1000 THEN NPixLines = 7
    
    IF BoxBot = 0 THEN BoxBot = ScRows - 1
    BoxBot = MinInt(BoxBot, ScRows - 1)
    
    'Find the Inverse box color for the Scroll Bar
    
    ScrlClr = 112
    
    Fill$ = SPACE$(MaxLen + 1)  'used in Dsply to pad spaces
    
    '---- determine how many choices to display, and where to put the menu
    Num = UBOUND(Items)         'find number of active items
    FOR MaxNum = Num TO 1 STEP -1
      IF NOT Null(Items(MaxNum).V) THEN EXIT FOR
    NEXT
    
    Top = CSRLIN + 1            'put the menu at the current
    LeftMargin = POS(0)         '  cursor location
    Bottom = Top + MaxNum - 1
    More = 0
    LastRow = 0
    
    IF Bottom >= BoxBot THEN    '<---/-- BoxBot sets lowest row to display
      More = -1 '   /
      Bottom = BoxBot - 1       '<-/
    END IF
    WinSize = Bottom - Top + 1
    RightMargin = LeftMargin + MaxLen + 3
    
    '----- Save the underlying screen area
    Size = (WinSize + 3) * (RightMargin - LeftMargin + 3)
    REDIM ScrBuf(Size + 1)
    MScrnSave Top - 1, LeftMargin, Bottom + 2, RightMargin + 2, ScrBuf(0)
    
    '----- Make sure the choice # is in range
    Offset = 0
    IF Choice = 0 OR Choice > MaxNum THEN
      Choice = 1
    END IF
    
    IF Choice > Offset + WinSize THEN
      Offset = Choice - WinSize
    END IF
    
    '----- draw border and display items on the screen
    HideCursor
    QPrintRC "É" + STRING$(MaxLen + 2, "Í") + "»", CSRLIN, LeftMargin, Cnf.MenBox
    FOR x = Top TO Bottom       'Draw sides, display items
      APos = x - Top + 1
      
      IF Items(APos + Offset).V = "-" THEN
        Items(APos + Offset).V = STRING$(MaxLen, "Ä")
      END IF
      LSET Fill$ = Items(APos + Offset).V
      'Left box side and choice
      QPrintRC "º ", x, LeftMargin, Cnf.MenBox
      QPrintRC Fill$, x, LeftMargin + 2, Cnf.ActivCh
      IF x = Top AND More = -1 THEN             'Draw Up arrow box side
        QPrintRC CHR$(24), x, RightMargin, ScrlClr
      ELSEIF x = Bottom AND More = -1 THEN      'Draw Down arrow box side
        QPrintRC CHR$(25), x, RightMargin, ScrlClr
      ELSEIF More THEN          'Draw Scroll bar
        QPrintRC "²", x, RightMargin, 7         'ScrlClr
      ELSE      'Draw normal box side
        QPrintRC "º", x, RightMargin, Cnf.MenBox
      END IF
    NEXT
    'Draw the bottom of box
    QPrintRC "È" + STRING$(MaxLen + 2, "Í") + "¼", x, LeftMargin, Cnf.MenBox
    
    MPaintBox Top, RightMargin + 1, Bottom + 1, RightMargin + 2, 8
    MPaintBox Bottom + 2, LeftMargin + 2, Bottom + 2, RightMargin + 2, 8
    
    ShowCursor
    
    Fill$ = ""
    LCol = LeftMargin + 1       'Set variables for Hilite
    RCol = LCol + MaxLen + 1
    OPtr = 0
    GOSUB VertHiLite
  END IF
  
  IF Action = 2 THEN            'Redisplay menu contents
    GOSUB VertDisplay
    GOSUB VertHiLite
  END IF
  
  IF Action THEN Action = 3     'Reset Action = 3 (Idle state)
  
  
  DO
    
    DO          'Look for keys
      Ky$ = INKEY$
      
      ButtonPress 1, Down, Presses, x, y
      IF Presses = 0 THEN GetCursor x, y, Button
      MRow = (y \ NPixLines) + 1                'Convert X, Y to Row and Column
      MCol = (x \ 8) + 1
      
      '----- If User pressed a button inside the Box, set flag.
      IF Presses THEN
        Inside = 0
        IF MRow >= Top - 1 AND MRow <= Bottom + 1 THEN
          IF MCol > LeftMargin AND MCol < RightMargin - More THEN
            MouseTrap Top - 1, LeftMargin + 1, Bottom + 1, RightMargin - More - 1
            Inside = -1
          END IF
        END IF
      END IF
      
      'IF Down = 0 THEN LetGo2& = PDQTimer&
      
      IF (Down AND 2) THEN
        DO
          GetCursor 0, 0, NewStatus
        LOOP WHILE NewStatus
        ButtonPress 0, 0, 0, 0, 0
        Ky$ = CHR$(27)
        
        '----- Button pressed outside the box
      ELSEIF Presses AND NOT Inside THEN
        IF PutBack THEN
          PutBack = 0
        ELSEIF Presses AND LEN(MouseStor$) > 0 THEN
          HideCursor
          'MSetState MouseStor$
          PutBack = -1
          Presses = 0
        END IF
        
        '----- Handle only Mouse actions that were initiated inside the box
      ELSEIF Inside THEN
        IF Presses THEN WasPressed = -1
        
        '----- See if button pressed on the Scroll bar
        IF Presses AND More THEN
          IF MCol = RightMargin AND MRow > Top AND MRow < Bottom THEN
            'See if button pressed on the
            IF MRow = BarPtr THEN               '  Scroll Bar Pointer
              OnPtr = -1
              MouseTrap Top + 1, MCol, Bottom - 1, MCol
            END IF
            ScrlBar = -1        'Set Scroll Bar action flag
          END IF
        END IF
        
        '----- Handle Scroll Bar Mouse actions
        IF ScrlBar THEN
          IF Down = 0 THEN      'If button let go
            IF MRow > BarPtr THEN               'If Below the Scroll Pointer
              Ky$ = CHR$(0) + CHR$(81)          'Page Down
            ELSEIF MRow < BarPtr THEN           'Must be above Scroll Pointer
              Ky$ = CHR$(0) + CHR$(73)          'Page Up
            END IF
            ScrlBar = 0         'Action complete so reset flag
            OnPtr = 0
            WasPressed = 0
            MouseTrap 1, 1, ScRows, 80
            
          ELSEIF OnPtr THEN     'Move the Scroll Bar Pointer
            IF LMRow <> MRow THEN
              BarRow = MRow
              IF BarRow >= Bottom THEN BarRow = Bottom - 1
              IF BarRow <= Top THEN BarRow = Top + 1
              Choice = (BarRow - Top - 1&) * (MaxNum - 1) \ (WinSize - 3) + 1
              Offset = Choice - 1
              IF Offset + WinSize > MaxNum THEN Offset = MaxNum - WinSize
              GOSUB VertDisplay
              GOSUB VertHiLite
            END IF
          END IF
          
          '----- See if button pressed on Top or Bottom Arrow icons
        ELSEIF Down AND MCol = RightMargin AND More THEN
          IF PDQTimer& > LstScrl& + 1 THEN      'Delay between updates
            IF MRow = Top THEN  'If on top, do Up
              IF Choice > Offset + 1 THEN Choice = Offset + 2
              Ky$ = CHR$(0) + CHR$(72)          '  Arrow proc.
            ELSEIF MRow = Bottom THEN           'If on Bottom, do
              IF Choice < Offset + WinSize - 1 THEN Choice = Offset + WinSize - 1
              Ky$ = CHR$(0) + CHR$(80)          '  Down Arrow Proc.
            END IF
            LstScrl& = PDQTimer&                'Update Delay
            IF Presses THEN LstScrl& = PDQTimer& + 4            'If first press,
          END IF                'leave more time
          
          '----- Button pressed on Box Bottom or Top
        ELSEIF More AND Down = 1 AND (MRow < Top OR MRow > Bottom) THEN
          IF PDQTimer& > LstScrl& + 1 THEN      'Delay between updates
            IF MRow < Top THEN  'If higher than Top,
              Ky$ = CHR$(0) + CHR$(72)          '  Arrow proc.
            ELSEIF MRow > Bottom THEN           'If Lower than Bottom,
              Ky$ = CHR$(0) + CHR$(80)          '  Down Arrow Proc.
            END IF
            LstScrl& = PDQTimer&                'Update Delay
            IF Presses THEN LstScrl& = PDQTimer& + 4            'If first press,
          END IF                'leave more time
          
          '----- Button Pressed or Down on a Choice
        ELSEIF Down THEN
          IF MRow <> LMRow OR MCol <> LMCol OR Presses THEN
            IF MRow <= Bottom AND MRow >= Top THEN
              TChoice = MRow - Top + 1 + Offset 'Calc. Choice number
              IF LEFT$(Items(TChoice).V, 1) <> "Ä" THEN
                Choice = TChoice
                GOSUB VertHiLite                'Hi-Lite new choice
              END IF
            END IF
          END IF
          
          '----- If the button was let go
        ELSEIF WasPressed AND Down = 0 THEN
          IF MRow >= Top AND MRow <= Bottom THEN
            IF MCol > LeftMargin AND MCol < RightMargin THEN
              IF PDQTimer& < LetGo& + 10 THEN   'If half second since the last
                IF LEFT$(Items(FstPrsChoice).V, 1) <> CHR$(196) THEN
                  Ky$ = CHR$(13)                '  press/release, do Enter
                  Choice = FstPrsChoice         '  proc.
                END IF
              ELSE              'Start last-press timer
                LetGo& = PDQTimer&
                FstPrsChoice = Choice
              END IF
            END IF
          END IF
          WasPressed = 0
          MouseTrap 1, 1, ScRows, 80
        END IF
        
        IF Down = 0 THEN
          LMRow = -1
          LMCol = -1
        ELSE
          LMRow = MRow          'Remember were we were so we
          LMCol = MCol          '  can detect movement.
        END IF
        
      END IF
      'MMMMMMMMMMMMMMMMMMMMMM End of Mouse Code MMMMMMMMMMMMMMMMMMMMMMMM
      
    LOOP WHILE Ky$ = "" AND Action = 0
    
    KyLen = LEN(Ky$)
    
    IF KyLen = 1 THEN           'Non-extended keys
      
      SELECT CASE ASC(Ky$)
      CASE 13, 27               'Enter or Escape key
        IF Action THEN
          Action = 4
        ELSE
          MScrnRest Top - 1, LeftMargin, Bottom + 2, RightMargin + 2, ScrBuf(0)
          ERASE ScrBuf          'Clean up memory
        END IF
        LOCATE , , 1
        EXIT SUB
        
      CASE 33 TO 122            'Letter key pressed
        'Reset color
        Ky = ASC(UCASE$(Ky$))   'Upper case version of letter
        K2 = ASC(LCASE$(Ky$))   'dito for lower case
        Pass = -1               'First pass through choices
        N = Choice              'Start at current choice + 1
        DO
          Pass = Pass + 1
          IF N < MaxNum THEN
            N = N + 1
          ELSE N = 1
          END IF
          
          DO WHILE N <= MaxNum
            A = ASC(LTRIM$(Items(N).V))
            IF A = 91 THEN A = ASC(MID$(LTRIM$(Items(N).V), 3, 1))
            IF Ky = A OR K2 = A THEN EXIT DO
            N = N + 1
          LOOP
          IF Ky = A OR K2 = A THEN
            Choice = N
            IF Choice > Offset + (Bottom - Top + 1) THEN
              Offset = Choice - (Bottom - Top + 1)
              GOSUB VertDisplay
            END IF
            IF Choice <= Offset THEN
              Offset = Choice - 1
              GOSUB VertDisplay
            END IF
            EXIT DO
          END IF
        LOOP WHILE Choice > 1 AND Pass = 0
        GOSUB VertHiLite
        
      CASE ELSE
      END SELECT
      
    ELSEIF KyLen = 2 THEN       'Reset color
      
      SELECT CASE ASC(RIGHT$(Ky$, 1))
        
      CASE 72
        IF Choice > 1 THEN
          DO
            Choice = Choice - 1
            IF Row - 1 < Top THEN Offset = Offset - 1
          LOOP WHILE LEFT$(Items(Choice).V, 1) = "Ä"
          GOSUB VertDisplay
        ELSEIF More = 0 THEN
          Choice = MaxNum
        END IF
        
      CASE 80
        IF Choice < MaxNum THEN
          DO
            Choice = Choice + 1
            IF Row + 1 > Bottom THEN Offset = Offset + 1
          LOOP WHILE LEFT$(Items(Choice).V, 1) = "Ä"
          GOSUB VertDisplay
        ELSEIF More = 0 THEN
          Choice = 1
        END IF
        
      CASE 73
        IF Offset > 0 THEN
          IF Offset - WinSize > 0 THEN
            Offset = Offset - WinSize
          ELSE Offset = 0
          END IF
          Choice = (Row - Top + 1) + Offset
          GOSUB VertDisplay
        END IF
        
      CASE 81
        IF Offset + 2 * WinSize < MaxNum THEN
          Offset = Offset + WinSize
        ELSE Offset = MaxNum - WinSize
        END IF
        Choice = (Row - Top + 1) + Offset
        GOSUB VertDisplay
        
      CASE 71
        Choice = 1
        Offset = 0
        GOSUB VertDisplay
        
      CASE 79
        Choice = MaxNum
        Offset = MaxNum - (Bottom - Top + 1)
        GOSUB VertDisplay
        
      CASE ELSE
        
      END SELECT
      GOSUB VertHiLite
    END IF
    
  LOOP WHILE Action = 0         'get again if not multi-tasking
  
  MouseStor$ = ""
  
  EXIT SUB
  
  
VertHiLite:
  Row = Top + (Choice - Offset) - 1
  IF LastRow > 0 AND Row <> LastRow THEN MPaintBox LastRow, LCol, LastRow, RCol, Cnf.ActivCh
  MPaintBox Row, LCol, Row, RCol, Cnf.HiLite    'set inverse color
  LastRow = Row
  
  IF More THEN
    
    BarPtr = Choice * (WinSize - 3&) \ (MaxNum - 1) + Top + 1
    
    IF BarPtr <> OPtr THEN
      IF OPtr THEN
        LOCATE OPtr, RightMargin                'Update the Scroll Bar position
        MQPrint "²", ScrlClr
      END IF
      LOCATE BarPtr, RightMargin
      MQPrint CHR$(8), ScrlClr  'Cnf.MenBox
    END IF
    OPtr = BarPtr
  END IF
  
  LOCATE Row, LCol + 1
  
  RETURN
  
  '----- Print the choices in the window
VertDisplay:
  IF Choice <= Offset THEN Offset = Choice - 1
  IF Choice > Offset + WinSize THEN
    Offset = Choice - WinSize
  END IF
  
  LOCATE Top, LeftMargin + 2
  
  ElSize = LEN(Items(1).V)
  
  APrintT0 Items(Offset + 1), ElSize, Bottom - Top + 1, 1, MaxLen + 1, -1
  RETURN
  
END SUB

FUNCTION ubAcctLookUp% (Acct$)
  
  REDIM TempAcctIdx(1) AS GLAcctIndexType
  REDIM GLAcct(1)  AS GLAcctRecType
  REDIM ScrBuf(0)
  SaveScrn ScrBuf()
  
  MatchAccts = 0
  Lookup$ = QPTrim$(Acct$)
  Lookuplen = LEN(Lookup$)
  
  ubGetAcctStruct GLFundLen, GLAcctLen, GLDetLen
  TotLen = GLFundLen + GLAcctLen + GLDetLen

  IF (TotLen <= 0) OR (Lookuplen <= 0) THEN
    GOTO ExitLookUp
  END IF

  
  ubOpenAcctFile AcctFileNum, NumAccts
  ubOpenAcctIdx AcctIdxFileNum, NumIdxRecs

  IF (NumAccts <= 0) OR (NumIdxRecs <= 0) THEN
    CLOSE AcctFileNum, AcctIdxFileNum
    GOTO ExitLookUp
  END IF

  REDIM AcctList(1 TO NumIdxRecs) AS FLen       'Dim array to hold trans list
  FOR Cnt = 1 TO NumIdxRecs     'Read the file
    GET AcctIdxFileNum, Cnt, TempAcctIdx(1)
    IF LEFT$(TempAcctIdx(1).AcctNum, Lookuplen) = Lookup$ THEN
      GET AcctFileNum, TempAcctIdx(1).RecNum, GLAcct(1)
      MatchAccts = MatchAccts + 1
      LSET AcctList(MatchAccts).V = List$
      AcctList(MatchAccts).V = TempAcctIdx(1).AcctNum
      MID$(AcctList(MatchAccts).V, TotLen + 4) = GLAcct(1).Title
      MID$(AcctList(MatchAccts).V, 41) = MKI$(TempAcctIdx(1).RecNum)
    END IF
  NEXT
  
  CLOSE AcctFileNum, AcctIdxFileNum
  
  SELECT CASE MatchAccts
  CASE 0
    Ok = MsgBox%("UB", "BADGLACT")
    Action = 1
  CASE 1
    AcctRecNum = CVI(RIGHT$(AcctList(1).V, 2))
    Choice = 1
  CASE ELSE
    Row = 5
    Col = 15
    LOCATE Row, Col
    Action = 0
    MaxLen = 39
    BoxBot = 18
    ListDescTxt$ = SPACE$(MaxLen + 4)
    LSET ListDescTxt$ = "  Account No.  Description"
    QPrintRC ListDescTxt$, Row - 1, Col, 112       'Cnf.ActivCh
    DO          'call vertmenu to display matching recs to choose from
      AcctPickList AcctList(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        EXIT DO
      END IF
      IF Choice > 0 THEN
        AcctRecNum = CVI(RIGHT$(AcctList(Choice).V, 2))
        EXIT DO
      END IF
    LOOP
  END SELECT
  
  IF AcctRecNum > 0 THEN
    Acct$ = LEFT$(QPTrim$(AcctList(Choice).V), TotLen + 2)
    ubAcctLookUp% = AcctRecNum
  ELSE
    ubAcctLookUp% = 0
  END IF
  
  RestScrn ScrBuf()
  ERASE AcctList

ExitLookUp:
  ERASE ScrBuf, TempAcctIdx, GLAcct

END FUNCTION

'*****************************************************************************
'Searches the acct index for a matching account number and returns the record
'number of that account.
'
'    Input: AcctNum$ as a formatted G/L account number string
'  Returns: Record number of the account
'*****************************************************************************
FUNCTION ubFindAcct (AcctNum$)
  
  REDIM AcctIdx(1) AS GLAcctIndexType
  
  AcctIdxLen = LEN(AcctIdx(1))
  AcctIdxFileNum = FREEFILE
  OPEN "GLACCT.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS AcctIdxFileNum LEN = AcctIdxLen
  NumIdxRecs = LOF(AcctIdxFileNum) \ AcctIdxLen
  
  IF NumIdxRecs = 0 THEN
    RecordNum = 0
    CLOSE AcctIdxFileNum
    EXIT FUNCTION
  END IF
  
  Match = False
  FirstRec = 1
  LastRec = NumIdxRecs
  Lookfor$ = QPTrim$(AcctNum$)
  
  DO UNTIL LastRec < FirstRec
    
    MiddleRec = (LastRec + FirstRec) \ 2
    
    GET AcctIdxFileNum, MiddleRec, AcctIdx(1)
    ThisAcct$ = QPTrim$(AcctIdx(1).AcctNum)
    
    IF Lookfor$ = ThisAcct$ THEN
      Match = -1
      EXIT DO
    ELSEIF Lookfor$ < ThisAcct$ THEN
      LastRec = MiddleRec - 1
    ELSE
      FirstRec = MiddleRec + 1
    END IF
    
  LOOP
  
  IF Match THEN
    RecNum = AcctIdx(1).RecNum
  ELSE
    RecNum = 0
  END IF
  
  ubFindAcct = RecNum
  
  CLOSE AcctIdxFileNum
  
END FUNCTION

'***********************************************
'formats an account number string with dashes. *
'***********************************************
FUNCTION ubFmtAcct$ (AN$, FundLen%, AcctLen%, DetLen%) STATIC
  
  FmtTotAcctLen = FundLen + AcctLen + DetLen
  
  AN$ = QPTrim$(AN$)
  ANLen = LEN(AN$)
  
  IF ANLen > FmtTotAcctLen THEN
    AN$ = LEFT$(AN$, FmtTotAcctLen)
    ANLen = FmtTotAcctLen
  END IF
  
  SELECT CASE ANLen
  CASE IS < FundLen
    ubFmtAcct$ = AN$
  CASE FundLen
    ubFmtAcct$ = AN$ + "-"
  CASE (FundLen + 1) TO (AcctLen + FundLen) - 1
    ubFmtAcct$ = LEFT$(AN$, FundLen) + "-" + MID$(AN$, FundLen + 1)
  CASE (AcctLen + FundLen)
    ubFmtAcct$ = LEFT$(AN$, FundLen) + "-" + MID$(AN$, FundLen + 1, AcctLen) + "-"
  CASE (AcctLen + FundLen + 1) TO (AcctLen + FundLen + DetLen) - 1
    ubFmtAcct$ = LEFT$(AN$, FundLen) + "-" + MID$(AN$, FundLen + 1, AcctLen) + "-" + MID$(AN$, FundLen + AcctLen + 1)
  CASE (AcctLen + FundLen + DetLen)
    ubFmtAcct$ = LEFT$(AN$, FundLen) + "-" + MID$(AN$, FundLen + 1, AcctLen) + "-" + MID$(AN$, FundLen + AcctLen + 1, DetLen)   'RIGHT$(AN$, DetLen)
  END SELECT
END FUNCTION

SUB ubGetAcctStruct (GLFundLen%, GLAcctLen%, GLDetLen%)
  
  REDIM GLSetUpRec(1) AS GLSetupRecType
  SetUpRecLen = LEN(GLSetUpRec(1))
  
  SetUpFile = FREEFILE
  OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetUpFile LEN = SetUpRecLen

  IF LOF(SetUpFile) = 0 THEN
  GLFundLen = 0
  GLAcctLen = 0
  GLDetLen = 0
  CLOSE
  EXIT SUB
  END IF

  GET SetUpFile, 1, GLSetUpRec(1)
  
  GLFundLen = GLSetUpRec(1).FundLen
  GLAcctLen = GLSetUpRec(1).AcctLen
  GLDetLen = GLSetUpRec(1).DetLen
  
  CLOSE SetUpFile
  ERASE GLSetUpRec
END SUB

FUNCTION ubGetAcctTitle$ (RecordNumber)
   REDIM TempAcct(1)  AS GLAcctRecType
   ubOpenAcctFile AcctFileNum, NumAccts
   GET AcctFileNum, RecordNumber, TempAcct(1)
   ubGetAcctTitle$ = TempAcct(1).Title
   CLOSE AcctFileNum
   ERASE TempAcct
END FUNCTION

SUB ubOpenAcctFile (AcctFileNum, NumAccts)
  REDIM TempAcct(1)  AS GLAcctRecType
  AcctRecLen = LEN(TempAcct(1))
  AcctFileNum = FREEFILE
  OPEN "GLACCT.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS AcctFileNum LEN = AcctRecLen
  NumAccts = LOF(AcctFileNum) \ AcctRecLen
  ERASE TempAcct
END SUB

SUB ubOpenAcctIdx (AcctIdxFileNum, NumIdxRecs)
  REDIM TempAcctIdx(1) AS GLAcctIndexType
  AcctIdxLen = LEN(TempAcctIdx(1))
  AcctIdxFileNum = FREEFILE
  OPEN "GLACCT.IDX" FOR RANDOM ACCESS READ WRITE SHARED AS AcctIdxFileNum LEN = AcctIdxLen
  NumIdxRecs = LOF(AcctIdxFileNum) \ AcctIdxLen
  ERASE TempAcctIdx
END SUB

FUNCTION ubStripAcct$ (AcctNum$) STATIC
  x$ = QPTrim$(AcctNum$)
  DO
    DashPos = INSTR(x$, "-")
    IF DashPos > 0 THEN
      x$ = LEFT$(x$, DashPos - 1) + MID$(x$, DashPos + 1)
    END IF
  LOOP WHILE DashPos
  ubStripAcct$ = x$
END FUNCTION

