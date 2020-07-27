DEFINT A-Z
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB BlockClear ()
DECLARE FUNCTION GetAdjFactor# ()
DECLARE FUNCTION GetBillBook% ()
DECLARE FUNCTION GetBillCycle% ()

'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'qscr.BI'

  CONST False = 0, True = NOT False

FUNCTION GetAdjFactor#
  
  LibName$ = "UB"
  ScrnName$ = "FUELADJ"
  
  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  Action = 1
  Frm(1).StayOnField = True
  
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key, 13
      AdjFactor# = VAL(Form$(1, 0))
      ExitFlag = True
    CASE EscKey
      AdjFactor# = -10000
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 15   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 28 TO 40           '--Cancel button
          PressButton 27, 15, 28, 40
        CASE 41 TO 53           '--F5 button
          PressButton 13, 15, 41, 53
        END SELECT
      END SELECT                'row
    END IF
    
  LOOP UNTIL ExitFlag
  
  GetAdjFactor# = AdjFactor#
  ERASE Frm, Form$, Fld
  
END FUNCTION

FUNCTION GetBillBook%

  LibName$ = "UB"
  ScrnName$ = "BOOKGET"
  
  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  Action = 1
  Frm(1).StayOnField = True
  
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key, 13
      BookNum = VAL(Form$(1, 0))
      ExitFlag = True
    CASE EscKey
      BookNum = -10000
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 15   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 28 TO 40           '--Cancel button
          PressButton 27, 15, 28, 40
        CASE 41 TO 53           '--F5 button
          PressButton 13, 15, 41, 53
        END SELECT
      END SELECT                'row
    END IF
    
  LOOP UNTIL ExitFlag
  
  IF BookNum = 99 THEN
    BookNum = -1
  END IF
  
  GetBillBook% = BookNum
  ERASE Frm, Form$, Fld
  
END FUNCTION

FUNCTION GetBillCycle%

  LibName$ = "UB"
  ScrnName$ = "CYCLEGET"
  
  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  Action = 1
  Frm(1).StayOnField = True
  
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key, 13
      CycleNum = VAL(Form$(1, 0))
      ExitFlag = True
    CASE EscKey
      CycleNum = -10000
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 15   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 28 TO 40           '--Cancel button
          PressButton 27, 15, 28, 40
        CASE 41 TO 53           '--F5 button
          PressButton 13, 15, 41, 53
        END SELECT
      END SELECT                'row
    END IF
    
  LOOP UNTIL ExitFlag
  
  GetBillCycle% = CycleNum
  ERASE Frm, Form$, Fld
  
  
END FUNCTION

