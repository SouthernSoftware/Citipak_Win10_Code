DEFINT A-Z
DECLARE SUB CFEditBankDraftRec ()
DECLARE SUB LoadSysDraftRec (Form$())
DECLARE SUB SaveSysDraftRec (Form$())
DECLARE SUB EditBankDraftRec ()
DECLARE FUNCTION AskAddNewTbl% ()
  
DECLARE SUB KillFile (FileName$)
DECLARE SUB CFNewYearProcess ()
DECLARE SUB CFPrinterMaint ()
DECLARE SUB CFZeroEmpYTD ()
  
DECLARE SUB PRCtrlMenu ()
  
DECLARE FUNCTION NotFromPRMain% ()
DECLARE SUB SaveField (FldNo%, Form$(), Fld() AS ANY, BadFld%)
DECLARE SUB CFLeaveFileMaint ()
DECLARE SUB CFRetireFileMaint ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB WaitForAction ()
DECLARE SUB FCreate (FileName$)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutRT (Handle%, Source AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Source AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutT (Handle%, Source AS ANY, NumBytes%)
  
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRT (Handle%, Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetT (Handle%, Dest AS ANY, NumBytes%)
  
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB FSeek (Handle%, Location&)
  
DECLARE SUB FFlush (Handle%)
  
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FEof% (Handle%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FLoc& (Handle%)
DECLARE FUNCTION FLof& (Handle%)
DECLARE SUB InsertT (SEG StartElement AS ANY, ElSize%, NumEls%)
DECLARE FUNCTION Monitor ()
DECLARE SUB MScrnSave (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnRest (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB GetCursor (x, Y, Button)
DECLARE SUB MQPrint (x$, Colr)
DECLARE SUB MPaintBox (UlRow, UlCol, LRRow, LRCol, Colr)
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, x%, Y%)
DECLARE SUB HideCursor ()
DECLARE SUB TextCursor (Fg, Bg)
DECLARE SUB ShowCursor ()
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE FUNCTION QPStrI$ (Num%)
  
DECLARE FUNCTION LibNumberOfFields (LibName$, FrmName$)
DECLARE FUNCTION LibScrName$ (Scr, ScrnLib())
DECLARE FUNCTION ScrnLibSize (LibName$)
DECLARE SUB EditForm (Form$(), Fld() AS ANY, Frm AS ANY, Cnf AS ANY, Action)
DECLARE SUB Lib2Scrn (NameInLib$, ScrnLib(), MonoCode, Attribute, ErrorCode)
DECLARE SUB LibGetFldDef (LibName$, FrmName$, StartEl, Fld() AS ANY, Form$(), ErrCode)
DECLARE SUB LoadScrnLib (LibName$, ScrnLib(), ErrorCode)
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintArray (FirstFld, LastFld, Form$(), Fld() AS ANY)
DECLARE SUB UnPackBuffer (FirstFld, LastFld, Form$(), Fld() AS ANY)
DECLARE FUNCTION FldNum (FldName$, Fld() AS ANY)
DECLARE FUNCTION EndOfForms (Fld() AS ANY)
DECLARE FUNCTION PDQTimer& ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB SearchT2 (SEG Array AS ANY, NumEls%, Match, Found%, Dir%, Code%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION QPRTrim$ (Text$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE SUB PayMenu (BYVAL MenuNum%, Choice%, NumOfItems%)
  
DECLARE SUB VertMenu (Item$(), Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB CursorOff ()
DECLARE SUB RestScrn (ScrnArray%())
DECLARE SUB SaveScrn (ScrnArray%())
DECLARE FUNCTION ASCII% (Strng$)
DECLARE SUB Pause (Ticks)
DECLARE SUB CFErnFileMaint ()
DECLARE SUB CFDedFileMaint ()
DECLARE SUB ClearScrn ()
DECLARE SUB BlockClear ()
DECLARE SUB CFStateTaxMaint ()
DECLARE SUB CFUnitFileMaint ()
DECLARE SUB CFSysFileMaint ()
DECLARE SUB CFEICTableMaint ()
DECLARE SUB CFFederalTaxMaint ()
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION QPValI% (Number$)
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB InitStateTaxDef ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)

  
  '$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.


  '$INCLUDE: 'FieldInf.bi'                '        "        field information
  '$INCLUDE: 'FormEdit.bi'                '        "        form editing info.
  '$INCLUDE: 'PRFiles.bi'
  '$INCLUDE: 'PREmpRec.bi'
  '$INCLUDE: 'PageInfo.bi'    '        "          Page information
  '$INCLUDE: 'Retire.Bi'
  '$INCLUDE: 'DedCodes.Bi'
  '$INCLUDE: 'ErnCodes.Bi'
  '$INCLUDE: 'Leave.Bi'
  '$INCLUDE: 'StateTax.Bi'
  '$INCLUDE: 'FedTax.Bi'
  '$INCLUDE: 'PREIC.Bi'
  '$INCLUDE: 'PRUNIT.Bi'
  '$INCLUDE: 'PRNSETUP.Bi'
  '$INCLUDE: 'PRSYSCTR.Bi'
  '$INCLUDE: 'PRdraft.BI'
  '$INCLUDE: 'setCnf.bi'                  'Defines TYPE for monitor/color info.
  ''$INCLUDE: 'PRTRANS.Bi'
  ''$INCLUDE: 'PRTIMEIN.Bi'

  CONST False = 0
  CONST True = NOT False

  PRCtrlMenu

FUNCTION AskAddNewTbl

  REDIM Frm(1) AS FormInfo

  REDIM TempScrn(0)
  SaveScrn TempScrn()

  FormName$ = "ADDLVTBL"
  NumFlds = LibNumberOfFields(MiscQLib, FormName$)
  REDIM Form$(NumFlds, 2)             'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo     'DIM the field information array
  StartEl = 0                         'Load first form at array start
  LibGetFldDef MiscQLib, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1

  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 2                 'Start editing on field #1
  Frm(1).InsStat = False           'Set insert state (True = Insert on)
  Frm(1).StartEl = 0               'Set form starting element to 0 and

  DisplayMiscScrn FormName$

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
      CASE F8KEY
        SaveFlag = True
        ExitFlag = True
      CASE EscKey
        ExitFlag = True
    END SELECT

  LOOP UNTIL ExitFlag    'proper key not set

  AskAddNewTbl = SaveFlag

  CursorOff

  RestScrn TempScrn()

  ERASE TempScrn, Form$, Fld, Frm

END FUNCTION

SUB CFDedFileMaint
  
  REDIM BlankDedRec(1)              AS DedCodeRecType
  REDIM Frm(1) AS FormInfo
  REDIM DedCodeDataRec(1 TO 12)    AS DedCodeRecType
  
  DedCodeRecLen = LEN(DedCodeDataRec(1)) * 12
  
  CursorOff
  
  'LibName$ = CtrlQLib     'Specify a Screen Library file name.
  FormName$ = "DEDCODES"
  LibSize = ScrnLibSize(CtrlQLib)               'Get scrn lib array size
  REDIM ScrnLib(LibSize)                        'Create scrn lib array
  LoadScrnLib CtrlQLib, ScrnLib(), ErrorCode    'load scrn lib
  
  NumFlds = LibNumberOfFields(CtrlQLib, FormName$)
  REDIM FormCFDC$(NumFlds, 2)                   'DIM the form data array
  REDIM FldCFDC(NumFlds) AS FieldInfo           'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef CtrlQLib, FormName$, StartEl, FldCFDC(), FormCFDC$(), ErrCode
  
  BlankDedRec(1).DCFWT1 = "N"
  BlankDedRec(1).DCSWT1 = "N"
  BlankDedRec(1).DCSOC1 = "N"
  BlankDedRec(1).DCMED1 = "N"
  
  FormCFDC$(0, 0) = SPACE$(DedCodeRecLen)
  FGetAH DedCodeFileName, DedCodeDataRec(1), LEN(DedCodeDataRec(1)), 12
  BCopy VARSEG(DedCodeDataRec(1)), VARPTR(DedCodeDataRec(1)), SSEG(FormCFDC$(0, 0)), SADD(FormCFDC$(0, 0)), DedCodeRecLen, 0
  UnPackBuffer 0, 0, FormCFDC$(), FldCFDC()
  
  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  ClearScrn
  DO            'Pole the editing procedure
    Lib2Scrn FormName$, ScrnLib(), MonoCode, -1, ErrCode
    
    DO
      EditForm FormCFDC$(), FldCFDC(), Frm(1), Cnf, Action
      
      IF Frm(1).Presses AND Frm(1).MRow = 25 THEN
        SELECT CASE Frm(1).MCol
        CASE 3 TO 12            'F1
          PressButton F1Key, 25, 3, 12
        CASE 14 TO 23           'F10
          PressButton F0Key, 25, 14, 23
        CASE 36 TO 45           'Escape
          PressButton EscKey, 25, 36, 45
'        CASE 58 TO 68           'F8
'          PressButton F8Key, 25, 58, 68
        END SELECT
      END IF
      
      SELECT CASE Frm(1).KeyCode
      CASE F0Key, EscKey
        SELECT CASE Frm(1).KeyCode
        CASE EscKey
          IF Frm(1).Edited THEN
            SaveFlag = PromptSaveData
          ELSE
            SaveFlag = False
          END IF
        CASE F0Key
          IF Frm(1).Edited THEN
            SaveFlag = True
          ELSE
            SaveFlag = False
          END IF
        END SELECT
        SELECT CASE SaveFlag
        CASE True
          GOSUB CFDedCodeRecPack
          FPutAH DedCodeFileName, DedCodeDataRec(1), LEN(DedCodeDataRec(1)), 12
          DisplayMiscScrn UpdatedOK
          WaitForAction
          ExitFlag = True
        CASE False
          ExitFlag = True
        CASE 1
          Action = 2
        END SELECT
'      CASE F8Key                'Erase rec
'        DescFld = FldNum("DCDESC" + RIGHT$(QPRTrim$(FldCFDC(Frm(1).FldNo).FldName), 1), FldCFDC())
'        LSET FormCFDC$(DescFld, 0) = ""
'        SaveField DescFld, FormCFDC$(), FldCFDC(), BadFld
'        GOSUB CFDedCodeRecPack
'        Frm(1).Edited = True
'        Frm(1).FldNo = DescFld
'        Action = 2
      END SELECT
      
    LOOP UNTIL ExitFlag = True
    
  LOOP UNTIL ExitFlag = True    '  the Escape key.
  
  CursorOff
  
  
  ERASE BlankDedRec, DedCodeDataRec, ScrnLib
  ERASE FormCFDC$, FldCFDC, Frm
  
  EXIT SUB
  
CFDedCodeRecPack:
  
  BCopy SSEG(FormCFDC$(0, 0)), SADD(FormCFDC$(0, 0)), VARSEG(DedCodeDataRec(1)), VARPTR(DedCodeDataRec(1)), LEN(FormCFDC$(0, 0)), 0
  FOR Cnt = 1 TO 12
    IF LEN(QPRTrim$(DedCodeDataRec(Cnt).DCDESC1)) = 0 THEN
      DedCodeDataRec(Cnt) = BlankDedRec(1)
      ThisRec = Cnt
      DO
        ThisRec = ThisRec + 1
        IF ThisRec = 13 THEN EXIT DO
        IF LEN(QPRTrim$(DedCodeDataRec(ThisRec).DCDESC1)) THEN
          SWAP DedCodeDataRec(ThisRec), DedCodeDataRec(Cnt)
          EXIT DO
        END IF
      LOOP
    END IF
  NEXT
  BCopy VARSEG(DedCodeDataRec(1)), VARPTR(DedCodeDataRec(1)), SSEG(FormCFDC$(0, 0)), SADD(FormCFDC$(0, 0)), DedCodeRecLen, 0
  UnPackBuffer 0, 0, FormCFDC$(), FldCFDC()
  RETURN
  
END SUB

SUB CFEditBankDraftRec
  
  'LibName$ =
  ScrnName$ = "PRDRAFT"
  NumScrns = 1
  
  '--define the multi-choice fields
  NumFlds = -1
  NumFlds = LibNumberOfFields(CtrlQLib, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef CtrlQLib, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  
  LoadSysDraftRec Form$()
  CALL UnPackBuffer(0, 0, Form$(), Fld())
  
  '--Set screen number to one and display screen
  Scr = 1
  ClearScrn
  LibFile2Scrn CtrlQLib, ScrnName$, MonoCode, Attribute%, ErrCode
  
  ShowCursor
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF Frm(1).Presses AND Frm(1).MRow = 25 THEN
      SELECT CASE Frm(1).MCol
      CASE 3 TO 12            'F1
        PressButton F1Key, 25, 3, 12
      CASE 14 TO 23           'F10
        PressButton F0Key, 25, 14, 23
      CASE 36 TO 45           'Escape
        PressButton EscKey, 25, 36, 45
      END SELECT
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      SaveSysDraftRec Form$()
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 21   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 43 TO 54           '--Save Button
          PressButton -68, 21, 43, 54
        CASE 62 TO 72           '--cancel button
          PressButton 27, 21, 57, 70
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
    
  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag
  
'  HideCursor
  
END SUB

SUB CFEICTableMaint
  
  REDIM EICRec(1)  AS EICRecType
  
  REDIM Frm(1)    AS FormInfo
  
  EICRecLen = LEN(EICRec(1))
  
  CursorOff
  
  REDIM TempScrn(0)
  
  NumOfScrns = 2
  F7Flag = False
  'LibName$ = CtrlQLib     'Specify a Screen Library file name.
  REDIM FormName$(1 TO NumOfScrns)
  FOR Scr = 1 TO NumOfScrns
    FormName$(Scr) = "EIC" + LTRIM$(STR$(Scr))
  NEXT
  LibSize = ScrnLibSize(CtrlQLib)               'Get scrn lib array size
  REDIM ScrnLib(LibSize)        'Create scrn lib array
  LoadScrnLib CtrlQLib, ScrnLib(), ErrorCode    'load scrn lib
  NumFlds = -1
  
  FOR Scr = 1 TO NumOfScrns     '-- Get total number of fields
    NumFlds = NumFlds + LibNumberOfFields(CtrlQLib, FormName$(Scr)) + 1
  NEXT
  
  REDIM FormCFEIC$(NumFlds, 2)  'DIM the form data array
  REDIM FldCFEIC(NumFlds) AS FieldInfo          'DIM the field information array
  REDIM PageInfo(1 TO NumOfScrns) AS PageInfoType
  
  StartEl = 0   'Load first form at array start
  FOR Scr = 1 TO NumOfScrns     'Get field defs
    LibGetFldDef CtrlQLib, FormName$(Scr), StartEl, FldCFEIC(), FormCFEIC$(), ErrCode
    PageInfo(Scr).PageNo = Scr
    PageInfo(Scr).FirstFld = StartEl + 1
    StartEl = StartEl + FldCFEIC(StartEl).Fields + 1            'Inc StartEl to next
    PageInfo(Scr).LastFld = StartEl - 1
  NEXT
  FOR n = 1 TO NumFlds          'For all fields on both pages
    LSET FormCFEIC$(n, 0) = ""  'Clear each field
  NEXT
  FormCFEIC$(0, 0) = SPACE$(EICRecLen)
  
  FGetAH EICFileName, EICRec(1), EICRecLen, 1
  BCopy VARSEG(EICRec(1)), VARPTR(EICRec(1)), SSEG(FormCFEIC$(0, 0)), SADD(FormCFEIC$(0, 0)), EICRecLen, 0
  UnPackBuffer 0, 0, FormCFEIC$(), FldCFEIC()
  
  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  Scr = 1       'Set screen number to one
  ClearScrn
  
  Lib2Scrn FormName$(Scr), ScrnLib(), MonoCode, -1, ErrCode
  'display first screen
  
  
  DO
    
    EditForm FormCFEIC$(), FldCFEIC(), Frm(1), Cnf, Action
    
    '*--------------------------------------------
    
    IF Frm(1).Presses AND Frm(1).MRow = 25 THEN
      SELECT CASE Frm(1).MCol
      CASE 3 TO 12              'F1
        PressButton F1Key, 25, 3, 12
      CASE 14 TO 23             'F2
        PressButton F0Key, 25, 14, 23
      CASE 36 TO 45             'Escape
        PressButton EscKey, 25, 36, 45
      CASE 58 TO 67             'PageDown
        IF Scr < NumOfScrns THEN
          PressButton PgDnKey, 25, 58, 67
        END IF
      CASE 69 TO 78             'PageUp
        IF Scr > 1 THEN
          PressButton PgUpKey, 25, 69, 78
        END IF
      END SELECT
    END IF
    
    
    '*************************************************************************
    '-------------------------------------------------------------------------
    SELECT CASE Frm(1).KeyCode
    CASE F0Key, EscKey
      
      
      SELECT CASE Frm(1).KeyCode
      CASE EscKey
        IF Frm(1).Edited THEN
          
          SaveFlag = PromptSaveData
        ELSE
          SaveFlag = False
        END IF
      CASE F0Key
        IF Frm(1).Edited THEN
          SaveFlag = True
        ELSE
          SaveFlag = False
        END IF
      END SELECT
      SELECT CASE SaveFlag
      CASE True
        'save true
        SaveScrn TempScrn()
        DisplayMiscScrn UpdatingDisk
        BCopy SSEG(FormCFEIC$(0, 0)), SADD(FormCFEIC$(0, 0)), VARSEG(EICRec(1)), VARPTR(EICRec(1)), EICRecLen, 0
        FPutAH EICFileName, EICRec(1), EICRecLen, 1
        DisplayMiscScrn UpdatedOK
        WaitForAction
        RestScrn TempScrn()
        ERASE TempScrn
        ExitFlag = True
      CASE False
        'save false
        ExitFlag = True
      CASE 1
        Action = 2
        ExitFlag = False
      END SELECT
    CASE ELSE
      
    END SELECT
    
    IF Frm(1).FldNo > PageInfo(Scr).LastFld OR Frm(1).FldNo < PageInfo(Scr).FirstFld THEN
      Action = 2
      SELECT CASE Frm(1).FldNo
      CASE IS < PageInfo(Scr).FirstFld
        IF Scr > 1 THEN Scr = Scr - 1
      CASE IS > PageInfo(Scr).LastFld
        IF Scr < NumOfScrns THEN Scr = Scr + 1
      END SELECT
      Lib2Scrn FormName$(Scr), ScrnLib(), MonoCode, -2, ErrCode
    END IF
    '*************************************************************************
    'Keep editing until the user presses
  LOOP UNTIL ExitFlag = True
  
  ERASE EICRec, TempScrn, FormName$, ScrnLib
  ERASE FormCFEIC$, FldCFEIC, PageInfo, Frm
  
  CursorOff
  
  
END SUB

SUB CFErnFileMaint
  
  REDIM BlankErnRec(1)   AS ErnCodeRecType
  REDIM Frm(1) AS FormInfo
  REDIM ErnCodeDataRec(1 TO 3)    AS ErnCodeRecType
  
  ErnRecLen = LEN(ErnCodeDataRec(1)) * 3
  
  REDIM TempScrn(0)
  
  'LibName$ = CtrlQLib     'Specify a Screen Library file name.
  FormName$ = "ERNCODES"
  LibSize = ScrnLibSize(CtrlQLib)               'Get scrn lib array size
  REDIM ScrnLib(LibSize)                        'Create scrn lib array
  LoadScrnLib CtrlQLib, ScrnLib(), ErrorCode    'load scrn lib
  NumFlds = LibNumberOfFields(CtrlQLib, FormName$)
  REDIM FormCFEC$(NumFlds, 2)                   'DIM the form data array
  REDIM FldCFEC(NumFlds) AS FieldInfo           'DIM the field information array
  StartEl = 0                                   'Load first form at array start
  LibGetFldDef CtrlQLib, FormName$, StartEl, FldCFEC(), FormCFEC$(), ErrCode
  
  BlankErnRec(1).ERNCODE1 = ""
  BlankErnRec(1).ERNFWT1 = "Y"
  BlankErnRec(1).ERNSWT1 = "Y"
  BlankErnRec(1).ERNSOC1 = "Y"
  BlankErnRec(1).ERNMED1 = "Y"
  BlankErnRec(1).ERNRET1 = "Y"
  
  FormCFEC$(0, 0) = SPACE$(ErnRecLen)
  FGetAH ErnCodeFileName, ErnCodeDataRec(1), LEN(ErnCodeDataRec(1)), 3
  BCopy VARSEG(ErnCodeDataRec(1)), VARPTR(ErnCodeDataRec(1)), SSEG(FormCFEC$(0, 0)), SADD(FormCFEC$(0, 0)), ErnRecLen, 0
  UnPackBuffer 0, 0, FormCFEC$(), FldCFEC()
  
  
  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set the insert state (0 = Insert off)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  ClearScrn
  
  DO            'Pole the editing procedure
    Lib2Scrn FormName$, ScrnLib(), MonoCode, -1, ErrCode
    
    DO
      EditForm FormCFEC$(), FldCFEC(), Frm(1), Cnf, Action
      
      IF Frm(1).Presses AND Frm(1).MRow = 25 THEN
        SELECT CASE Frm(1).MCol
        CASE 3 TO 12            'F1
          PressButton F1Key, 25, 3, 12
        CASE 14 TO 23           'F10
          PressButton F0Key, 25, 14, 23
        CASE 36 TO 45           'Escape
          PressButton EscKey, 25, 36, 45
        CASE 58 TO 68           'F8
          PressButton F8KEY, 25, 58, 68
        END SELECT
      END IF
      
      SELECT CASE Frm(1).KeyCode
      CASE F0Key, EscKey
        SELECT CASE Frm(1).KeyCode
        CASE EscKey
          IF Frm(1).Edited THEN
            SaveFlag = PromptSaveData
          ELSE
            SaveFlag = False
          END IF
        CASE F0Key
          IF Frm(1).Edited THEN
            SaveFlag = True
          ELSE
            SaveFlag = False
          END IF
        END SELECT
        SELECT CASE SaveFlag
        CASE True
          GOSUB ErnCodeRecPack
          FPutAH ErnCodeFileName, ErnCodeDataRec(1), LEN(ErnCodeDataRec(1)), 3
          DisplayMiscScrn UpdatedOK
          WaitForAction
          ExitFlag = True
        CASE False
          ExitFlag = True
        CASE 1
          Action = 2
        END SELECT
      CASE F8KEY                'Erase rec
        DescFld = FldNum("ERNCODE" + RIGHT$(QPRTrim$(FldCFEC(Frm(1).FldNo).FldName), 1), FldCFEC())
        LSET FormCFEC$(DescFld, 0) = ""
        SaveField DescFld, FormCFEC$(), FldCFEC(), BadFld
        GOSUB ErnCodeRecPack
        Frm(1).Edited = True
        Frm(1).FldNo = DescFld
        Action = 2
      END SELECT
      
    LOOP UNTIL ExitFlag = True
    
  LOOP UNTIL ExitFlag = True    '  the Escape key.
  
  CursorOff
  
  ERASE BlankErnRec, ErnCodeDataRec, TempScrn, ScrnLib
  ERASE FormCFEC$, FldCFEC, Frm
  
  EXIT SUB
  
ErnCodeRecPack:
  
  BCopy SSEG(FormCFEC$(0, 0)), SADD(FormCFEC$(0, 0)), VARSEG(ErnCodeDataRec(1)), VARPTR(ErnCodeDataRec(1)), LEN(FormCFEC$(0, 0)), 0
  FOR Cnt = 1 TO 3
    IF LEN(QPRTrim$(ErnCodeDataRec(Cnt).ERNCODE1)) = 0 THEN
      ErnCodeDataRec(Cnt) = BlankErnRec(1)
      ThisRec = Cnt
      DO
        ThisRec = ThisRec + 1
        IF ThisRec = 4 THEN EXIT DO
        IF LEN(QPRTrim$(ErnCodeDataRec(ThisRec).ERNCODE1)) THEN
          SWAP ErnCodeDataRec(ThisRec), ErnCodeDataRec(Cnt)
          EXIT DO
        END IF
      LOOP
    END IF
  NEXT
  BCopy VARSEG(ErnCodeDataRec(1)), VARPTR(ErnCodeDataRec(1)), SSEG(FormCFEC$(0, 0)), SADD(FormCFEC$(0, 0)), ErnRecLen, 0
  UnPackBuffer 0, 0, FormCFEC$(), FldCFEC()
  RETURN
  
END SUB

SUB CFFederalTaxMaint
  
  REDIM Frm(1)    AS FormInfo
  
  REDIM FedTaxRec(1)  AS FederalTaxRecType
  
  FederalTaxRecLen = LEN(FedTaxRec(1))
  
  CursorOff
  
  REDIM TempScrn(0)
  
  NumOfScrns = 2
  F7Flag = False
  'LibName$ = CtrlQLib     'Specify a Screen Library file name.
  REDIM FormName$(1 TO NumOfScrns)
  FOR Scr = 1 TO NumOfScrns
    FormName$(Scr) = "FDTAX" + LTRIM$(STR$(Scr))
  NEXT
  LibSize = ScrnLibSize(CtrlQLib)               'Get scrn lib array size
  REDIM ScrnLib(LibSize)        'Create scrn lib array
  LoadScrnLib CtrlQLib, ScrnLib(), ErrorCode    'load scrn lib
  NumFlds = -1
  
  FOR Scr = 1 TO NumOfScrns     '-- Get total number of fields
    NumFlds = NumFlds + LibNumberOfFields(CtrlQLib, FormName$(Scr)) + 1
  NEXT
  
  REDIM FormCFFT$(NumFlds, 2)   'DIM the form data array
  REDIM FldCFFT(NumFlds) AS FieldInfo           'DIM the field information array
  REDIM PageInfo(1 TO NumOfScrns) AS PageInfoType
  
  StartEl = 0   'Load first form at array start
  FOR Scr = 1 TO NumOfScrns     'Get field defs
    LibGetFldDef CtrlQLib, FormName$(Scr), StartEl, FldCFFT(), FormCFFT$(), ErrCode
    PageInfo(Scr).PageNo = Scr
    PageInfo(Scr).FirstFld = StartEl + 1
    StartEl = StartEl + FldCFFT(StartEl).Fields + 1             'Inc StartEl to next
    PageInfo(Scr).LastFld = StartEl - 1
  NEXT
  FOR n = 1 TO NumFlds          'For all fields on both pages
    LSET FormCFFT$(n, 0) = ""   'Clear each field
  NEXT
  
  FormCFFT$(0, 0) = SPACE$(FederalTaxRecLen)
  
  FGetAH FederalTaxFileName, FedTaxRec(1), FederalTaxRecLen, 1
  BCopy VARSEG(FedTaxRec(1)), VARPTR(FedTaxRec(1)), SSEG(FormCFFT$(0, 0)), SADD(FormCFFT$(0, 0)), FederalTaxRecLen, 0
  UnPackBuffer 0, 0, FormCFFT$(), FldCFFT()
  
  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  Scr = 1       'Set screen number to one
  ClearScrn
  
  Lib2Scrn FormName$(Scr), ScrnLib(), MonoCode, -1, ErrCode
  'display first screen
  
  
  DO
    
    EditForm FormCFFT$(), FldCFFT(), Frm(1), Cnf, Action
    
    '*--------------------------------------------
    
    IF Frm(1).Presses AND Frm(1).MRow = 25 THEN
      SELECT CASE Frm(1).MCol
      CASE 3 TO 12              'F1
        PressButton F1Key, 25, 3, 12
      CASE 14 TO 23             'F2
        PressButton F0Key, 25, 14, 23
      CASE 36 TO 45             'Escape
        PressButton EscKey, 25, 36, 45
      CASE 58 TO 67             'PageDown
        IF Scr < NumOfScrns THEN
          PressButton PgDnKey, 25, 58, 67
        END IF
      CASE 69 TO 78             'PageUp
        IF Scr > 1 THEN
          PressButton PgUpKey, 25, 69, 78
        END IF
      END SELECT
    END IF
    
    
    '*************************************************************************
    '-------------------------------------------------------------------------
    SELECT CASE Frm(1).KeyCode
    CASE F0Key, EscKey
      SELECT CASE Frm(1).KeyCode
      CASE EscKey
        IF Frm(1).Edited THEN
          SaveFlag = PromptSaveData
        ELSE
          SaveFlag = False
        END IF
      CASE F0Key
        IF Frm(1).Edited THEN
          SaveFlag = True
        ELSE
          SaveFlag = False
        END IF
      END SELECT
      SELECT CASE SaveFlag
      CASE True
        'save true
        SaveScrn TempScrn()
        DisplayMiscScrn UpdatingDisk
        BCopy SSEG(FormCFFT$(0, 0)), SADD(FormCFFT$(0, 0)), VARSEG(FedTaxRec(1)), VARPTR(FedTaxRec(1)), FederalTaxRecLen, 0
        FPutAH FederalTaxFileName, FedTaxRec(1), FederalTaxRecLen, 1
        DisplayMiscScrn UpdatedOK
        WaitForAction
        RestScrn TempScrn()
        ERASE TempScrn
        ExitFlag = True
      CASE False
        'save false
        ExitFlag = True
      CASE 1
        Action = 2
        ExitFlag = False
      END SELECT
    CASE ELSE
      
    END SELECT
    
    IF Frm(1).FldNo > PageInfo(Scr).LastFld OR Frm(1).FldNo < PageInfo(Scr).FirstFld THEN
      Action = 2
      SELECT CASE Frm(1).FldNo
      CASE IS < PageInfo(Scr).FirstFld
        IF Scr > 1 THEN Scr = Scr - 1
      CASE IS > PageInfo(Scr).LastFld
        IF Scr < NumOfScrns THEN Scr = Scr + 1
      END SELECT
      Lib2Scrn FormName$(Scr), ScrnLib(), MonoCode, -2, ErrCode
    END IF
    '*************************************************************************
    'Keep editing until the user presses
  LOOP UNTIL ExitFlag = True
  CursorOff
  
  ERASE FedTaxRec, TempScrn, FormName$
  ERASE ScrnLib, FormCFFT$, FldCFFT, PageInfo, Frm
  
  
END SUB

SUB CFLeaveFileMaint

  CursorOff
  
  FirstTime = True
  
  REDIM Frm(1)    AS FormInfo
  
  FormName$ = "LEAVE"
  
  REDIM LeaveRec(1 TO 1)   AS LeaveRecType
  LeaveRecLen = LEN(LeaveRec(1))
  
  DO 'start of main loop. For init & add a new table

    NumLeaveRec = FileSize(LeaveFileName) \ LeaveRecLen
    
    IF NumLeaveRec > 1 THEN
      REDIM LeaveRec(1 TO NumLeaveRec)   AS LeaveRecType
    END IF
    
    FGetAH LeaveFileName, LeaveRec(1), LeaveRecLen, NumLeaveRec
    
    IF FirstTime THEN
      OrgLeaveName$ = "prdata\OrgLeave.dat"
      FCreate OrgLeaveName$
      FPutAH OrgLeaveName$, LeaveRec(1), LeaveRecLen, NumLeaveRec
    END IF
    
    NumOfScrns = NumLeaveRec
    
    NumFlds = -1
    
    NumPageFlds = LibNumberOfFields(CtrlQLib, FormName$)
    
    FOR Scr = 1 TO NumOfScrns   '-- Get total number of fields
      NumFlds = NumFlds + NumPageFlds + 1
    NEXT
    
    REDIM Form$(NumFlds, 2)             'DIM the form data array
    REDIM Fld(NumFlds) AS FieldInfo     'DIM the field information array
    REDIM PageInfo(1 TO NumOfScrns) AS PageInfoType
    
    StartEl = 0                         'Load first form at array start
    
    FOR Scr = 1 TO NumOfScrns   'Get field defs
      LibGetFldDef CtrlQLib, FormName$, StartEl, Fld(), Form$(), ErrCode
      PageInfo(Scr).PageNo = Scr
      PageInfo(Scr).FirstFld = StartEl + 1
      StartEl = StartEl + Fld(StartEl).Fields + 1               'Inc StartEl to next
      PageInfo(Scr).LastFld = StartEl - 1
    NEXT
    
    FOR n = 1 TO NumFlds        'For all fields on both pages
      LSET Form$(n, 0) = ""     'Clear each field
    NEXT
    
    Form$(0, 0) = SPACE$(LeaveRecLen * NumLeaveRec)
    BCopy VARSEG(LeaveRec(1)), VARPTR(LeaveRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), LeaveRecLen * NumLeaveRec, 0
    UnPackBuffer 0, 0, Form$(), Fld()
    
    '----- Setup TYPE for setting and reading form editing information.
    IF FirstTime THEN
      Frm(1).FldNo = 1          'Start editing on field #1
      Frm(1).InsStat = False    'Set the insert state (-1 = Insert on)
      Frm(1).StartEl = 0        'Set form starting element to 0 and
      Action = 1
      Scr = 1
      ClearScrn
    ELSE
      Frm(1).FldNo = PageInfo(NumLeaveRec).FirstFld             '+ FldInPage
      Frm(1).StartEl = PageInfo(NumLeaveRec).FirstFld - 1
      Action = 2
      Scr = NumLeaveRec
    END IF
    '----- Set the "Action" flag to force the editor to initialize itself and
    '      display the data on the form.
    
    IF FirstTime THEN
      FirstTime = False
      LibFile2Scrn CtrlQLib, FormName$, MonoCode, -1, ErrCode   'display screen
    END IF
    GOSUB ShowWhatTable
    
    DO
      EditForm Form$(), Fld(), Frm(1), Cnf, Action
      
      IF Frm(1).Edited THEN
        BeenEdited = True
      END IF
      
      IF Frm(1).Presses AND Frm(1).MRow = 25 THEN
        SELECT CASE Frm(1).MCol
        CASE 3 TO 12            'F1
          PressButton F1Key, 25, 3, 12
        CASE 14 TO 23           'F10
          PressButton F0Key, 25, 14, 23
        CASE 36 TO 45           'Escape
          PressButton EscKey, 25, 36, 45
        CASE 58 TO 68           'f8
          PressButton F8KEY, 25, 58, 68
        END SELECT
      END IF
      
      SELECT CASE Frm(1).KeyCode
      CASE F0Key, EscKey
        SELECT CASE Frm(1).KeyCode
        CASE EscKey
          IF BeenEdited THEN
            SaveFlag = PromptSaveData
          ELSE
            SaveFlag = False
          END IF
        CASE F0Key
          IF BeenEdited THEN
            SaveFlag = True
          ELSE
            SaveFlag = False
          END IF
        END SELECT
        SELECT CASE SaveFlag
        CASE True
          BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(LeaveRec(1)), VARPTR(LeaveRec(1)), LeaveRecLen * NumLeaveRec, 0
          FPutAH LeaveFileName, LeaveRec(1), LeaveRecLen, NumLeaveRec
          KillFile OrgLeaveName$
          DisplayMiscScrn UpdatedOK
          WaitForAction
          ExitFlag = True
        CASE False
          KillFile LeaveFileName
          NAME OrgLeaveName$ AS LeaveFileName
          ExitFlag = True
        CASE 1
          Action = 2
        END SELECT
      CASE F8KEY
        IF AskAddNewTbl THEN
          BeenEdited = True
          FldInPage = Frm(1).FldNo - PageInfo(Scr).FirstFld
          REDIM LeaveRec(1 TO NumLeaveRec + 1) AS LeaveRecType
          BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(LeaveRec(1)), VARPTR(LeaveRec(1)), LeaveRecLen * NumLeaveRec, 0
          FPutAH LeaveFileName, LeaveRec(1), LeaveRecLen, NumLeaveRec + 1
          InitFlag = True
        ELSE
          Action = 2
        END IF
      END SELECT
      IF Frm(1).FldNo > PageInfo(Scr).LastFld OR Frm(1).FldNo < PageInfo(Scr).FirstFld THEN
        FOR Cnt = 1 TO NumLeaveRec
          IF Frm(1).FldNo >= PageInfo(Cnt).FirstFld AND Frm(1).FldNo <= PageInfo(Cnt).LastFld THEN
            Scr = Cnt
            EXIT FOR
          END IF
        NEXT
        GOSUB ShowWhatTable
      END IF
      
    LOOP UNTIL ExitFlag = True OR InitFlag = True

    InitFlag = False

  LOOP UNTIL ExitFlag = True    '  the Escape key.
  
  CursorOff
  
  ERASE LeaveRec, Frm
  ERASE Form$, Fld
  EXIT SUB
  
ShowWhatTable:
  Text$ = QPStrI$(Scr) + " of " + QPStrI(NumLeaveRec) + "  "
  QPrintRC Text$, 1, 64, 112
RETURN
  
END SUB

SUB CFNewYearProcess
  
  ExitFlag = False
  OkFlag = False
  
  StartEl = 0   'Load first form at array start
  FormName$ = ZeroEmpYTD
  NumFlds = LibNumberOfFields(MiscQLib, FormName$)
  
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  LibGetFldDef MiscQLib, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  '----- Setup TYPE for setting and reading form editing information.
  
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set insert state (True = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  BlockClear
  DisplayMiscScrn FormName$
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE 89, 121                'yes
      CFZeroEmpYTD

      OkFlag = True
      
    CASE EscKey, 78, 110        'no
      OkFlag = True
      ExitFlag = True
    END SELECT
    
  LOOP UNTIL OkFlag             'proper key not set
  
  CursorOff
  
  ERASE Form$, Fld, Frm
  
END SUB

SUB CFPrinterMaint
  
  SHARED PrnDef$()
  
  REDIM PRNSetupRec(1)   AS PRNSetupRecType
  REDIM Frm(1) AS FormInfo
  REDIM TempScrn(0)
  REDIM Choice$(13, 0)          '
  
  'LibName$ = CtrlQLib     'Specify a Screen Library file name.
  FormName$ = "PRNSETUP"
  LibSize = ScrnLibSize(CtrlQLib)               'Get scrn lib array size
  REDIM ScrnLib(LibSize)        'Create scrn lib array
  LoadScrnLib CtrlQLib, ScrnLib(), ErrorCode    'load scrn lib
  NumFlds = LibNumberOfFields(CtrlQLib, FormName$)
  REDIM FormCFL$(NumFlds, 2)    'DIM the form data array
  REDIM FldCFL(NumFlds) AS FieldInfo            'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef CtrlQLib, FormName$, StartEl, FldCFL(), FormCFL$(), ErrCode
  
  PRNSetupRecLen = LEN(PRNSetupRec(1))
  FormCFL$(0, 0) = SPACE$(PRNSetupRecLen)
  'PrinterSetUPFile
  FGetAH PrinterSetUPFile, PRNSetupRec(1), PRNSetupRecLen, 1
  
  BCopy VARSEG(PRNSetupRec(1)), VARPTR(PRNSetupRec(1)), SSEG(FormCFL$(0, 0)), SADD(FormCFL$(0, 0)), PRNSetupRecLen, 0
  UnPackBuffer 0, 0, FormCFL$(), FldCFL()
  
  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  ClearScrn
  
  DO            'Pole the editing procedure
    Lib2Scrn FormName$, ScrnLib(), MonoCode, -1, ErrCode
    QPrintRC PrnDef$(0), 4, 31, -1
    
    DO
      EditForm FormCFL$(), FldCFL(), Frm(1), Cnf, Action
      
      IF Frm(1).Presses AND Frm(1).MRow = 25 THEN
        SELECT CASE Frm(1).MCol
        CASE 3 TO 12            'F1
          PressButton F1Key, 25, 3, 12
        CASE 14 TO 23           'F10
          PressButton F0Key, 25, 14, 23
        CASE 36 TO 45           'Escape
          PressButton EscKey, 25, 36, 45
        END SELECT
      END IF
      
      SELECT CASE Frm(1).KeyCode
      CASE F0Key, EscKey
        SELECT CASE Frm(1).KeyCode
        CASE EscKey
          IF Frm(1).Edited THEN
            SaveFlag = PromptSaveData
          ELSE
            SaveFlag = False
          END IF
        CASE F0Key
          IF Frm(1).Edited THEN
            SaveFlag = True
          ELSE
            SaveFlag = False
          END IF
        END SELECT
        SELECT CASE SaveFlag
        CASE True
          BCopy SSEG(FormCFL$(0, 0)), SADD(FormCFL$(0, 0)), VARSEG(PRNSetupRec(1)), VARPTR(PRNSetupRec(1)), PRNSetupRecLen, 0
          FPutAH PrinterSetUPFile, PRNSetupRec(1), PRNSetupRecLen, 1
          DisplayMiscScrn UpdatedOK
          WaitForAction
          ExitFlag = True
        CASE False
          ExitFlag = True
        CASE 1
          Action = 2
        END SELECT
      END SELECT
      
    LOOP UNTIL ExitFlag = True
    
  LOOP UNTIL ExitFlag = True    '  the Escape key.
  
  CursorOff
  
  ERASE PRNSetupRec, TempScrn, Frm
  ERASE ScrnLib, FormCFL$, FldCFL
  
END SUB

SUB CFRetireFileMaint
  
  REDIM BlankRetRec(1)   AS RetireRecType
  REDIM Frm(1) AS FormInfo
  
  REDIM RetireDataRec(1 TO 6)    AS RetireRecType
  REDIM TempScrn(0)
  
  'LibName$ = CtrlQLib     'Specify a Screen Library file name.
  FormName$ = "RETIRE"
  LibSize = ScrnLibSize(CtrlQLib)               'Get scrn lib array size
  REDIM ScrnLib(LibSize)        'Create scrn lib array
  LoadScrnLib CtrlQLib, ScrnLib(), ErrorCode    'load scrn lib
  NumFlds = LibNumberOfFields(CtrlQLib, FormName$)
  REDIM FormCFR$(NumFlds, 2)    'DIM the form data array
  REDIM FldCFR(NumFlds) AS FieldInfo            'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef CtrlQLib, FormName$, StartEl, FldCFR(), FormCFR$(), ErrCode
  BlankRetRec(1).TYPEDES1 = ""
  BlankRetRec(1).TYPEWH1 = 0
  BlankRetRec(1).TYPEM1 = 0
  BlankRetRec(1).TYPEOT1 = "Y"
  BlankRetRec(1).TYPETD1 = "Y"
  
  RetireRecLen = LEN(RetireDataRec(1)) * 6
  FormCFR$(0, 0) = SPACE$(RetireRecLen)
  FGetAH RetireFileName, RetireDataRec(1), LEN(RetireDataRec(1)), 6
  BCopy VARSEG(RetireDataRec(1)), VARPTR(RetireDataRec(1)), SSEG(FormCFR$(0, 0)), SADD(FormCFR$(0, 0)), RetireRecLen, 0
  UnPackBuffer 0, 0, FormCFR$(), FldCFR()
  
  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  ClearScrn
  
  
    Lib2Scrn FormName$, ScrnLib(), MonoCode, -1, ErrCode
    
    DO
      EditForm FormCFR$(), FldCFR(), Frm(1), Cnf, Action
      
      IF Frm(1).Presses AND Frm(1).MRow = 25 THEN
        SELECT CASE Frm(1).MCol
        CASE 3 TO 12            'F1
          PressButton F1Key, 25, 3, 12
        CASE 14 TO 23           'F10
          PressButton F0Key, 25, 14, 23
        CASE 36 TO 45           'Escape
          PressButton EscKey, 25, 36, 45
        CASE 58 TO 68           'F8
          PressButton F8KEY, 25, 58, 68
        END SELECT
      END IF
      
      SELECT CASE Frm(1).KeyCode
      CASE F0Key, EscKey
        SELECT CASE Frm(1).KeyCode
        CASE EscKey
          IF Frm(1).Edited THEN
            SaveFlag = PromptSaveData
          ELSE
            SaveFlag = False
          END IF
        CASE F0Key
          IF Frm(1).Edited THEN
            SaveFlag = True
          ELSE
            SaveFlag = False
          END IF
        END SELECT
        SELECT CASE SaveFlag
        CASE True
          GOSUB CFRetireRecPack
          FPutAH RetireFileName, RetireDataRec(1), LEN(RetireDataRec(1)), 6
          DisplayMiscScrn UpdatedOK
          WaitForAction
          ExitFlag = True
        CASE False
          ExitFlag = True
        CASE 1
          Action = 2
        END SELECT
      CASE F8KEY                'Erase rec
        DescFld = FldNum("TYPEDES" + RIGHT$(QPRTrim$(FldCFR(Frm(1).FldNo).FldName), 1), FldCFR())
        LSET FormCFR$(DescFld, 0) = ""
        SaveField DescFld, FormCFR$(), FldCFR(), BadFld
        GOSUB CFRetireRecPack
        Frm(1).Edited = True
        Frm(1).FldNo = DescFld
        Action = 2
      END SELECT
    
  LOOP UNTIL ExitFlag = True    '
  
  CursorOff
  
  ERASE BlankRetRec, RetireDataRec, TempScrn, ScrnLib
  ERASE FormCFR$, FldCFR, Frm
  
  EXIT SUB
  
CFRetireRecPack:
  
  BCopy SSEG(FormCFR$(0, 0)), SADD(FormCFR$(0, 0)), VARSEG(RetireDataRec(1)), VARPTR(RetireDataRec(1)), LEN(FormCFR$(0, 0)), 0
  FOR Cnt = 1 TO 6
    IF LEN(QPRTrim$(RetireDataRec(Cnt).TYPEDES1)) = 0 THEN
      RetireDataRec(Cnt) = BlankRetRec(1)
      ThisRec = Cnt
      DO
        ThisRec = ThisRec + 1
        IF ThisRec = 7 THEN EXIT DO
        IF LEN(QPRTrim$(RetireDataRec(ThisRec).TYPEDES1)) THEN
          SWAP RetireDataRec(ThisRec), RetireDataRec(Cnt)
          EXIT DO
        END IF
      LOOP
    END IF
  NEXT
  BCopy VARSEG(RetireDataRec(1)), VARPTR(RetireDataRec(1)), SSEG(FormCFR$(0, 0)), SADD(FormCFR$(0, 0)), RetireRecLen, 0
  UnPackBuffer 0, 0, FormCFR$(), FldCFR()
  RETURN
  
END SUB

SUB CFStateTaxMaint
  
  SHARED TaxText$()
  
  REDIM Frm(1)    AS FormInfo
  REDIM StatTaxRec(1)  AS StateTaxRecType

  StateTaxRecLen = LEN(StatTaxRec(1))
  
  CursorOff
  
  REDIM TempScrn(0)
  
  NumOfScrns = 3
  F7Flag = False
  
  REDIM FormName$(1 TO NumOfScrns)
  FOR Scr = 1 TO NumOfScrns
    FormName$(Scr) = "STTAX" + LTRIM$(STR$(Scr))
  NEXT
  LibSize = ScrnLibSize(CtrlQLib)               'Get scrn lib array size
  REDIM ScrnLib(LibSize)        'Create scrn lib array
  LoadScrnLib CtrlQLib, ScrnLib(), ErrorCode    'load scrn lib
  NumFlds = -1
  
  FOR Scr = 1 TO NumOfScrns     '-- Get total number of fields
    NumFlds = NumFlds + LibNumberOfFields(CtrlQLib, FormName$(Scr)) + 1
  NEXT
  
  REDIM FormCFST$(NumFlds, 2)   'DIM the form data array
  REDIM FldCFST(NumFlds) AS FieldInfo           'DIM the field information array
  REDIM PageInfo(1 TO NumOfScrns) AS PageInfoType
  
  StartEl = 0   'Load first form at array start
  FOR Scr = 1 TO NumOfScrns     'Get field defs
    LibGetFldDef CtrlQLib, FormName$(Scr), StartEl, FldCFST(), FormCFST$(), ErrCode
    PageInfo(Scr).PageNo = Scr
    PageInfo(Scr).FirstFld = StartEl + 1
    StartEl = StartEl + FldCFST(StartEl).Fields + 1             'Inc StartEl to next
    PageInfo(Scr).LastFld = StartEl - 1
  NEXT
  FOR n = 1 TO NumFlds          'For all fields on both pages
    LSET FormCFST$(n, 0) = ""   'Clear each field
  NEXT
  FormCFST$(0, 0) = SPACE$(StateTaxRecLen)
  
  FGetAH StateTaxFileName, StatTaxRec(1), StateTaxRecLen, 1
  BCopy VARSEG(StatTaxRec(1)), VARPTR(StatTaxRec(1)), SSEG(FormCFST$(0, 0)), SADD(FormCFST$(0, 0)), StateTaxRecLen, 0
  UnPackBuffer 0, 0, FormCFST$(), FldCFST()
  
  
  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  Scr = 1       'Set screen number to one
  ClearScrn
  
  Lib2Scrn FormName$(Scr), ScrnLib(), MonoCode, -1, ErrCode
  GOSUB FinishTaxScrn
  'display first screen
  
  
  DO
    
    EditForm FormCFST$(), FldCFST(), Frm(1), Cnf, Action
    
    '*--------------------------------------------
    
    IF Frm(1).Presses AND Frm(1).MRow = 25 THEN
      SELECT CASE Frm(1).MCol
      CASE 3 TO 12              'F1
        PressButton F1Key, 25, 3, 12
      CASE 14 TO 23             'F2
        PressButton F0Key, 25, 14, 23
      CASE 36 TO 45             'Escape
        PressButton EscKey, 25, 36, 45
      CASE 58 TO 67             'PageDown
        IF Scr < NumOfScrns THEN
          PressButton PgDnKey, 25, 58, 67
        END IF
      CASE 69 TO 78             'PageUp
        IF Scr > 1 THEN
          PressButton PgUpKey, 25, 69, 78
        END IF
      END SELECT
    END IF
    
    
    '*************************************************************************
    '-------------------------------------------------------------------------
    SELECT CASE Frm(1).KeyCode
    CASE F0Key, EscKey
      
      
      SELECT CASE Frm(1).KeyCode
      CASE EscKey
        IF Frm(1).Edited THEN
          
          SaveFlag = PromptSaveData
        ELSE
          SaveFlag = False
        END IF
      CASE F0Key
        IF Frm(1).Edited THEN
          SaveFlag = True
        ELSE
          SaveFlag = False
        END IF
      END SELECT
      SELECT CASE SaveFlag
      CASE True
        'save true
        SaveScrn TempScrn()
        DisplayMiscScrn UpdatingDisk
        BCopy SSEG(FormCFST$(0, 0)), SADD(FormCFST$(0, 0)), VARSEG(StatTaxRec(1)), VARPTR(StatTaxRec(1)), StateTaxRecLen, 0
        FPutAH StateTaxFileName, StatTaxRec(1), StateTaxRecLen, 1
        DisplayMiscScrn UpdatedOK
        WaitForAction
        RestScrn TempScrn()
        ERASE TempScrn
        ExitFlag = True
      CASE False
        'save false
        ExitFlag = True
      CASE 1
        Action = 2
        ExitFlag = False
      END SELECT
    CASE ELSE
      
    END SELECT
    
    IF Frm(1).FldNo > PageInfo(Scr).LastFld OR Frm(1).FldNo < PageInfo(Scr).FirstFld THEN
      Action = 2
      SELECT CASE Frm(1).FldNo
      CASE IS < PageInfo(Scr).FirstFld
        IF Scr > 1 THEN Scr = Scr - 1
      CASE IS > PageInfo(Scr).LastFld
        IF Scr < NumOfScrns THEN Scr = Scr + 1
      END SELECT
      Lib2Scrn FormName$(Scr), ScrnLib(), MonoCode, -2, ErrCode
      GOSUB FinishTaxScrn
    END IF
    '*************************************************************************
    'Keep editing until the user presses
  LOOP UNTIL ExitFlag = True
  
  CursorOff
  
  ERASE StatTaxRec, TempScrn, FormName$, ScrnLib
  ERASE FormCFST$, FldCFST, PageInfo, Frm
  
  EXIT SUB
  
FinishTaxScrn:
  
  FOR Cnt = 4 TO 10
    QPrintRC TaxText$(Cnt), Cnt - 1, 21, -1
  NEXT
  QPrintRC TaxText$(Scr), 1, 21, 112
  
  
  RETURN
  
END SUB

SUB CFSysFileMaint
  
  REDIM SysFileRec(1)   AS RegDSysFileRecType
  REDIM Frm(1) AS FormInfo
  
  REDIM TempScrn(0)
  
  'LibName$ = CtrlQLib                   'Specify a Screen Library file name.
  
  FormName$ = "REGDSYS"
  LibSize = ScrnLibSize(CtrlQLib)             'Get scrn lib array size
  REDIM ScrnLib(LibSize)                      'Dim scrn lib array
  LoadScrnLib CtrlQLib, ScrnLib(), ErrorCode  'load scrn lib
  NumFlds = LibNumberOfFields(CtrlQLib, FormName$)

  REDIM Form$(NumFlds, 2)                  'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo          'DIM the field information array
  StartEl = 0                              'Load first form at array start
  LibGetFldDef CtrlQLib, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  SysFileRecLen = LEN(SysFileRec(1))
  
  Form$(0, 0) = SPACE$(SysFileRecLen)
  
  FGetAH SysFileName, SysFileRec(1), SysFileRecLen, 1
  BCopy VARSEG(SysFileRec(1)), VARPTR(SysFileRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), SysFileRecLen, 0
  UnPackBuffer 0, 0, Form$(), Fld()
  
  
  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  ClearScrn
  
  
    Lib2Scrn FormName$, ScrnLib(), MonoCode, -1, ErrCode
    
    DO
      
      EditForm Form$(), Fld(), Frm(1), Cnf, Action
      IF Frm(1).Presses AND Frm(1).MRow = 25 THEN
        SELECT CASE Frm(1).MCol
        CASE 3 TO 12            'F1
          PressButton F1Key, 25, 3, 12
        CASE 14 TO 23           'F10
          PressButton F0Key, 25, 14, 23
        CASE 36 TO 45           'Escape
          PressButton EscKey, 25, 36, 45
        END SELECT
      END IF
      
      SELECT CASE Frm(1).KeyCode
      CASE F0Key, EscKey
        SELECT CASE Frm(1).KeyCode
        CASE EscKey
          IF Frm(1).Edited THEN
            SaveFlag = PromptSaveData
          ELSE
            SaveFlag = False
          END IF
        CASE F0Key
          IF Frm(1).Edited THEN
            SaveFlag = True
          ELSE
            SaveFlag = False
          END IF
        END SELECT
        SELECT CASE SaveFlag
        CASE True
          BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(SysFileRec(1)), VARPTR(SysFileRec(1)), SysFileRecLen, 0
          FPutAH SysFileName, SysFileRec(1), SysFileRecLen, 1
          DisplayMiscScrn UpdatedOK
          WaitForAction
          ExitFlag = True
        CASE False
          ExitFlag = True
        CASE 1
          Action = 2
        END SELECT
      END SELECT
      
    LOOP UNTIL ExitFlag = True
  
  CursorOff
  ERASE SysFileRec, TempScrn, ScrnLib, Form$, Fld, Frm
  
END SUB

SUB CFUnitFileMaint
  SHARED TaxText$()
  
  REDIM UnitFileRec(1)   AS UnitFileRecType
  REDIM Frm(1) AS FormInfo
  REDIM TempScrn(0)
  
  FormName$ = "UNITFILE"
  LibSize = ScrnLibSize(CtrlQLib)               'Get scrn lib array size
  REDIM ScrnLib(LibSize)        'Create scrn lib array
  LoadScrnLib CtrlQLib, ScrnLib(), ErrorCode    'load scrn lib
  NumFlds = LibNumberOfFields(CtrlQLib, FormName$)
  REDIM Form$(NumFlds, 2)     'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo             'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef CtrlQLib, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  UnitFileRecLen = LEN(UnitFileRec(1))
  Form$(0, 0) = SPACE$(UnitFileRecLen)
  
  FGetAH UnitFileName, UnitFileRec(1), UnitFileRecLen, 1
  BCopy VARSEG(UnitFileRec(1)), VARPTR(UnitFileRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), UnitFileRecLen, 0
  UnPackBuffer 0, 0, Form$(), Fld()

  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1               'Start editing on field #1
  Frm(1).InsStat = False         'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0             'Set form starting element to 0 and
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  WBaseFld = FldNum("TAXWBASE", Fld())
  ESCTypeFld = FldNum("ESCRTYPE", Fld())
  'LastFld = EndOfForms(Fld())
  
  ClearScrn
  
  'DO            'Pole the editing procedure
    Lib2Scrn FormName$, ScrnLib(), MonoCode, -1, ErrCode
    
    DO
      EditForm Form$(), Fld(), Frm(1), Cnf, Action
      'no vb eqv
      IF Frm(1).PrevFld <> Frm(1).FldNo THEN
        IF QPValI(Form$(ESCTypeFld, 0)) = 1 THEN
          Fld(WBaseFld).Protected = True
        ELSE
          Fld(WBaseFld).Protected = False
        END IF
      END IF
      
      IF Frm(1).Presses AND Frm(1).MRow = 25 THEN
        SELECT CASE Frm(1).MCol
        CASE 3 TO 12            'F1
          PressButton F1Key, 25, 3, 12
        CASE 14 TO 23           'F10
          PressButton F0Key, 25, 14, 23
        CASE 36 TO 45           'Escape
          PressButton EscKey, 25, 36, 45
        END SELECT
      END IF
      
      SELECT CASE Frm(1).KeyCode
      CASE F0Key, EscKey
        SELECT CASE Frm(1).KeyCode
        CASE EscKey
          IF Frm(1).Edited THEN
            SaveFlag = PromptSaveData
          ELSE
            SaveFlag = False
          END IF
        CASE F0Key
          IF Frm(1).Edited THEN
            SaveFlag = True
          ELSE
            SaveFlag = False
          END IF
        END SELECT
        SELECT CASE SaveFlag
        CASE True
          BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UnitFileRec(1)), VARPTR(UnitFileRec(1)), UnitFileRecLen, 0

          'UnitFileRec(1).FileVer = 9.1

          FPutAH UnitFileName, UnitFileRec(1), UnitFileRecLen, 1
          DisplayMiscScrn UpdatedOK
          InitStateTaxDef
          WaitForAction
          ExitFlag = True
        CASE False
          ExitFlag = True
        CASE 1
          Action = 2
        END SELECT
      END SELECT
      
    LOOP UNTIL ExitFlag = True
    
  'LOOP UNTIL ExitFlag = True    '  the Escape key.
  
  CursorOff
  
END SUB

SUB CFZeroEmpYTD
  
  REDIM E3Rec(1) AS EmpData3Type
  E3Len = LEN(E3Rec(1))
  NumOfRecs = FileSize(EmpData3Name) \ E3Len
  BlockClear
  ShowProcessingScrn "Initializing New Payroll Year"
  
  'clear employee YTD file.
  
  FOpenS EmpData3Name, EHandle
  FOR Cnt = 1 TO NumOfRecs
    E3Rec(1).Data1RecNum = Cnt
    FPutRTA EHandle, E3Rec(1), CLNG(Cnt), E3Len
    ShowPctComp Cnt, NumOfRecs
  NEXT
  FClose EHandle
  
  ERASE E3Rec

'This will ZAP the trans pointers
'  REDIM E2Rec(1) AS EmpData2Type
'  E2Len = LEN(E2Rec(1))
'  NumOfRecs = FileSize(EmpData2Name) \ E2Len
'
'  FOpenS EmpData2Name, EHandle
'  FOR Cnt = 1 TO NumOfRecs
'    FGetRTA EHandle, E2Rec(1), CLNG(Cnt), E2Len
'    E2Rec(1).LastTransRec = 0
'    FPutRTA EHandle, E2Rec(1), CLNG(Cnt), E2Len
'    ShowPctComp Cnt, NumOfRecs
'  NEXT
'  FClose EHandle

'  ERASE E2Rec

  BlockClear
  DisplayMiscScrn UpdatedOK
  WaitForAction
  
END SUB

SUB LoadSysDraftRec (Form$())
  
  REDIM PRSysDraftRec(1) AS PRDraftRecType
  PRSysDraftRecLen = LEN(PRSysDraftRec(1))
  Form$(0, 0) = SPACE$(PRSysDraftRecLen)
  
  PRFile = FREEFILE
  OPEN DraftInfoFileName FOR RANDOM SHARED AS PRFile LEN = PRSysDraftRecLen
  GET PRFile, 1, PRSysDraftRec(1)
  CLOSE
  
  BCopy VARSEG(PRSysDraftRec(1)), VARPTR(PRSysDraftRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), PRSysDraftRecLen, 0
  
  ERASE PRSysDraftRec
  
END SUB

SUB PRCtrlMenu
  
  Choice = 1
  
  DO
    PayMenu ControlMenu, Choice, 12
    
    SELECT CASE Choice
    CASE 1
      CFUnitFileMaint
    CASE 2
      CFSysFileMaint
    CASE 3
      CFStateTaxMaint
    CASE 4
      CFFederalTaxMaint
    CASE 5
      CFEICTableMaint
    CASE 6
      CFLeaveFileMaint
    CASE 7
      CFDedFileMaint
    CASE 8
      CFErnFileMaint
    CASE 9
      CFRetireFileMaint
    CASE 10
      CFEditBankDraftRec
    CASE 11
      CFPrinterMaint
    CASE 12
      CFNewYearProcess
      Choice = 1
    CASE 12
      
    END SELECT
    
  LOOP UNTIL Choice = EscKey
  
END SUB

SUB SaveSysDraftRec (Form$())
  
  REDIM PRSysDraftRec(1) AS PRDraftRecType
  PRSysDraftRecLen = LEN(PRSysDraftRec(1))
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(PRSysDraftRec(1)), VARPTR(PRSysDraftRec(1)), PRSysDraftRecLen, 0
  
  PRFile = FREEFILE
  OPEN DraftInfoFileName FOR RANDOM SHARED AS PRFile LEN = PRSysDraftRecLen
  PUT PRFile, 1, PRSysDraftRec(1)
  CLOSE
  
  DisplayMiscScrn UpdatedOK
  WaitForAction
  
  ERASE PRSysDraftRec
  
END SUB

