DEFINT A-Z
DECLARE FUNCTION GetItemRecCnt& ()
DECLARE FUNCTION FAAssetLookUp% (TempCode$)
DECLARE SUB FALookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
DECLARE SUB DisplayFAScrn (ScrnName$)
DECLARE SUB ItemMaint ()
DECLARE SUB AddEditItem (RecNo&)
DECLARE SUB ClearScrn ()
DECLARE SUB PrintHelp (H$)
DECLARE SUB BlockClear ()
DECLARE SUB PrintTitle (Title$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Num2Date$ (TheDate%)
DECLARE SUB ubGetAcctStruct (GLFundLen%, GLAcctLen%, GLDetLen%)
DECLARE FUNCTION ubFindAcct% (AcctNum$)
DECLARE FUNCTION ubFmtAcct$ (AN$, FundLen%, AcctLen%, DetLen%)
DECLARE FUNCTION ubStripAcct$ (AcctNum$)
DECLARE FUNCTION ubAcctLookUp% (Acct$)
DECLARE SUB WaitForAction ()
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB SaveScrn (Array())
DECLARE SUB RestScrn (Array())

  
  CONST False = 0, True = NOT False
  CONST BlankInt% = -32767      'Specifies blank integer

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'newfa.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'GL.BI'
  
  STACK 5000
  

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 5)
  
  MChoice$(1) = " Item Maintenance"
  MChoice$(2) = " Reports Menu"
  MChoice$(3) = " Year End Processing  "
  MChoice$(4) = " Asset Code Maintenance "
  MChoice$(5) = " Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
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
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    
    TitleBox 2, Col, MaxLen + 3, "Fixed Assets Main Menu", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      ItemMaint
    CASE 2
      RUN "FAREPORT"
    CASE 3
      RUN "FAYREND"
    CASE 4
      RUN "FACODE"
    CASE 5
      HideCursor
      CLS
      END
    END SELECT
  LOOP
  
  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF
  IF Exist("cpmenu" + Ext$) THEN
    RUN "cpmenu"
  ELSE
    HideCursor
    ClearScrn
  END IF
  
  END

SUB AddEditItem (RecNo&)
  
  SHARED Choice$()
  
  REDIM Choice$(0 TO 4, 0)
  
  REDIM FAItemRec(1) AS FAItemRecType
  REDIM TempScrn(0)

  FAItemRecLen = LEN(FAItemRec(1))
  
  Choice$(0, 0) = "2"
  Choice$(1, 0) = "Active"
  Choice$(2, 0) = "Inactive"
  
  LibName$ = "FA"
  ScrnName$ = "FAITEMS"
  
  '--define the multi-choice fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  Frm(1).PageDisable = True

  TagFld = FldNum%("ITEMTAG", Fld())
  AcctFld = FldNum%("GLACCT", Fld())
  CodeFld = FldNum%("ASETCODE", Fld())
  LifeFld = FldNum%("ILIFE", Fld())
  AQDateFld = FldNum%("AQURDATE", Fld())
  EOLDateFld = FldNum%("EOLDATE", Fld())

  GOSUB LoadGLAcctInfo
  
  '--Set screen number to one and display screen
  DisplayFAScrn ScrnName$
  
  ShowCursor
  FirstTime = True
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      IF RecNo& = 0 THEN
        Form$(4, 0) = DATE$
      ELSE
        GOSUB LoadItemRec
      END IF
      Action = 1
    END IF

'Section to check for dupe tag           **********************************
    IF Frm(1).FldNo <> Frm(1).PrevFld AND Frm(1).PrevFld = TagFld THEN
      IF RecNo& > 0 THEN                 'if it's not a new item
        NewTag$ = QPTrim$(Form$(1, 0))   'get current item tag
        IF NewTag$ <> OrgTag$ THEN       'if the user changed this tag
          GOSUB CheckItemTag             'check to see if it's a dupe
          IF NOT TagOK THEN              'it was a dupe
            LSET Form$(1, 0) = OrgTag$   'set item tag back to the original
          END IF
        END IF
      ELSE                               'not editing, adding a new item
        GOSUB CheckItemTag               'check new tag see if it's a dupe
      END IF
    END IF
'***********************************************************

'Section to Calc EOLDate
    IF Frm(1).FldNo <> Frm(1).PrevFld AND Frm(1).PrevFld = LifeFld THEN
      CurIYears = Value(Form$(LifeFld, 0), ECode)
      IF CurIYears <> OrgIYears THEN
        AQDate = Date2Num(Form$(AQDateFld, 0))
        IF AQDate > -32000 THEN
          AYears = VAL(RIGHT$(Form$(AQDateFld, 0), 4))
          AYears = AYears + CurIYears
          NewEOLDate$ = LEFT$(Form$(AQDateFld, 0), 6) + QPTrim$(STR$(AYears))
          LSET Form$(EOLDateFld, 0) = NewEOLDate$
          Action = 1
          CurIYears = OrgIYears
        END IF
      END IF
    END IF
'************************************************

'Section to validate GL Accounts
    IF Frm(1).FldNo <> Frm(1).PrevFld AND Frm(1).PrevFld = AcctFld THEN
      IF Frm(1).FldNo < NumFlds THEN
        TempAcct$ = ubStripAcct$(Form$(AcctFld, 0))
        'strip out any dashes in the account number
        IF LEN(TempAcct$) > 0 THEN
          'if there was an account
          TempAcct$ = ubFmtAcct$(TempAcct$, GLFundLen, GLAcctLen, GLDetLen)
          'reformat to make sure it's ok for ubAcctLookUp
          AcctOK = ubAcctLookUp(TempAcct$)      'go look for it
          IF AcctOK > 0 THEN    'if it was a valid account
            LSET Form$(AcctFld, 0) = TempAcct$
            'set it back into the form formated correctly
          ELSE  'nope not a valid account
            Frm(1).FldNo = AcctFld
            LSET Form$(AcctFld, 0) = ""
          END IF
          Action = 1
        END IF
      END IF
    END IF
'****************************************************

'Section to Pick Asset Code
    IF Frm(1).FldNo <> Frm(1).PrevFld AND Frm(1).PrevFld = CodeFld THEN
      IF Frm(1).FldNo < NumFlds THEN
        TempCode$ = Form$(CodeFld, 0)
        CodeRecNo = FAAssetLookUp%(TempCode$)      'go look for it
        IF CodeRecNo > 0 THEN    'if it was a valid account
          LSET Form$(CodeFld, 0) = TempCode$
        ELSE  'nope not a valid account
          Frm(1).FldNo = CodeFld
          LSET Form$(CodeFld, 0) = ""
        END IF
        Action = 1
      END IF
    END IF
'**************************

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      GOSUB CheckAndSave
      IF ExitFlag THEN
        DisplayFAScrn "UPDATEOK"
        WaitForAction
      END IF
    END SELECT

    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 52 TO 61           '--Save Button
          PressButton -68, 22, 52, 61
        CASE 63 TO 74           '--cancel button
          PressButton 27, 22, 63, 74
        END SELECT
      END SELECT                'row
    END IF
    '--Check screen page
  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag
  
  EXIT SUB
  
CheckAndSave:
  
  ExitFlag = True    'assume all is well

  IF RecNo& = 0 THEN
    REDIM FAItemRec(1) AS FAItemRecType
  END IF
  
  Temp$ = QPTrim$(Form$(1, 0))
  IF LEN(Temp$) = 0 THEN
    FldNo = 1
    GOTO BadParm
  END IF
  LSET FAItemRec(1).ItemTag = Temp$

  Temp$ = QPTrim$(Form$(2, 0))
  IF LEN(Temp$) = 0 THEN
    FldNo = 2
    GOTO BadParm
  END IF
  LSET FAItemRec(1).ISTATUS = Temp$            'only left most char

  Temp$ = QPTrim$(Form$(3, 0))
  IF LEN(Temp$) = 0 OR NOT (Temp$ = "Y" OR Temp$ = "N") THEN
    FldNo = 3
    GOTO BadParm
  END IF
  LSET FAItemRec(1).DEPYN = Temp$            'only left most char
  
  TempDate = Date2Num(Form$(4, 0))
  IF TempDate <= -32000 THEN
    FldNo = 4
    GOTO BadParm
  END IF
  FAItemRec(1).AQURDATE = TempDate
  
  LSET FAItemRec(1).IDESC1 = QPTrim$(Form$(5, 0))
  LSET FAItemRec(1).IDESC2 = QPTrim$(Form$(6, 0))

  Temp$ = QPTrim$(Form$(7, 0))
  'IF LEN(Temp$) = 0 THEN
  '  FldNo = 7
  '  GOTO BadParm
  'END IF

  LSET FAItemRec(1).GLAcct = Temp$
  LSET FAItemRec(1).IDEPT = QPTrim$(Form$(8, 0))

  FAItemRec(1).AssetCode = QPTrim$(Form$(9, 0))
  FAItemRec(1).CODEREC = CodeRecNo

'11
  Temp$ = QPTrim$(Form$(11, 0))
  IF LEN(Temp$) = 0 THEN
    FldNo = 11
    GOTO BadParm
  END IF
  DblNum# = Value#(Temp$, ECode)
  FAItemRec(1).ILife = DblNum#

'12
  Temp$ = QPTrim$(Form$(12, 0))
  IF LEN(Temp$) = 0 THEN
    FldNo = 12
    GOTO BadParm
  END IF
  DblNum# = Value#(Temp$, ECode)
  FAItemRec(1).ORGCOST = DblNum#

'13
  FAItemRec(1).DEP2DATE = Value#(Form$(13, 0), ECode)   'Total Depreciation
'14
  Temp$ = QPTrim$(Form$(14, 0))
  DblNum# = Value#(Temp$, ECode)
  FAItemRec(1).CurrVal = DblNum#

'15
  TempDate = Date2Num(Form$(15, 0))
  IF TempDate <= -10000 THEN
    FldNo = 15
    GOTO BadParm
  END IF
  FAItemRec(1).CDEPDATE = TempDate

'16
  FAItemRec(1).DISPDATE = Date2Num(Form$(16, 0))                       'Disposal Date

  IF FAItemRec(1).DISPDATE < 0 THEN
    'FAItemRec(1).DISPDATE =
    FAItemRec(1).DISPDATE = BlankInt%
  END IF

'17
  FAItemRec(1).VENDOR = QPTrim$(Form$(17, 0))
'18
  FAItemRec(1).SERIALNO = QPTrim$(Form$(18, 0))
'19
  FAItemRec(1).ITEMMFG = QPTrim$(Form$(19, 0))
'20 21 22
  FAItemRec(1).CONTACT = QPTrim$(Form$(20, 0))
  FAItemRec(1).ItemLoc = QPTrim$(Form$(21, 0))
  FAItemRec(1).EOLDate = Date2Num(Form$(22, 0))
  LSET FAItemRec(1).Fill1 = ""

  FAFile = FREEFILE
  OPEN FAItemFile FOR RANDOM SHARED AS FAFile LEN = FAItemRecLen
  IF RecNo& = 0 THEN
    RecNo& = (LOF(FAFile) / FAItemRecLen) + 1
  END IF
  PUT FAFile, RecNo&, FAItemRec(1)
  CLOSE
  GOTO SaveReturn

BadParm:
  Ok = MsgBox("FA", "NOBLANK")
  Action = 1
  ExitFlag = False
  Frm(1).FldNo = FldNo

SaveReturn:
  RETURN
  
LoadItemRec:
  
  REDIM FAItemRec(1) AS FAItemRecType
  FAFile = FREEFILE
  OPEN FAItemFile FOR RANDOM AS FAFile LEN = FAItemRecLen
  GET FAFile, RecNo&, FAItemRec(1)
  CLOSE

  BCopy VARSEG(FAItemRec(1)), VARPTR(FAItemRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), LEN(Form$(0, 0)), 0
  UnPackBuffer 0, 0, Form$(), Fld()
  OrgIYears = FAItemRec(1).ILife
  OrgTag$ = QPTrim$(FAItemRec(1).ItemTag)

RETURN
  
LoadGLAcctInfo:
  ubGetAcctStruct GLFundLen, GLAcctLen, GLDetLen
RETURN

CheckItemTag:
  NewTag$ = QPTrim$(Form$(1, 0))
  IF LEN(NewTag$) > 0 THEN       'if there is one
    TagOK = True                 'assume all is well
    REDIM FAItemRec(1) AS FAItemRecType
    NumItem& = GetItemRecCnt&
    FAFile = FREEFILE
    OPEN FAItemFile FOR RANDOM AS FAFile LEN = FAItemRecLen
    FOR Cnt& = 1 TO NumItem&
      GET FAFile, Cnt&, FAItemRec(1)
      ThisTag$ = QPTrim$(FAItemRec(1).ItemTag)
      IF NewTag$ = ThisTag$ THEN 'oops no good!
        TagOK = False            'all is not well
        EXIT FOR                 'no need to check any more
      END IF
    NEXT
    CLOSE                        'close-up shop
  END IF
  IF NOT TagOK THEN              'if we found a dupe
    SaveScrn TempScrn()          'save underling scrn
    DisplayFAScrn "DUPENUMB"     'display dupe error scrn
    WaitForAction                'wait for user action
    RestScrn TempScrn()          'restore to underling scrn.
    Frm(1).KeyCode = 0           'clear forms key return code
    Frm(1).FldNo = TagFld        'move back to Tag filed
    Action = 1                   'reset form action
  END IF
RETURN                           'back to caller

END SUB

FUNCTION GetItemRecCnt&

  REDIM FAItemRec(1) AS FAItemRecType
  FAItemRecLen = LEN(FAItemRec(1))
  ItemFile = FREEFILE
  OPEN FAItemFile FOR RANDOM SHARED AS ItemFile LEN = FAItemRecLen
  RecCnt& = LOF(ItemFile) \ FAItemRecLen
  CLOSE ItemFile
  GetItemRecCnt& = RecCnt&

END FUNCTION

SUB ItemMaint
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 3)
  
  MChoice$(1) = " Add a New Item"
  MChoice$(2) = " Edit an Existing Item "
'  MChoice$(3) = " Sell/Discard Item"
  MChoice$(3) = " Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 1
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    
    TitleBox 2, Col, MaxLen + 3, "Item Maintenance Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      'IF GetItemRecCnt& > 2 THEN
      '  CursorOff
      '  BlockClear
      '  DisplayFAScrn "DEMOONLY"
      '  WaitForAction
      'ELSE
        AddEditItem 0
      'END IF
    CASE 2
      BlockClear
      FALookUp RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%
      IF RecNo& > 0 THEN
        AddEditItem RecNo&
      END IF
    CASE 3
      HideCursor
      CLS
      END
    END SELECT
  LOOP
  
END SUB

