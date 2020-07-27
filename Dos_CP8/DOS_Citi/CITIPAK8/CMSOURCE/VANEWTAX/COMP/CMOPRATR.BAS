DEFINT A-Z
DECLARE SUB DisplayCMScrn (ScrnName$)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB AddEditOperator ()
DECLARE SUB AuthorizeOperator (oktoadd%)
DECLARE SUB ValidateOperator (OperNum%, OperPassword$, OperOKFlag%)
DECLARE SUB DisplayOperEntry ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB BlockClear ()
DECLARE SUB QPrintRC (t$, r%, c%, clr%)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'

  '*** Citipak Include Files Needed ***
  '$INCLUDE: 'CMFILES.BI'
  ''$INCLUDE: 'ubtrans.BI'
  ''$INCLUDE: 'gl.bi'

  CONST False = 0, True = NOT False

SUB AddEditOperator
  
  AuthorizeOperator oktoadd
  
  IF oktoadd = 0 THEN EXIT SUB
  
  LibName$ = "CM"
  ScrnName$ = "CMNEWOP"

  BlockClear
  DisplayCMScrn ScrnName$

  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True

  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  Action = 1
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM SHARED AS CMFile LEN = CMOperRecLen
  NumRecs = LOF(CMFile) \ CMOperRecLen
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF Frm(1).PrevFld = 2 AND edit <> True THEN
      FOR Cnt = 1 TO NumRecs
        GET CMFile, Cnt, CMOperRec(1)
        IF VAL(Form$(1, 0)) = CMOperRec(1).OperatorNumber AND VAL(Form$(1, 0)) > 0 THEN
          Form$(2, 0) = CMOperRec(1).OperatorName
          Form$(3, 0) = CMOperRec(1).OperatorPassword
          Fld(1).Protected = True
          RecNo = Cnt
          edit = True
          Action = 1
          EXIT FOR
        END IF
      NEXT Cnt
    END IF
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      CMOperRec(1).OperatorNumber = VAL(Form$(1, 0))
      CMOperRec(1).OperatorName = Form$(2, 0)
      CMOperRec(1).OperatorPassword = Form$(3, 0)
      IF RecNo = 0 THEN
        PUT CMFile, LOF(CMFile) \ CMOperRecLen + 1, CMOperRec(1)
      ELSE
        PUT CMFile, RecNo, CMOperRec(1)
      END IF
      
      CLOSE CMFile
      Done = True
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
END SUB

SUB AuthorizeOperator (oktoadd)
  
  LibName$ = "CM"
  ScrnName$ = "OPERAUTH"

  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F

  BlockClear
  DisplayCMScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF Form$(1, 0) = "CASHMGMT" THEN
        oktoadd = 1
      ELSE
        oktoadd = 0
      END IF
      Done = True
    CASE ESC
      Done = True
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done

END SUB

SUB DisplayOperEntry
  BlockClear
  DisplayCMScrn "CHOOSEOP"
END SUB

SUB DisplayUBScrn (ScrnName$)
  LibFile2Scrn "UB", ScrnName$, MonoCode, Attribute%, ErrCode
END SUB

SUB OperatorEntry (OperOKFlag, OperNum, PostDate$)
  
  Action = 1
  added = 0
  
  LibName$ = "CM"
  ScrnName$ = "CHOOSEOP"

  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  DisplayOperEntry

  FirstTime = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(3, 0) = DATE$
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F7KEY
      AddEditOperator
      Action = 1
      DisplayOperEntry
    CASE F10Key
      OperNum = VAL(Form$(1, 0))
      OperPassword$ = Form$(2, 0)
      PostDate$ = Form$(3, 0)
      GOSUB CheckPostDate
      IF BadDate THEN
        QPrintRC "<-- DATE NOT SET", 12, 49, 4
        BEEP
        SLEEP 1
        DisplayOperEntry
        Frm(1).FldNo = 1
        LSET Form$(3, 0) = DATE$
        Action = 1
        GOTO Skip1
      END IF
      
      IF OperNum > 0 THEN
        ValidateOperator OperNum, OperPassword$, OperOKFlag
        IF OperOKFlag > 0 THEN
          Done = True
        ELSE
          DisplayOperEntry
          Action = 1
          Frm(1).FldNo = 1
        END IF
      ELSE
        OperOKFlag = 0
        Done = True
      END IF
    CASE ESC
      OperOKFlag = -1
      Done = True
    CASE ELSE
      Done = False
    END SELECT
Skip1:
  LOOP UNTIL Done
  
  EXIT SUB
  
CheckPostDate:
  IF VAL(LEFT$(PostDate$, 2)) < 1 OR VAL(LEFT$(PostDate$, 2)) > 12 THEN
    BadDate = True
  ELSEIF VAL(MID$(PostDate$, 4, 2)) < 1 OR VAL(MID$(PostDate$, 4, 2)) > 31 THEN
    BadDate = True
  ELSE
    BadDate = False
  END IF
RETURN
  
END SUB

SUB ValidateOperator (OperNum, OperPassword$, OperOKFlag)
  
  REDIM CMOperRec(1) AS CMOperRecType
  OperFileLength = LEN(CMOperRec(1))
  
  FileHandle = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM SHARED AS #FileHandle LEN = OperFileLength

  NumOperRecs = LOF(FileHandle) \ OperFileLength

  IF NumOperRecs = 0 THEN
    OperOKFlag = 0
    GOTO ExitValidate
  END IF

  FOR Cnt = 1 TO NumOperRecs
    GET FileHandle, Cnt, CMOperRec(1)
    IF OperNum = CMOperRec(1).OperatorNumber AND OperPassword$ = CMOperRec(1).OperatorPassword THEN
      OperOKFlag = Cnt
      EXIT FOR
    END IF
  NEXT

ExitValidate:
  CLOSE
  IF OperOKFlag = 0 THEN
    BEEP
    OK = MsgBox("CM", "BADOPER")
  END IF
  
END SUB

