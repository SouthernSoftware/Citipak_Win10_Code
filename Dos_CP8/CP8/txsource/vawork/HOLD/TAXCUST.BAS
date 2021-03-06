
DEFINT A-Z
DECLARE SUB AddEditRETaxCust (RecNo&)
DECLARE SUB AddEditPPTaxCust (RecNo&)
DECLARE SUB CustHistoryRpt ()
DECLARE SUB ShowCustHistory (CustRec&, TaxType%)
DECLARE SUB DelPersAbstract (PersRecs() AS LONG, WhatPers%, CustRec&)
DECLARE SUB DelPropAbstract (PropRecs() AS LONG, WhatProp, CustRec&)
DECLARE SUB MakePersPINFile ()
DECLARE SUB MakeRealPINFile ()
DECLARE SUB CustPersonalAbstract (CustRec&)
DECLARE SUB GetPersRecList (PersRecs() AS LONG, CustRec&)
DECLARE SUB GetPropRecList (PropRecs() AS LONG, CustRec&)
DECLARE SUB ClearScrn ()
DECLARE SUB HideCursor ()
DECLARE SUB AddEditTaxCust (RecNo&)
DECLARE SUB ClearBack ()
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB CustPropertyAbstract (RecNo&)
  
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Compare3% (BYVAL Seg1, BYVAL Adr1, BYVAL Seg2, BYVAL Adr2, NumBytes)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION OK2UPDateCust% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Round# (DblNumber#)
DECLARE FUNCTION ParseBillNum$ (Text$)

DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB CursorOff ()
DECLARE SUB DeleteCust ()
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB InsertT (SEG StartElement AS ANY, ElSize%, NumEls%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%, TaxType%)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION GetCustName$ (CustRec&, TaxType%)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'
  
  
  CONST False = 0, True = NOT False
  
  CrLf$ = CHR$(13) + CHR$(10)

SUB AddEditPPTaxCust (RecNo&)

  SHARED Choice$()
  REDIM Choice$(0 TO 4, 0)

  REDIM ScrnArray(0)

  LibName$ = "TAX"
  ScrnName$ = "VATXCSPP"

  '--define the multi-choice fields
  '--Get the total number of fields from all pages
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
    CustFormLen = CustFormLen + Fld(F).StorLen
  NEXT

  '--Must fill Form$(0, 0) to the correct len, before
  'using the BCopy function.
  Form$(0, 0) = SPACE$(CustFormLen)

  Frm(1).PageDisable = True
  Frm(1).StayOnField = True

  IF RecNo& > 0 THEN
    GOSUB PPLoadCustInfo
    UnPackBuffer 0, 0, Form$(), Fld()
    LSET Form$(1, 0) = QPTrim$(STR$(RecNo&))
  ELSE
    Form$(2, 0) = DATE$
    Form$(15, 0) = "Y"
    Form$(16, 0) = "Y"
    Form$(17, 0) = "N"
    Form$(18, 0) = "Y"
  END IF

  '--Set screen number to one and display screen
  Scr = 1
  Action = 1
  FirstTime = True

  DisplayTaxScrn ScrnName$

  ShowCursor

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      SaveScrn ScrnArray()
    END IF

    IF Frm(1).FldNo <> Frm(1).PrevFld THEN
      IF Frm(1).Edited THEN
        EditedFlag = True
      END IF
    END IF

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key 'F10
      GOSUB PPCheckCustInfo
      IF CustInfoOK THEN
        GOSUB PPUPDateCustFile
      END IF
    CASE F4KEY
      IF RecNo& > 0 THEN
        ShowCustHistory RecNo&, TaxType%
      END IF
    CASE F6KEY
      IF RecNo& > 0 THEN
        CustPersonalAbstract RecNo&
      ELSE
        GOSUB PPCheckCustInfo
        IF CustInfoOK THEN
          IF OK2UPDateCust% THEN
            GOSUB PPUPDateCustFile
            CustPersonalAbstract RecNo&
          END IF
        END IF
      END IF
      Action = 1


    CASE EscKey
      IF EditedFlag THEN
        SaveFlag = PromptSaveData
        SELECT CASE SaveFlag
        CASE True               'user wants to save
          StuffBuf CHR$(0) + CHR$(ABS(F10Key))
        CASE False              'user wants to abandon
          IF RecNo& = 0 THEN
          ELSE
          END IF
          ExitFlag = True
        CASE ELSE
          'continue editing
        END SELECT
        Action = 1
      ELSE
        ExitFlag = True
      END IF
    END SELECT

    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 6 TO 14            '--F1 Help
          PressButton F1Key, 22, 6, 14
        CASE 15 TO 26           '--F4
          PressButton F4KEY, 22, 15, 26
        CASE 27 TO 37           '--F6
          PressButton F6KEY, 22, 27, 37
        CASE 38 TO 51           '--F7 Msgs
          IF RecNo& > 0 THEN
            PressButton F7KEY, 22, 38, 51
          END IF
        CASE 52 TO 61           '--F10 Save Button
          PressButton F10Key, 22, 52, 61
        CASE 62 TO 74           '--ESC cancel button
          PressButton 27, 22, 62, 74
        END SELECT
      END SELECT                'row
    END IF

    '--Check screen page
  LOOP UNTIL ExitFlag

  EXIT SUB

PPCheckCustInfo:

  CustInfoOK = True             'assume all is well

  'FOR zz = 3 TO 5
  '  IF LEN(QPTrim$(Form$(zz, 0))) = 0 THEN
  '    Frm(1).FldNo = zz
  '    CustInfoOK = False        'all is not well
  '    GOTO PPCustInfoExit
  '  END IF
  'NEXT

  FOR zz = 10 TO 14
    IF zz <> 11 THEN
      IF LEN(QPTrim$(Form$(zz, 0))) = 0 THEN
        Frm(1).FldNo = zz
        CustInfoOK = False      'all is not well
        EXIT FOR
      END IF
    END IF
  NEXT

PPCustInfoExit:
  IF NOT CustInfoOK THEN
    SaveScrn ScrnArray()
    DisplayTaxScrn "ERRSCRN1"
    QPrintRC "This field can not be blank!", 10, 26, -1
    QPrintRC "Please correct and try again.", 12, 26, -1
    WaitForAction
    RestScrn ScrnArray()
    Frm(1).Edited = False
    EditedFlag = False
  END IF
  RETURN


PPUPDateCustFile:
  IF RecNo& > 0 THEN
    ExitFlag = True
  END IF
  SaveScrn ScrnArray()
  DisplayTaxScrn "UPDATDSK"
  GOSUB PPSaveTaxCustRec
  RestScrn ScrnArray()
  DisplayTaxScrn "UPDATEOK"
  WaitForAction
  RestScrn ScrnArray()
  Action = 1
  EditedFlag = False
  RETURN


PPSaveTaxCustRec:

  REDIM CustPin(1) AS PINRecType
  REDIM TaxCust(1)  AS TaxCustType

  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure
  FormLen = LEN(Form$(0, 0))    'Length of Cust Form$(0,0) Element


  TaxFile = FREEFILE
  OPEN PPTaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  NumOfRecs& = LOF(TaxFile) \ TaxCustLen

  IF RecNo& > 0 THEN
    GET TaxFile, RecNo&, TaxCust(1)
  ELSEIF RecNo& = 0 THEN
    RecNo& = NumOfRecs& + 1
    TaxCust(1).Acct = RecNo&
    CustPin(1).PIN = RecNo&
    TaxCust(1).PIN = RecNo&
  END IF
  'Copy Cust info from Form$(0,0) to the Cust Rec Structure
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(TaxCust(1)), VARPTR(TaxCust(1)), FormLen, 0

  'stamp this record with the file version number
  TaxCust(1).FileVer = CustRecVerNO
  PUT TaxFile, RecNo&, TaxCust(1)

  CLOSE TaxFile
  CLOSE #PHandle

  ERASE TaxCust, CustPin

  RETURN

PPLoadCustInfo:

  REDIM TaxCust(1) AS TaxCustType
  TaxRecLen = LEN(TaxCust(1))
  Form$(0, 0) = SPACE$(CustFormLen)
  TaxFile = FREEFILE
  OPEN PPTaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, RecNo&, TaxCust(1)
  CLOSE
  BCopy VARSEG(TaxCust(1)), VARPTR(TaxCust(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), CustFormLen, 0

  ERASE TaxCust

  RETURN

END SUB

SUB AddEditRETaxCust (RecNo&)
  
  SHARED Choice$()
  REDIM Choice$(0 TO 4, 0)
  
  REDIM ScrnArray(0)
  
  LibName$ = "TAX"
  ScrnName$ = "VATXCSRE"
  
  '--define the multi-choice fields
  '--Get the total number of fields from all pages
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
    CustFormLen = CustFormLen + Fld(F).StorLen
  NEXT
  
  '--Must fill Form$(0, 0) to the correct len, before
  'using the BCopy function.
  Form$(0, 0) = SPACE$(CustFormLen)
  
  Frm(1).PageDisable = True
  Frm(1).StayOnField = True
  
  IF RecNo& > 0 THEN
    GOSUB LoadCustInfo
    UnPackBuffer 0, 0, Form$(), Fld()
    LSET Form$(1, 0) = QPTrim$(STR$(RecNo&))
  ELSE
    Form$(2, 0) = DATE$
    Form$(15, 0) = "Y"
    Form$(16, 0) = "Y"
    Form$(17, 0) = "N"
    Form$(18, 0) = "Y"
  END IF
  
  '--Set screen number to one and display screen
  Scr = 1
  Action = 1
  FirstTime = True
  
  DisplayTaxScrn ScrnName$
  
  ShowCursor
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      SaveScrn ScrnArray()
    END IF
    
    IF Frm(1).FldNo <> Frm(1).PrevFld THEN
      IF Frm(1).Edited THEN
        EditedFlag = True
      END IF
    END IF
    
    '--Check for Key presses
    TaxType% = 1
    SELECT CASE Frm(1).KeyCode
    CASE F10Key 'F10
      GOSUB CheckCustInfo
      IF CustInfoOK THEN
        GOSUB UPDateCustFile
      END IF
    CASE F4KEY
      IF RecNo& > 0 THEN
        ShowCustHistory RecNo&, TaxType%
      END IF
 
    CASE F7KEY
      IF RecNo& > 0 THEN
        CustPropertyAbstract RecNo&
      ELSE
        GOSUB CheckCustInfo
        IF CustInfoOK THEN
          IF OK2UPDateCust% THEN
            GOSUB UPDateCustFile
            CustPropertyAbstract RecNo&
          END IF
        END IF
      END IF
      Action = 1
    CASE EscKey
      IF EditedFlag THEN
        SaveFlag = PromptSaveData
        SELECT CASE SaveFlag
        CASE True               'user wants to save
          StuffBuf CHR$(0) + CHR$(ABS(F10Key))
        CASE False              'user wants to abandon
          IF RecNo& = 0 THEN
          ELSE
          END IF
          ExitFlag = True
        CASE ELSE
          'continue editing
        END SELECT
        Action = 1
      ELSE
        ExitFlag = True
      END IF
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 6 TO 14            '--F1 Help
          PressButton F1Key, 22, 6, 14
        CASE 15 TO 26           '--F4
          PressButton F4KEY, 22, 15, 26
        CASE 27 TO 37           '--F6
          PressButton F6KEY, 22, 27, 37
        CASE 38 TO 51           '--F7 Msgs
          IF RecNo& > 0 THEN
            PressButton F7KEY, 22, 38, 51
          END IF
        CASE 52 TO 61           '--F10 Save Button
          PressButton F10Key, 22, 52, 61
        CASE 62 TO 74           '--ESC cancel button
          PressButton 27, 22, 62, 74
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
  LOOP UNTIL ExitFlag
  
  EXIT SUB
  
CheckCustInfo:
  
  CustInfoOK = True             'assume all is well
  
  FOR zz = 3 TO 5
    IF LEN(QPTrim$(Form$(zz, 0))) = 0 THEN
      Frm(1).FldNo = zz
      CustInfoOK = False        'all is not well
      GOTO CustInfoExit
    END IF
  NEXT
  
  FOR zz = 10 TO 14
    IF zz <> 11 THEN
      IF LEN(QPTrim$(Form$(zz, 0))) = 0 THEN
        Frm(1).FldNo = zz
        CustInfoOK = False      'all is not well
        EXIT FOR
      END IF
    END IF
  NEXT
  
CustInfoExit:
  IF NOT CustInfoOK THEN
    SaveScrn ScrnArray()
    DisplayTaxScrn "ERRSCRN1"
    QPrintRC "This field can not be blank!", 10, 26, -1
    QPrintRC "Please correct and try again.", 12, 26, -1
    WaitForAction
    RestScrn ScrnArray()
    Frm(1).Edited = False
    EditedFlag = False
  END IF
  RETURN
  
  
UPDateCustFile:
  IF RecNo& > 0 THEN
    ExitFlag = True
  END IF
  SaveScrn ScrnArray()
  DisplayTaxScrn "UPDATDSK"
  GOSUB SaveTaxCustRec
  RestScrn ScrnArray()
  DisplayTaxScrn "UPDATEOK"
  WaitForAction
  RestScrn ScrnArray()
  Action = 1
  EditedFlag = False
  RETURN
  
  
SaveTaxCustRec:
  

  REDIM TaxCust(1) AS TaxCustType
  
  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure
  FormLen = LEN(Form$(0, 0))    'Length of Cust Form$(0,0) Element
  
  
  TaxFile = FREEFILE
  OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  NumOfRecs& = LOF(TaxFile) \ TaxCustLen
  
  IF RecNo& > 0 THEN
    GET TaxFile, RecNo&, TaxCust(1)
  ELSEIF RecNo& = 0 THEN
    RecNo& = NumOfRecs& + 1
    TaxCust(1).Acct = RecNo&
    TaxCust(1).PIN = RecNo&
    END IF
  'Copy Cust info from Form$(0,0) to the Cust Rec Structure
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(TaxCust(1)), VARPTR(TaxCust(1)), FormLen, 0
  
  'stamp this record with the file version number
  TaxCust(1).FileVer = CustRecVerNO
  PUT TaxFile, RecNo&, TaxCust(1)
  
  CLOSE TaxFile
  CLOSE #PHandle
  
  ERASE TaxCust
  
  RETURN
  
LoadCustInfo:
  
  REDIM TaxCust(1) AS TaxCustType
  TaxRecLen = LEN(TaxCust(1))
  Form$(0, 0) = SPACE$(CustFormLen)
  TaxFile = FREEFILE
  OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, RecNo&, TaxCust(1)
  CLOSE
  BCopy VARSEG(TaxCust(1)), VARPTR(TaxCust(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), CustFormLen, 0
  
  ERASE TaxCust
  
  RETURN
  
END SUB

SUB CustPersonalAbstract (CustRec&)
  
  SHARED Choice$()
  REDIM Choice$(0 TO 4, 0)
  ShowRec$ = SPACE$(10)

  REDIM EntryScrn(0)
  REDIM ScrnArray(0)
  SaveScrn EntryScrn()
  
  REDIM PersEDRec(1) AS PersonalRecType
  REDIM BlankPers(1) AS PersonalRecType
  PersEDRecLen = LEN(PersEDRec(1))
  
  REDIM TaxCust(1) AS TaxCustType
  TaxRecLen = LEN(TaxCust(1))

  CustName$ = GetCustName$(CustRec&, TaxType%)


PersRestart:
  REDIM PersRecs(0 TO 0) AS LONG
  
  GetPersRecList PersRecs(), CustRec&
  
  Adding = 0
  Editing = 1
  
  LibName$ = "TAX"
  ScrnName$ = "VAPPABST"
  
  '--Get the total number of fields from all pages
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
    PersFormLen = PersFormLen + Fld(F).StorLen
  NEXT
  
  '**Must fill Form$(0, 0) to the correct len, before using BCopy function.
  Form$(0, 0) = SPACE$(PersFormLen)
  
  IF PersRecs(0) > 0 THEN
    CurPersRec = 1
    Mode = Editing
    GOSUB LoadPersRec
  ELSE
    CurPersRec = 0
    Mode = Adding
    FirstTime = True
  END IF
  
  Frm(1).PageDisable = True
  
  '--Set screen number to one and display screen
  Action = 1
  ShowRecFlag = True
  
  DisplayTaxScrn ScrnName$
  
  ShowCursor
  
  Frm(1).StayOnField = True
  
  DO
    
    IF Action = 1 THEN
      Frm(1).FldNo = 1
    END IF
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF ShowRecFlag THEN
      ShowRecFlag = False
      GOSUB ShowPersRecInfo
      GOSUB ShowPersEDMode
    END IF
    
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(2, 0) = DATE$
      LSET Form$(10, 0) = "N"
      LSET Form$(11, 0) = "N"
      Action = 1
    END IF
    
    IF Frm(1).FldNo <> Frm(1).PrevFld THEN
      IF Frm(1).Edited THEN
        EditedFlag = True
      END IF
      SELECT CASE Frm(1).PrevFld
      CASE 3 TO 9
        IF LEN(QPTrim$(Form$(Frm(1).PrevFld, 0))) = 0 THEN
          LSET Form$(Frm(1).PrevFld, 0) = "0.00"
          SaveField Frm(1).PrevFld, Form$(), Fld(), BadField
          Action = 2
        END IF
      END SELECT
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F3Key
      IF Mode = Editing THEN
        OK = MsgBox%("TAX.QSL", "DELCPERS")
        SELECT CASE OK
        CASE 1  'nope don't detach
          Action = 2  'reset action so property form will resume correctly
        CASE 2  'do detach
          DelPersAbstract PersRecs(), CurPersRec, CustRec&
          GOTO PersRestart
        END SELECT
      END IF
    
    CASE F10Key 'F10
      GOSUB CheckPersInfo
      IF PersInfoOK THEN
        GOSUB SavePersRec
      END IF
      
    CASE F8KEY
      IF Mode = Editing THEN
        BCopy VARSEG(BlankPers(1)), VARPTR(BlankPers(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), PersFormLen, 0
        UnPackBuffer 0, 0, Form$(), Fld()
        Mode = Adding
        Action = 1
        FirstTime = True
        ShowRecFlag = True
        EditedFlag = False
        OCurPersRec = CurPersRec
        CurPersRec = PersRecs(0) + 1
      END IF
      
    CASE PgDnKey
      IF PersRecs(0) > 0 AND Mode = Editing THEN
        IF CurPersRec < PersRecs(0) THEN
          CurPersRec = CurPersRec + 1
          GOSUB LoadPersRec
          Action = 1
          ShowRecFlag = True
        END IF
      END IF
      
    CASE PgUpKey
      IF Mode = Editing THEN
        IF CurPersRec > 1 THEN
          CurPersRec = CurPersRec - 1
          GOSUB LoadPersRec
          Action = 1
          Frm(1).FldNo = 1
          ShowRecFlag = True
        END IF
      END IF
      
    CASE EscKey
      IF EditedFlag THEN
        IF Mode = Editing THEN
          GOSUB PersPromptSave
          ExitFlag = True
        ELSE
          GOSUB PersPromptSave
        END IF
      ELSE
        IF Mode = Adding AND PersRecs(0) > 0 THEN
          Mode = Editing
          CurPersRec = OCurPersRec
          ShowRecFlag = True
          EditedFlag = False
          Action = 1
          GOSUB LoadPersRec
        ELSE
          ExitFlag = True
        END IF
      END IF
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 23 TO 32           'ESC Cancel
          PressButton EscKey, 22, 23, 32
        CASE 33 TO 43           'Detach property
          PressButton F3Key, 22, 33, 43
        CASE 44 TO 51           'Add New
          PressButton F8KEY, 22, 44, 51
        CASE 52 TO 61           '--F10 Save Button
          PressButton F10Key, 22, 52, 61
        CASE 62 TO 67           'PgDn Next
          PressButton PgDnKey, 22, 62, 67
        CASE 68 TO 73           'PgUp Previous
          PressButton PgUpKey, 22, 68, 73
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
  LOOP UNTIL ExitFlag
  
  RestScrn EntryScrn()
  ERASE EntryScrn
  
  EXIT SUB
  
ShowPersRecInfo:
  IF PersRecs(0) = 0 THEN
    LSET ShowRec$ = "1 of 0"
  ELSE
    LSET ShowRec$ = STR$(CurPersRec) + " of" + STR$(PersRecs(0))
  END IF
  QPrintRC CustName$, 4, 18, -1
  QPrintRC ShowRec$, 4, 57, -1
  RETURN
  
  
CheckPersInfo:
  
  PersInfoOK = True
  
  IF LEN(QPTrim$(Form$(1, 0))) = 0 THEN
    Frm(1).FldNo = 1
    PersInfoOK = False
    GOTO CheckPersExit
  END IF
  
  FOR zz = 3 TO 11
    IF LEN(QPTrim$(Form$(zz, 0))) = 0 THEN
      Frm(1).FldNo = zz
      PersInfoOK = False
      EXIT FOR
    END IF
  NEXT
  
CheckPersExit:
  IF NOT PersInfoOK THEN
    SaveScrn ScrnArray()
    DisplayTaxScrn "ERRSCRN1"
    QPrintRC "This field can not be blank!", 10, 26, -1
    QPrintRC "Please correct and try again.", 12, 26, -1
    WaitForAction
    RestScrn ScrnArray()
    Frm(1).Edited = False
    EditedFlag = False
  END IF
  Action = 2
  RETURN
  
  
  RETURN
  
ShowPersEDMode:

  SELECT CASE Mode
  CASE Adding
    Mode$ = "ADDING "
    LibFile2Scrn LibName$, "ADBTNSR", MonoCode%, -2, ErrorCode
  CASE Editing
    Mode$ = "EDITING"
    LibFile2Scrn LibName$, "EDBTNSR", MonoCode%, -2, ErrorCode
  END SELECT
  QPrintRC Mode$, 22, 12, -1
  RETURN
  
SavePersRec:
  SaveScrn ScrnArray()
  DisplayTaxScrn "UPDATDSK"
  
  REDIM PersEDRec(1) AS PersonalRecType
  
  TaxFile = FREEFILE
  OPEN PPTaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, CustRec&, TaxCust(1)
  CustPin& = CustRec&
  
  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersEDRecLen
  
  IF CurPersRec = 0 THEN        'adding a new Pers cust has no others
    BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(PersEDRec(1)), VARPTR(PersEDRec(1)), PersFormLen, 0
    PersEDRec(1).CustPin& = CustPin&
    NumPersRec& = LOF(PersFile) \ PersEDRecLen
    WhatPers& = NumPersRec& + 1
    PUT #PersFile, WhatPers&, PersEDRec(1)
    TaxCust(1).FirstPersRec = WhatPers&
    PUT TaxFile, CustRec&, TaxCust(1)
    REDIM PRESERVE PersRecs(0 TO 1) AS LONG
    PersRecs(0) = 1
    PersRecs(1) = WhatPers&
    CurPersRec = 1
    Mode = Editing
    ShowRecFlag = True
    Action = 1
    GOTO SavePersReturn
  END IF
  
  SELECT CASE Mode
  CASE Adding
    BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(PersEDRec(1)), VARPTR(PersEDRec(1)), PersFormLen, 0
    PersEDRec(1).CustPin& = CustPin&
    NumPersRec& = LOF(PersFile) \ PersEDRecLen
    WhatPers& = NumPersRec& + 1
    PUT #PersFile, WhatPers&, PersEDRec(1)
    LastPers& = PersRecs(PersRecs(0))
    GET #PersFile, LastPers&, PersEDRec(1)
    PersEDRec(1).NextRec = WhatPers&
    PUT #PersFile, LastPers&, PersEDRec(1)
    PCnt = PersRecs(0) + 1
    REDIM PRESERVE PersRecs(0 TO PCnt) AS LONG
    PersRecs(0) = PCnt
    PersRecs(PCnt) = WhatPers&
    Mode = Editing
    CurPersRec = PCnt
    ShowRecFlag = True
    Action = 1
  CASE Editing
    GET #PersFile, PersRecs(CurPersRec), PersEDRec(1)
    BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(PersEDRec(1)), VARPTR(PersEDRec(1)), PersFormLen, 0
    PUT #PersFile, PersRecs(CurPersRec), PersEDRec(1)
    Action = 1
    ShowRecFlag = True
  END SELECT
  
SavePersReturn:
  
  MakePersPINFile
  
  EditedFlag = False
  CLOSE TaxFile, PersFile
  RestScrn ScrnArray()
  DisplayTaxScrn "UPDATEOK"
  WaitForAction
  RestScrn ScrnArray()
  RETURN
  
LoadPersRec:
  REDIM PersEDRec(1) AS PersonalRecType
  PersFile = FREEFILE
  OPEN TaxPersFile FOR RANDOM SHARED AS PersFile LEN = PersEDRecLen
  GET #PersFile, PersRecs(CurPersRec), PersEDRec(1)
  CLOSE PersFile
  BCopy VARSEG(PersEDRec(1)), VARPTR(PersEDRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), PersFormLen, 0
  UnPackBuffer 0, 0, Form$(), Fld()
  RETURN
  
PersPromptSave:
  SaveFlag = PromptSaveData
  SELECT CASE SaveFlag
  CASE True     'user wants to save
    GOSUB SavePersRec
  CASE False    'user wants to abandon
    IF Mode = Adding AND PersRecs(0) > 0 THEN
      Mode = Editing
      CurPersRec = OCurPersRec
      ShowRecFlag = True
      EditedFlag = False
      Action = 1
      GOSUB LoadPersRec
    ELSE
      ExitFlag = True
    END IF
  CASE ELSE
    'continue editing
  END SELECT
  Action = 1
  
  RETURN
  
  
END SUB

SUB CustPropertyAbstract (CustRec&)
  
  SHARED Choice$()
  ShowRec$ = SPACE$(10)

  REDIM Choice$(0 TO 1, 0)
  
  REDIM EntryScrn(0)
  REDIM ScrnArray(0)
  SaveScrn EntryScrn()
  
  REDIM PropEDRec(1) AS PropertyRecType
  REDIM BlankProp(1) AS PropertyRecType
  PropEDRecLen = LEN(PropEDRec(1))
  
  REDIM TaxCust(1) AS TaxCustType
  TaxRecLen = LEN(TaxCust(1))
  TaxType% = 1
  CustName$ = GetCustName$(CustRec&, TaxType%)
  
PropRestart:

  REDIM PropRecs(0 TO 0) AS LONG
  
  GetPropRecList PropRecs(), CustRec&
  
  Adding = 0
  Editing = 1
  
  LibName$ = "TAX"
  ScrnName$ = "VAPROPAB"
  
  '--Get the total number of fields from all pages
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
    PropFormLen = PropFormLen + Fld(F).StorLen
  NEXT
  
  '**Must fill Form$(0, 0) to the correct len, before using BCopy function.
  'or, the system WILL CRASH!!!!
  Form$(0, 0) = SPACE$(PropFormLen)
  
  IF PropRecs(0) > 0 THEN
    CurPropRec = 1
    Mode = Editing
    GOSUB LoadPropRec
  ELSE
    CurPropRec = 0
    Mode = Adding
    FirstTime = True
  END IF
  
  Frm(1).PageDisable = True
  
  '--Set screen number to one and display screen
  Action = 1
  ShowRecFlag = True
  GOSUB InitMultiChoice
  DisplayTaxScrn ScrnName$
  ShowCursor
  Frm(1).StayOnField = True
  
  DO
    IF Action = 1 THEN
      Frm(1).FldNo = 1
    END IF
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF ShowRecFlag THEN
      ShowRecFlag = False
      GOSUB ShowPropRecInfo
      GOSUB ShowPropEDMode
    END IF
    
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(2, 0) = DATE$
      LSET Form$(9, 0) = "N"
      LSET Form$(10, 0) = "N"
      FOR zz = 12 TO 14
        LSET Form$(zz, 0) = "0.00"
      NEXT
      Action = 1
    END IF
    
    IF Frm(1).FldNo <> Frm(1).PrevFld THEN
      IF Frm(1).Edited THEN
        EditedFlag = True
      END IF
      SELECT CASE Frm(1).PrevFld
      CASE 12 TO 14
        IF LEN(QPTrim$(Form$(Frm(1).PrevFld, 0))) = 0 THEN
          LSET Form$(Frm(1).PrevFld, 0) = "0.00"
          SaveField Frm(1).PrevFld, Form$(), Fld(), BadField
          Action = 2
        END IF
      END SELECT
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key 'F10
      GOSUB CheckPropInfo
      IF PropInfoOK THEN
        GOSUB SavePropRec
      END IF

    CASE F3Key

      IF Mode = Editing THEN
        OK = MsgBox%("TAX.QSL", "DELCPROP")
        SELECT CASE OK
        CASE 1  'nope don't detach
          Action = 2  'reset action so property form will resume correctly
        CASE 2  'do detach
          DelPropAbstract PropRecs(), CurPropRec, CustRec&
          GOTO PropRestart
        END SELECT
      END IF

    CASE F8KEY
      IF Mode = Editing THEN
        'GOSUB AddPropertyHow
        BCopy VARSEG(BlankProp(1)), VARPTR(BlankProp(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), PropFormLen, 0
        UnPackBuffer 0, 0, Form$(), Fld()
        Mode = Adding
        Action = 1
        FirstTime = True
        ShowRecFlag = True
        EditedFlag = False
        OCurPropRec = CurPropRec
        CurPropRec = PropRecs(0) + 1
      END IF
      
    CASE PgDnKey
      IF PropRecs(0) > 0 AND Mode = Editing THEN
        IF CurPropRec < PropRecs(0) THEN
          CurPropRec = CurPropRec + 1
          GOSUB LoadPropRec
          Action = 1
          ShowRecFlag = True
        END IF
      END IF
      
    CASE PgUpKey
      IF Mode = Editing THEN
        IF CurPropRec > 1 THEN
          CurPropRec = CurPropRec - 1
          GOSUB LoadPropRec
          Action = 1
          Frm(1).FldNo = 1
          ShowRecFlag = True
        END IF
      END IF
    CASE EscKey
      IF EditedFlag THEN
        IF Mode = Editing THEN
          GOSUB PropPromptSave
          ExitFlag = True
        ELSE
          GOSUB PropPromptSave
        END IF
      ELSE
        IF Mode = Adding AND PropRecs(0) > 0 THEN
          Mode = Editing
          CurPropRec = OCurPropRec
          ShowRecFlag = True
          EditedFlag = False
          Action = 1
          GOSUB LoadPropRec
        ELSE
          ExitFlag = True
        END IF
      END IF
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 23 TO 32           'ESC Cancel
          PressButton EscKey, 22, 23, 32
        CASE 33 TO 43           'Detach property
          PressButton F3Key, 22, 33, 43
        CASE 44 TO 51           'Add New
          PressButton F8KEY, 22, 44, 51
        CASE 52 TO 61           '--F10 Save Button
          PressButton F10Key, 22, 52, 61
        CASE 62 TO 67           'PgDn Next
          PressButton PgDnKey, 22, 62, 67
        CASE 68 TO 73           'PgUp Previous
          PressButton PgUpKey, 22, 68, 73
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
  LOOP UNTIL ExitFlag
  
  RestScrn EntryScrn()
  ERASE EntryScrn
  
  EXIT SUB
  
CheckPropInfo:
  PropInfoOK = True
  IF LEN(QPTrim$(Form$(1, 0))) = 0 THEN
    Frm(1).FldNo = 1
    PropInfoOK = False
    GOTO CheckPropExit
  END IF
  
  FOR zz = 9 TO 10
    IF LEN(QPTrim$(Form$(zz, 0))) = 0 THEN
      Frm(1).FldNo = zz
      PropInfoOK = False
      GOTO CheckPropExit
    END IF
  NEXT
  
  FOR zz = 12 TO 14
    IF LEN(QPTrim$(Form$(zz, 0))) = 0 THEN
      Frm(1).FldNo = zz
      PropInfoOK = False
      EXIT FOR
    END IF
  NEXT
  
CheckPropExit:
  IF NOT PropInfoOK THEN
    SaveScrn ScrnArray()
    DisplayTaxScrn "ERRSCRN1"
    QPrintRC "This field can not be blank!", 10, 26, -1
    QPrintRC "Please correct and try again.", 12, 26, -1
    WaitForAction
    RestScrn ScrnArray()
    Frm(1).Edited = False
    EditedFlag = False
    Action = 2
  END IF
  RETURN
  
ShowPropRecInfo:
  IF PropRecs(0) = 0 THEN
    LSET ShowRec$ = "1 of 0"
  ELSE
    LSET ShowRec$ = STR$(CurPropRec) + " of" + STR$(PropRecs(0))
  END IF
  QPrintRC CustName$, 4, 18, -1
  QPrintRC ShowRec$, 4, 57, -1
  RETURN
  
ShowPropEDMode:
  SELECT CASE Mode
  CASE Adding
    Mode$ = "ADDING "
    LibFile2Scrn LibName$, "ADBTNSR", MonoCode%, -2, ErrorCode
  CASE Editing
    Mode$ = "EDITING"
    LibFile2Scrn LibName$, "EDBTNSR", MonoCode%, -2, ErrorCode
  END SELECT
  QPrintRC Mode$, 22, 12, -1
  RETURN
  
SavePropRec:
  SaveScrn ScrnArray()
  DisplayTaxScrn "UPDATDSK"
  
  REDIM PropEDRec(1) AS PropertyRecType
  
  TaxFile = FREEFILE
  OPEN RETaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, CustRec&, TaxCust(1)
  CustPin& = CustRec&
  
  PropFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS PropFile LEN = PropEDRecLen
  
  IF CurPropRec = 0 THEN        'adding a new prop cust has no others
    BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(PropEDRec(1)), VARPTR(PropEDRec(1)), PropFormLen, 0
    PropEDRec(1).CustPin& = CustPin&
    NumPropRec& = LOF(PropFile) \ PropEDRecLen
    WhatProp& = NumPropRec& + 1
    PUT #PropFile, WhatProp&, PropEDRec(1)
    TaxCust(1).FirstPropRec = WhatProp&
    PUT TaxFile, CustRec&, TaxCust(1)
    REDIM PRESERVE PropRecs(0 TO 1) AS LONG
    PropRecs(0) = 1
    PropRecs(1) = WhatProp&
    CurPropRec = 1
    Mode = Editing
    ShowRecFlag = True
    Action = 1
    GOTO SavePropReturn
  END IF
  
  SELECT CASE Mode
  CASE Adding
    BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(PropEDRec(1)), VARPTR(PropEDRec(1)), PropFormLen, 0
    PropEDRec(1).CustPin& = CustPin&
    NumPropRec& = LOF(PropFile) \ PropEDRecLen
    WhatProp& = NumPropRec& + 1
    PUT #PropFile, WhatProp&, PropEDRec(1)
    LastProp& = PropRecs(PropRecs(0))
    GET #PropFile, LastProp&, PropEDRec(1)
    PropEDRec(1).NextRec = WhatProp&
    PUT #PropFile, LastProp&, PropEDRec(1)
    PCnt = PropRecs(0) + 1
    REDIM PRESERVE PropRecs(0 TO PCnt) AS LONG
    PropRecs(0) = PCnt
    PropRecs(PCnt) = WhatProp&
    Mode = Editing
    CurPropRec = PCnt
    ShowRecFlag = True
    Action = 1
  CASE Editing
    GET #PropFile, PropRecs(CurPropRec), PropEDRec(1)
    BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(PropEDRec(1)), VARPTR(PropEDRec(1)), PropFormLen, 0
    PUT #PropFile, PropRecs(CurPropRec), PropEDRec(1)
    Action = 1
    ShowRecFlag = True
  END SELECT
  
SavePropReturn:
  
  MakeRealPINFile
  
  EditedFlag = False
  CLOSE TaxFile, PropFile
  RestScrn ScrnArray()
  DisplayTaxScrn "UPDATEOK"
  WaitForAction
  RestScrn ScrnArray()
RETURN
  
LoadPropRec:
  REDIM PropEDRec(1) AS PropertyRecType
  PropFile = FREEFILE
  OPEN TaxPropFile FOR RANDOM SHARED AS PropFile LEN = PropEDRecLen
  GET #PropFile, PropRecs(CurPropRec), PropEDRec(1)
  CLOSE PropFile
  BCopy VARSEG(PropEDRec(1)), VARPTR(PropEDRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), PropFormLen, 0
  UnPackBuffer 0, 0, Form$(), Fld()
RETURN
  
PropPromptSave:
  SaveFlag = PromptSaveData
  SELECT CASE SaveFlag
  CASE True     'user wants to save
    GOSUB CheckPropInfo
    IF PropInfoOK THEN
      GOSUB SavePropRec
    END IF
  CASE False    'user wants to abandon
    IF Mode = Adding AND PropRecs(0) > 0 THEN
      Mode = Editing
      CurPropRec = OCurPropRec
      ShowRecFlag = True
      EditedFlag = False
      Action = 1
      GOSUB LoadPropRec
    ELSE
      ExitFlag = True
    END IF
  CASE ELSE
    'continue editing
  END SELECT
  Action = 2
  
RETURN
  
InitMultiChoice:
  MinArraySize = 1
  MortFld = FldNum%("MORTCODE", Fld())
  
  REDIM MCodeRec(1) AS MortCodeRecType
  MCodeLen = LEN(MCodeRec(1))
  
  MCodeFile = FREEFILE
  OPEN TaxMCodeFile FOR RANDOM SHARED AS MCodeFile LEN = MCodeLen
  
  NumOfMCRec = LOF(MCodeFile) \ MCodeLen
  IF NumOfMCRec > MinArraySize THEN
    MinArraySize = NumOfMCRec
  END IF
  
  REDIM Choice$(0 TO MinArraySize, 0)
  
  FOR Cnt = 1 TO NumOfMCRec
    GET MCodeFile, Cnt, MCodeRec(1)
    Choice$(Cnt, 0) = MCodeRec(1).MORTCODE + "  " + MCodeRec(1).BName
  NEXT
  CLOSE MCodeFile
  Choice$(0, 0) = STR$(MortFld)
RETURN

'AddPropertyHow:
'RETURN

END SUB

FUNCTION OK2UPDateCust%
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  LibName$ = "TAX"
  SaveFlag = 2
  
  FormName$ = "OK2SAVE"
  NumFlds = LibNumberOfFields(LibName$, FormName$)
  
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)       'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo               'DIM the field information array
  StartEl = 0   'Load first form at array start
  LibGetFldDef LibName$, FormName$, StartEl, Fld(), Form$(), ErrCode
  
  
  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
  
  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1              'Start editing on field #1
  Frm(1).InsStat = False        'Set insert state (True = Insert on)
  Frm(1).StartEl = 0            'Set form starting element to 0 and
  
  DisplayTaxScrn FormName$
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE F0Key
      OK2UPDateCust% = True
      EXIT DO
    CASE IS <> 0
      OK2UPDateCust% = False
      EXIT DO
    END SELECT
  LOOP          'proper key not set
  
  CursorOff
  
  RestScrn TempScrn()
  
  ERASE TempScrn, Form$, Fld, Frm
  
END FUNCTION

SUB TAXCustomerMenu
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 5)
  
  MChoice$(1) = "Add a New Real Estate Customer"
  MChoice$(2) = "Add a New Personal Property Customer"
  MChoice$(3) = "Edit Existing Real Estate Customer"
  MChoice$(4) = "Edit Existing Personal Property Customer"
  MChoice$(5) = "Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 20   'limit the box length to go no lower than line 20
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
    
    ClearBack
    
    TitleBox 2, Col, MaxLen + 3, "TAX Customer Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      AddEditRETaxCust 0
    CASE 2
      AddEditPPTaxCust 0
     CASE 3
      ClearBack
      LookUp RecNo&, "Customer", False, True, False, 1
      IF RecNo& > 0 THEN
        AddEditRETaxCust RecNo&
      END IF
     CASE 4
      ClearBack
      LookUp RecNo&, "Customer", False, True, False, 2
      IF RecNo& > 0 THEN
        AddEditPPTaxCust RecNo&
      END IF

    CASE 5
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
END SUB

