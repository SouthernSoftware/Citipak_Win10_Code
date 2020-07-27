DECLARE SUB TXLog (Text$)
DECLARE SUB DeleteTaxCustomer (RecNo&)
DEFINT A-Z
DECLARE SUB ExportData ()
DECLARE SUB CustHistoryRpt ()
DECLARE SUB ShowCustHistory (CustRec&)
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
DECLARE SUB LookUp (RecNo&, Text$, ChkBalFlag%, CLSFlag%, SSNFlag%)
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
DECLARE FUNCTION GetCustName$ (CustRec&)
DECLARE FUNCTION DoesCustOwe% (TaxCustRec AS ANY)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TaxCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'


  



  
  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE

  
  CONST False = 0, True = NOT False
  
  CrLf$ = CHR$(13) + CHR$(10)

SUB AddEditTaxCust (RecNo&)

  SHARED Choice$()
  REDIM Choice$(0 TO 4, 0)

  TXLog "TAX: IN ADD/EDIT Customer"

  IF RecNo& > 0 THEN
    TXLog "TAX: EDIT Customer:" + STR$(RecNo&)
  ELSE
    TXLog "TAX: ADDING NEW Customer."
  END IF

  REDIM ScrnArray(0)
  
  LibName$ = "TAX"
  ScrnName$ = "TAXCUST1"
  
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
    SELECT CASE Frm(1).KeyCode
    CASE F10Key 'F10
      GOSUB CheckCustInfo
      IF CustInfoOK THEN
        GOSUB UPDateCustFile
      END IF
    CASE F3Key
      IF RecNo& > 0 THEN
        DeleteTaxCustomer RecNo&
        Action = 1
        IF RecNo& < 0 THEN    'exit edit routine if they
          EXIT SUB            'deleted this account.
        END IF
      END IF
    CASE F4Key
      IF RecNo& > 0 THEN
        ShowCustHistory RecNo&
        Action = 1
      END IF
    CASE F6KEY
      IF RecNo& > 0 THEN
        CustPersonalAbstract RecNo&
      ELSE
        GOSUB CheckCustInfo
        IF CustInfoOK THEN
          IF OK2UPDateCust% THEN
            GOSUB UPDateCustFile
            CustPersonalAbstract RecNo&
          END IF
        END IF
      END IF
      Action = 1
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
          'IF RecNo& = 0 THEN
          'ELSE
          'END IF
          ExitFlag = True
        CASE ELSE
          'continue editing
        END SELECT
        Action = 1
      ELSE
        ExitFlag = True
      END IF
    CASE -106    'Alt-F3
      IF RecNo& > 0 THEN
        GOSUB DeleteCust
        IF KilledCust THEN
          GOTO AddEditExit
        END IF
      END IF
    
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 6 TO 14            '--F1 Help
          PressButton F1Key, 22, 6, 14
        CASE 15 TO 22            '--F3 Delete
          PressButton F3Key, 22, 15, 22
        CASE 23 TO 33           '--F4
          PressButton F4Key, 22, 23, 33
        CASE 34 TO 42           '--F6
          PressButton F6KEY, 22, 34, 42
        CASE 43 TO 51           '--F7 Msgs
          IF RecNo& > 0 THEN
            PressButton F7KEY, 22, 43, 51
          END IF

        CASE 52 TO 61           '--F10 Save Button
          PressButton F10Key, 22, 52, 61
        CASE 62 TO 73           '--ESC cancel button
          PressButton 27, 22, 62, 73
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
  LOOP UNTIL ExitFlag

AddEditExit:
EXIT SUB
  
CheckCustInfo:
  
  CustInfoOK = True             'assume all is well
  
  FOR zz = 4 TO 5
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
  
  REDIM CustPin(1) AS PINRecType
  REDIM TaxCust(1)  AS TaxCustType
  
  TaxCustLen = LEN(TaxCust(1))  'Length of Cust Record Structure
  FormLen = LEN(Form$(0, 0))    'Length of Cust Form$(0,0) Element
  
  PHandle = FREEFILE
  OPEN CustPinFile FOR RANDOM SHARED AS #PHandle LEN = 4
  IF LOF(PHandle) \ 4 > 0 THEN
    GET #PHandle, 1, CustPin(1) 'get last pin used info
  END IF
  
  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  NumOfRecs& = LOF(TaxFile) \ TaxCustLen
  IF RecNo& > 0 THEN
    TXLog "TAX: UPDATE Customer Rec:" + STR$(RecNo&)
    GET TaxFile, RecNo&, TaxCust(1)
  ELSEIF RecNo& = 0 THEN
    RecNo& = NumOfRecs& + 1
    TXLog "TAX: ADDING Customer Rec:" + STR$(RecNo&)
    TaxCust(1).Acct = RecNo&
    CustPin(1).PIN = RecNo&         'Increment last pin used
    TaxCust(1).PIN = RecNo&
    PUT #PHandle, 1, CustPin(1)
  END IF
  'Copy Cust info from Form$(0,0) to the Cust Rec Structure
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(TaxCust(1)), VARPTR(TaxCust(1)), FormLen, 0

  IF DeleteFlag THEN
    TaxCust(1).Deleted = -1
  END IF

  'stamp this record with the file version number
  TaxCust(1).FileVer = CustRecVerNO
  PUT TaxFile, RecNo&, TaxCust(1)
  
  CLOSE TaxFile
  CLOSE #PHandle

  TXLog "TAX: DONE Customer Rec:" + STR$(RecNo&)

  ERASE TaxCust, CustPin
  
RETURN
  
LoadCustInfo:
  REDIM TaxCust(1) AS TaxCustType
  TaxRecLen = LEN(TaxCust(1))
  Form$(0, 0) = SPACE$(CustFormLen)
  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, RecNo&, TaxCust(1)
  CLOSE
  BCopy VARSEG(TaxCust(1)), VARPTR(TaxCust(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), CustFormLen, 0
  ERASE TaxCust
RETURN

DeleteCust:
  ECode = 0
  REDIM TaxCust(1) AS TaxCustType
  TaxRecLen = LEN(TaxCust(1))
  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, RecNo&, TaxCust(1)
  CLOSE
  IF TaxCust(1).LastTrans > 0 THEN
    ECode = 1
  ELSEIF TaxCust(1).FirstPropRec > 0 THEN
    ECode = 2
  ELSEIF TaxCust(1).FirstPersRec > 0 THEN
    ECode = 3
  END IF
  IF ECode = 0 THEN
    OK = MsgBox%("TAX", "TAXCDEL")
    IF OK = 2 THEN
      OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
      GET TaxFile, RecNo&, TaxCust(1)
      TaxCust(1).Deleted = -1
      PUT TaxFile, RecNo&, TaxCust(1)
      KilledCust = -1
    END IF
  ELSE
    SaveScrn ScrnArray()
    DisplayTaxScrn "ErrScrn1"
    QPrintRC "Can not DELETE this Account!", 10, 26, -1
    SELECT CASE ECode
    CASE 1
      QPrintRC "Account Has Transaction History!", 12, 24, -1
    CASE 2, 3
      QPrintRC "Account Has Abstracts!", 12, 29, -1
    END SELECT
    WaitForAction
    RestScrn ScrnArray()
  END IF

Action = 1
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

  CUSTNAME$ = GetCustName$(CustRec&)


PersRestart:
  REDIM PersRecs(0 TO 0) AS LONG
  
  GetPersRecList PersRecs(), CustRec&
  
  Adding = 0
  Editing = 1
  
  LibName$ = "TAX"
  ScrnName$ = "PERSABST"
  
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
  
  '**Must fill Form$(0, 0) to the correct len, before using BCopy function!
  '  A lock-up is assured if not done!!!
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
      
    CASE F8Key
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
          PressButton F8Key, 22, 44, 51
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
  QPrintRC CUSTNAME$, 4, 18, -1
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
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, CustRec&, TaxCust(1)
  CustPin& = TaxCust(1).PIN
  
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

  CUSTNAME$ = GetCustName$(CustRec&)
  
PropRestart:

  REDIM PropRecs(0 TO 0) AS LONG
  
  GetPropRecList PropRecs(), CustRec&
  
  Adding = 0
  Editing = 1
  
  LibName$ = "TAX"
  ScrnName$ = "PROPABST"
  
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
          DelPropAbstract PropRecs(), CINT(PropRecs(CurPropRec)), CustRec&
          GOTO PropRestart
        END SELECT
      END IF

    CASE F8Key
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
          PressButton F8Key, 22, 44, 51
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
  QPrintRC CUSTNAME$, 4, 18, -1
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
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET TaxFile, CustRec&, TaxCust(1)
  CustPin& = TaxCust(1).PIN
  
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
    Choice$(Cnt, 0) = MCodeRec(1).MortCode + "  " + MCodeRec(1).BName
  NEXT
  CLOSE MCodeFile
  Choice$(0, 0) = STR$(MortFld)
RETURN

'AddPropertyHow:
'RETURN

END SUB

SUB DeleteTaxCustomer (RecNo&)

  TXLog "TAX: IN DELETE CUSTOMER"

  REDIM ScrnArray(0)
  SaveScrn ScrnArray()

  REDIM TaxCust(1) AS TaxCustType
  TaxRecLen = LEN(TaxCust(1))

  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  GET #TaxFile, RecNo&, TaxCust(1)
  CLOSE

  IF TaxCust(1).FirstPropRec > 0 OR TaxCust(1).FirstPersRec > 0 THEN
    ErrFlag = True
  END IF
  IF DoesCustOwe%(TaxCust(1)) THEN
    ErrFlag = True
  END IF
  IF ErrFlag THEN
    TXLog "ERR: CAN NOT DELETE CUSTOMER:" + STR$(RecNo&)
    OK = MsgBox%("TAX", "NODELSCR")
    RestScrn ScrnArray()
    GOTO DeleteExit
  END IF
  TXLog "TAX: ASK OK TO DELETE:" + STR$(RecNo&)
  OK = MsgBox%("TAX.QSL", "DELCUST")
  SELECT CASE OK
  CASE 2  'ok delete this account
    TXLog "TAX: YES DELETE:" + STR$(RecNo&)
    GOSUB DeleteThisAccount
  CASE ELSE 'nope don't delete
    TXLog "TAX: NO DO NOT DELETE:" + STR$(RecNo&)
  END SELECT

DeleteExit:
  TXLog "TAX: OUT DELETE CUSTOMER"
EXIT SUB

DeleteThisAccount:

  TaxFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxRecLen
  'NumOfRecs& = LOF(TaxFile) \ TaxCustLen
  GET #TaxFile, RecNo&, TaxCust(1)
  TaxCust(1).Deleted = True
  TaxCust(1).Active = "N"
  PUT #TaxFile, RecNo&, TaxCust(1)
  CLOSE
  DisplayTaxScrn "UPDATEOK"
  WaitForAction
  RecNo& = -1

RETURN
END SUB

SUB ExportData

  REDIM TaxSetupRec(1) AS TaxMasterType
  REDIM TaxCustRec(1) AS TaxCustType
  REDIM PropertyRec(1) AS PropertyRecType
  REDIM PersRec(1) AS PersonalRecType
  

  SHARED Choice$()

  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "TaxCust.PRN"   'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)


  LineCnt = 0
  CustCnt = 0

  LibName$ = "TAX"
  ScrnName$ = "EXPORT"

  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F

  REDIM Choice$(0 TO 2, 0 TO 0)

  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name Order"
  Choice$(2, 0) = "Account Number"

  Form$(2, 0) = "Y"
  Form$(3, 0) = "TAXDATA1"

  Action = 1
  ClearBack

  ShowCursor

  DisplayTaxScrn ScrnName$

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      IF LEFT$(Form$(1, 0), 1) = "N" THEN
        UsingIndex = True
      ELSE
        UsingIndex = False
      END IF
      IF LEFT$(Form$(2, 0), 1) = "Y" THEN
        DetailFlag = True
        
      ELSE
        DetailFlag = False
        
      END IF
      ExitFlag = True
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag

  IF AbortFlag THEN EXIT SUB

  RptHandle = FREEFILE

  OPEN ReportFile$ FOR OUTPUT AS #RptHandle

  FileName$ = RTRIM$(Form$(3, 0)) + ".TXT"
  Oh = FREEFILE
  OPEN FileName$ FOR OUTPUT AS #Oh

  GOSUB OpenTaxFile
  GOSUB OpenPropFile
  GOSUB OpenPersFile

  IF UsingIndex AND NumOfTaxRecs > 0 THEN
    GOSUB GetNameIndex
  END IF

  ClearBack
  ShowProcessingScrn "Exporting Customer Data"

  FOR Cnt = 1 TO NumOfTaxRecs
    IF UsingIndex THEN
      CustRecNo = Array(Cnt).RecNum
    ELSE
      CustRecNo = Cnt
    END IF

    GET TaxFile, CustRecNo, TaxCustRec(1)

    IF TaxCustRec(1).Deleted = 0 THEN
      ELSE
      TaxCustRec(1).Deleted = -1
      PUT TaxFile, CustRecNo, TaxCustRec(1)
    END IF

    IF NOT TaxCustRec(1).Deleted THEN
        PRINT #Oh, USING "#####"; CustRecNo;
        PRINT #Oh, "]";
        PRINT #Oh, (TaxCustRec(1).FName); "]";
        PRINT #Oh, (TaxCustRec(1).LName); "]";
        PRINT #Oh, (TaxCustRec(1).Addr1); "]";
        PRINT #Oh, (TaxCustRec(1).Addr2); "]";
        PRINT #Oh, (TaxCustRec(1).City); "]";
        PRINT #Oh, (TaxCustRec(1).State); "]";
        PRINT #Oh, (TaxCustRec(1).Zip); "]";
        PRINT #Oh, (TaxCustRec(1).CSSN);
        IF DetailFlag = False THEN
         PRINT #Oh, ""
         GOTO SkipEm
         ELSE
         PRINT #Oh, "]";
        END IF
        'Now Show Property Records Next

        IF TaxCustRec(1).FirstPropRec > 0 THEN

          PropertyRecord! = TaxCustRec(1).FirstPropRec

          WHILE PropertyRecord! <> 0


            GET #PropTaxFile, PropertyRecord!, PropertyRec(1)

            IF PropertyRecord! <> TaxCustRec(1).FirstPropRec THEN
                PRINT #Oh, USING "#####"; CustRecNo;
                PRINT #Oh, "]";
                PRINT #Oh, (TaxCustRec(1).FName); "]";
                PRINT #Oh, (TaxCustRec(1).LName); "]";
                PRINT #Oh, (TaxCustRec(1).Addr1); "]";
                PRINT #Oh, (TaxCustRec(1).Addr2); "]";
                PRINT #Oh, (TaxCustRec(1).City); "]";
                PRINT #Oh, (TaxCustRec(1).State); "]";
                PRINT #Oh, (TaxCustRec(1).Zip); "]";
                PRINT #Oh, (TaxCustRec(1).CSSN); "]";
             ELSE
            END IF

            PRINT #Oh, PropertyRec(1).REALPIN; "]";
            PRINT #Oh, PropertyRec(1).PROPNOT1; "]";
            PRINT #Oh, PropertyRec(1).PROPNOT2; "]";
            PRINT #Oh, PropertyRec(1).PROPNOT3; "]";
            PRINT #Oh, USING "##########"; PropertyRec(1).PropValu;
             PRINT #Oh, "]";
            PRINT #Oh, USING "##########"; PropertyRec(1).EXMPSENI;
             PRINT #Oh, "]";
            PRINT #Oh, PropertyRec(1).MortCode; "]";
            PRINT #Oh, PropertyRec(1).Map; "]";
            PRINT #Oh, PropertyRec(1).Block; "]";
            PRINT #Oh, PropertyRec(1).LotNumb; "]";
            PropertyRecord! = PropertyRec(1).NextRec
            
           WEND
           
           ELSE
            PRINT #Oh, "]"; "]"; "]"; "]"; "]"; "]"; "]"; "]"; "]"; "]";
           END IF


        'NOW CHECK PERSONAL PROPERTY

        IF TaxCustRec(1).FirstPersRec > 0 THEN

          
          PropertyRecord! = TaxCustRec(1).FirstPersRec


          WHILE PropertyRecord! <> 0

           GET #PersTaxFile, PropertyRecord!, PersRec(1)

            PRINT #Oh, PersRec(1).PropPin; "]";
            PRINT #Oh, USING "########"; PersRec(1).PersVal;
             PRINT #Oh, "]";
            PRINT #Oh, USING "########"; PersRec(1).EXMPSENI;
             PRINT #Oh, "]";
            PRINT #Oh, PersRec(1).DESC1; "]";
            PRINT #Oh, PersRec(1).DESC2; "]";
            PRINT #Oh, PersRec(1).DESC3; "]";
            PRINT #Oh, PersRec(1).DESC4; "]";
            PRINT #Oh, PersRec(1).DESC5; "]"
            PropertyRecord! = PersRec(1).NextRec
            IF OldRecord! = PropertyRecord! THEN PropertyRecord! = 0
          WEND
          ELSE
           PRINT #Oh, " ]"; " ]"; " ]"; " ]"; " ]"; " ]"; " ]"; " ]"
        END IF
        END IF
SkipEm:
      CustCnt = CustCnt + 1
    
    ShowPctComp Cnt, NumOfTaxRecs
  NEXT

  GOSUB PrintLayOutTable

  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi

  CLOSE         'Close all open files now

  IF DevSpec$ = "P" THEN
    EntryPoint = 4
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 1
  END IF

  ERASE Array, Frm, Form$, Fld, TaxCustRec

  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint

  KILL ReportFile$

  EXIT SUB


PrintLayOutTable:
  PRINT #RptHandle, "Data File Export Facility : Taxes"
  PRINT #RptHandle, "            Date of Export: "; DATE$
  PRINT #RptHandle, "  FileName Containing Data: "; FileName$
  PRINT #RptHandle, "           Field Delimiter:  ]"
  PRINT #RptHandle,
  PRINT #RptHandle, " 1 Customer Acct Number"
  PRINT #RptHandle, " 2 First Name"
  PRINT #RptHandle, " 3 Last Name"
  PRINT #RptHandle, " 4 Address Line 1"
  PRINT #RptHandle, " 5 Address Line 2"
  PRINT #RptHandle, " 6 City"
  PRINT #RptHandle, " 7 State"
  PRINT #RptHandle, " 8 Zip Code"
  PRINT #RptHandle, " 9 Customer Soc Sec #"
  IF DetailFlag THEN
   PRINT #RptHandle, "10 Real Property PIN"
   PRINT #RptHandle, "11 Real Prop Desc Line 1"
   PRINT #RptHandle, "12 Real Prop Desc Line 2"
   PRINT #RptHandle, "13 Real Prop Desc Line 3"
   PRINT #RptHandle, "14 Real Prop Value"
   PRINT #RptHandle, "15 Real Prop Senior Citizen Exemption"
   PRINT #RptHandle, "16 Mortage Code"
   PRINT #RptHandle, "17 Map "
   PRINT #RptHandle, "18 Block"
   PRINT #RptHandle, "19 Lot"
   PRINT #RptHandle, "20 Personal Property Pin #"
   PRINT #RptHandle, "21 Personal Value"
   PRINT #RptHandle, "22 Exemption Value"
   PRINT #RptHandle, "23 Desc Line 1"
   PRINT #RptHandle, "24 Desc Line 2"
   PRINT #RptHandle, "25 Desc Line 3"
   PRINT #RptHandle, "26 Desc Line 4"
   PRINT #RptHandle, "27 Desc Line 5"
  END IF

  PRINT #RptHandle, FF$
  RETURN

GetNameIndex:
  REDIM Array(1 TO NumOfTaxRecs) AS Struct
  FOR Cnt = 1 TO NumOfTaxRecs
    GET TaxFile, Cnt, TaxCustRec(1)
    Array(Cnt).who = UCASE$(TaxCustRec(1).SName) + " "
    Array(Cnt).RecNum = Cnt
  NEXT

  'Sort Them Here
  SortT Array(1), NumOfTaxRecs, 0, LEN(Array(1)), 0, 14
  RETURN

OpenTaxFile:
  TaxFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #TaxFile LEN = LEN(TaxCustRec(1))
  NumOfTaxRecs = LOF(TaxFile) / LEN(TaxCustRec(1))
  RETURN
OpenPersFile:
  PersTaxFile = FREEFILE
  OPEN "TAXPERS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PersTaxFile LEN = LEN(PersRec(1))
  NumOfPersRecs = LOF(PersTaxFile) / LEN(PersRec(1))
  RETURN
OpenPropFile:
  PropTaxFile = FREEFILE
  OPEN "TAXPROP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS #PropTaxFile LEN = LEN(PropertyRec(1))
  NumOfPropRecs = LOF(PropTaxFile) / LEN(PropertyRec(1))
  RETURN


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
  
  MChoice$(1) = "Add a New Customer"
  MChoice$(2) = "Edit Existing Customer"
  MChoice$(3) = "Customer Transaction History "
  MChoice$(4) = "Export Customer Data"
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
      REDIM TaxCust(1) AS TaxCustType
      TaxCustLen = LEN(TaxCust(1))
      NumOfCust& = FileSize&(TaxCustFile) / TaxCustLen
'      IF NumOfCust& > 2 THEN                   'UnRem for Demo Only
'        ok = MsgBox("TAX", "DEMOONLY")
'      ELSE
        AddEditTaxCust 0
'      END IF
    CASE 2
      ClearBack
      LookUp RecNo&, "Customer", False, True, False
      IF RecNo& > 0 THEN
        AddEditTaxCust RecNo&
      END IF
    CASE 3
      CustHistoryRpt
    CASE 4
      ExportData
    CASE 5
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
END SUB

