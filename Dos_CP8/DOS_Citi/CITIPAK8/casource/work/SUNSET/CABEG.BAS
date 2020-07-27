DEFINT A-Z
DECLARE SUB CABegMenu ()
DECLARE SUB CACustomerMenu ()
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
DECLARE SUB DisplayCaScrn (ScrnName$)
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

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'CACust.BI'

  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE

  
  CONST False = 0, True = NOT False
  
  CrLf$ = CHR$(13) + CHR$(10)

   CABegMenu

  RUN "camenu"

SUB AddEditTaxCust (RecNo&)

  REDIM CaCust(1) AS CACUSTType
  REDIM CaTrans(1)  AS CABegTransType
  SHARED Choice$()

  REDIM Choice$(0 TO 4, 0)
  
  REDIM ScrnArray(0)
  
  LibName$ = "TAX"
  ScrnName$ = "CABEGTR1"
  
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

  
  Frm(1).PageDisable = True
  Frm(1).StayOnField = True
  
  IF RecNo& > 0 THEN
    GOSUB LoadCustInfo
    Form$(1, 0) = QPTrim$(STR$(RecNo&))
  END IF
  
  '--Set screen number to one and display screen
  Scr = 1
  Action = 1
  FirstTime = True
  
  DisplayCaScrn ScrnName$
  
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
        ExitFlag = True
      END IF
      
    CASE EscKey
      IF EditedFlag THEN
        SaveFlag = PromptSaveData
        SELECT CASE SaveFlag
        CASE True               'user wants to save
          StuffBuf CHR$(0) + CHR$(ABS(F10Key))
        CASE False              'user wants to abandon
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
        CASE 15 TO 25           '--F4
          PressButton F4KEY, 22, 15, 25
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

AddEditExit:
EXIT SUB
  

CheckCustInfo:
  
  CustInfoOK = True             'assume all is well
  
  FOR zz = 2 TO 5
    IF LEN(QPTrim$(Form$(zz, 0))) = 0 THEN
      Frm(1).FldNo = zz
      CustInfoOK = False        'all is not well
      GOTO CustInfoExit
    END IF
  NEXT
  
  
CustInfoExit:
  IF NOT CustInfoOK THEN
    SaveScrn ScrnArray()
    DisplayCaScrn "ERRSCRN1"
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
  DisplayCaScrn "UPDATDSK"
  GOSUB SaveTaxCustRec
  RestScrn ScrnArray()
  DisplayCaScrn "UPDATEOK"
  WaitForAction
  RestScrn ScrnArray()
  Action = 1
  EditedFlag = False
RETURN
  
  
SaveTaxCustRec:
  
  
  CaTrLen = LEN(CaTrans(1))  'Length of Trans Record Structure
  FormLen = LEN(Form$(0, 0))    'Length of Cust Form$(0,0) Element

  CaTrFile = FREEFILE
  OPEN "CABEGTR.DAT" FOR RANDOM SHARED AS CaTrFile LEN = CaTrLen
  TrNumOfRecs& = LOF(CaTrFile) \ CaTrLen
  TrRecNo& = TrNumOfRecs& + 1

 'Update Customer Pointer & Running Balances
  CaCust(1).BegTrans = TrRecNo&
 'Figure Running Balance and Store
  CaCust(1).RPBStreet = CaCust(1).RPBStreet - Value(Form$(5, 0), ecode)
  CaCust(1).RPBSide = CaCust(1).RPBSide - Value(Form$(9, 0), ecode)
  CaCust(1).RPBDrive = CaCust(1).RPBDrive - Value(Form$(13, 0), ecode)
  CaCust(1).RPBDemo = CaCust(1).RPBDemo - Value(Form$(17, 0), ecode)
  CaCust(1).RPBWater = CaCust(1).RPBWater - Value(Form$(21, 0), ecode)
  CaCust(1).RPBSewer = CaCust(1).RPBSewer - Value(Form$(25, 0), ecode)
  CaCust(1).IStreet = (Value(Form$(4, 0), ecode) + Value(Form$(6, 0), ecode))
  CaCust(1).ISide = (Value(Form$(8, 0), ecode) + Value(Form$(10, 0), ecode))
  CaCust(1).IDrive = (Value(Form$(12, 0), ecode) + Value(Form$(14, 0), ecode))
  CaCust(1).IDemo = (Value(Form$(16, 0), ecode) + Value(Form$(18, 0), ecode))
  CaCust(1).IWater = (Value(Form$(20, 0), ecode) + Value(Form$(22, 0), ecode))
  CaCust(1).ISewer = (Value(Form$(24, 0), ecode) + Value(Form$(26, 0), ecode))
  CaCust(1).RIBStreet = Value(Form$(4, 0), ecode)
  CaCust(1).RIBSide = Value(Form$(8, 0), ecode)
  CaCust(1).RIBDrive = Value(Form$(12, 0), ecode)
  CaCust(1).RIBDemo = Value(Form$(16, 0), ecode)
  CaCust(1).RIBWater = Value(Form$(20, 0), ecode)
  CaCust(1).RIBSewer = Value(Form$(24, 0), ecode)
  PUT CaFile, RecNo&, CaCust(1)

  'Copy Cust info from Form$(0,0) to the Cust Rec Structure
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(CaCust(1)), VARPTR(CaCust(1)), FormLen, 0
  CaTrans(1).Pad = ""
  PUT CaTrFile, TrRecNo&, CaTrans(1)
  CLOSE CaFile
  CLOSE CaTrFile
  ERASE CaTrans
  ERASE CaCust
  RETURN
  
LoadCustInfo:
  
  CaRecLen = LEN(CaCust(1))
  Form$(0, 0) = SPACE$(CustFormLen)
  CaFile = FREEFILE
  OPEN "CAMASTER.DAT" FOR RANDOM SHARED AS CaFile LEN = CaRecLen
  GET CaFile, RecNo&, CaCust(1)
   IF CaCust(1).STRAMT < 0 THEN CaCust(1).STRAMT = 0
   IF CaCust(1).SIDEAMT < 0 THEN CaCust(1).SIDEAMT = 0
   IF CaCust(1).DRIVEAMT < 0 THEN CaCust(1).DRIVEAMT = 0
   IF CaCust(1).DEMOAMT < 0 THEN CaCust(1).DEMOAMT = 0
   IF CaCust(1).WATERAMT < 0 THEN CaCust(1).WATERAMT = 0
   IF CaCust(1).SEWERAMT < 0 THEN CaCust(1).SEWERAMT = 0
  PUT CaFile, RecNo&, CaCust(1)

 'Do Not Allow But One Beg Trans
  IF CaCust(1).BegTrans > 0 THEN
    SaveScrn ScrnArray()
    DisplayCaScrn "ERRSCRN1"
    QPrintRC "Beginning Transaction File Exits!", 10, 24, -1
    QPrintRC "Only One Beginning Transaction Allowed", 12, 21, -1
    QPrintRC "Per Customer!", 13, 34, -1
    WaitForAction
    RestScrn ScrnArray()
    CLOSE
    EXIT SUB
  END IF

  IF CaCust(1).LastTrans > 0 THEN
    SaveScrn ScrnArray()
    DisplayCaScrn "ERRSCRN1"
    QPrintRC "Other Transactions Already On File!", 10, 24, -1
    QPrintRC "Beginning Transaction Not Allowed Now", 12, 21, -1
    
    WaitForAction
    RestScrn ScrnArray()
    CLOSE
    EXIT SUB
  END IF

  Form$(2, 0) = CaCust(1).CUSTNAME
  Form$(3, 0) = QPTrim$(STR$(CaCust(1).STRAMT))
  Form$(4, 0) = "0"
  Form$(5, 0) = "0"
  Form$(6, 0) = "0"
  Form$(7, 0) = QPTrim$(STR$(CaCust(1).SIDEAMT))
  Form$(8, 0) = "0"
  Form$(9, 0) = "0"
  Form$(10, 0) = "0"
  Form$(11, 0) = QPTrim$(STR$(CaCust(1).DRIVEAMT))
  Form$(12, 0) = "0"
  Form$(13, 0) = "0"
  Form$(14, 0) = "0"
  Form$(15, 0) = QPTrim$(STR$(CaCust(1).DEMOAMT))
  Form$(16, 0) = "0"
  Form$(17, 0) = "0"
  Form$(18, 0) = "0"
  Form$(19, 0) = QPTrim$(STR$(CaCust(1).WATERAMT))
  Form$(20, 0) = "0"
  Form$(21, 0) = "0"
  Form$(22, 0) = "0"
  Form$(23, 0) = QPTrim$(STR$(CaCust(1).SEWERAMT))
  Form$(24, 0) = "0"
  Form$(25, 0) = "0"
  Form$(26, 0) = "0"
  RETURN

END SUB

SUB CABegMenu

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 3)

  MChoice$(1) = "Add a New Transaction      "
  MChoice$(2) = "Exit to OS"

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

    TitleBox 2, Col, MaxLen + 3, "Beginning Transaction Menu", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
     CASE 1
      ClearBack
      LookUp RecNo&, "Customer", False, True, False
      IF RecNo& > 0 THEN
        AddEditTaxCust RecNo&
      END IF
    CASE 2
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
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
  
  DisplayCaScrn FormName$
  
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

