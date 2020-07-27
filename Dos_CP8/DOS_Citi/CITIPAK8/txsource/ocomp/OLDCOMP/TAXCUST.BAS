DEFINT A-Z
DECLARE SUB SaveTaxCustRec (Form$(), RecNo&)
DECLARE SUB ClearScrn ()
DECLARE SUB HideCursor ()
DECLARE SUB AddEditTaxCust (RecNo&)
DECLARE SUB ClearBack ()
DECLARE SUB DisplayTaxScrn (ScrnName$)

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
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
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

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'TaxCust.BI'

  CONST False = 0, True = NOT False
  
  CrLf$ = CHR$(13) + CHR$(10)

SUB AddEditTaxCust (RecNo&)

  LibName$ = "TAX"
  ScrnName$ = "TAXCUST1"
  REDIM ScrnArray(0)

  '--define the multi-choice fields
  SHARED Choice$()
  REDIM Choice$(0 TO 4, 0)

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
  FOR f = 1 TO NumFlds
    TFormLen = TFormLen + Fld(f).StorLen
  NEXT

  '--Must fill Form$(0, 0) to the correct len, before
  'using the BCopy function.
  Form$(0, 0) = SPACE$(TFormLen)

  IF RecNo& > 0 THEN
    GOSUB LoadCustInfo
    UnPackBuffer 0, 0, Form$(), Fld()
    LSET Form$(1, 0) = QPTrim$(STR$(RecNo&))
  ELSE
    Frm(1).PageDisable = True
  END IF

  '--Set screen number to one and display screen
  Scr = 1
  Action = 1
  FirstTime = True

  DisplayTaxScrn ScrnName$
  
  ShowCursor

  'Frm(1).StayOnField = True

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
      SaveScrn ScrnArray()
      DisplayTaxScrn "UPDATDSK"
      SaveTaxCustRec Form$(), RecNo&
      'UBLog "UPDATED:" + STR$(RecNo&) + " " + TempName$
      RestScrn ScrnArray()
      DisplayTaxScrn "UPDATEOK"
      WaitForAction
      ExitFlag = True
      RestScrn ScrnArray()
    CASE F3Key

    CASE F4KEY

    CASE F6KEY

    CASE F7KEY

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
        CASE 27 TO 37           '--F6 Cons Hist
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
  GOTO ExitAddEdit

LoadCustInfo:

  REDIM TAXCust(1) AS TaxCustType
  TaxRecLen = LEN(TAXCust(1))
  Form$(0, 0) = SPACE$(TFormLen)
  TAXFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM SHARED AS TAXFile LEN = TaxRecLen
  GET TAXFile, RecNo&, TAXCust(1)
  CLOSE
  BCopy VARSEG(TAXCust(1)), VARPTR(TAXCust(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), TFormLen, 0
  ERASE TAXCust

RETURN

ExitAddEdit:

END SUB

SUB SaveTaxCustRec (Form$(), RecNo&)

  REDIM TAXCust(1)  AS TaxCustType

  TaxCustLen = LEN(TAXCust(1))              'Length of Cust Record Structure
  FormLen = LEN(Form$(0, 0))    'Length of Cust Form$(0,0) Element

  TAXFile = FREEFILE
  OPEN "TAXCUST.DAT" FOR RANDOM SHARED AS TAXFile LEN = TaxCustLen
  NumOfRecs& = LOF(TAXFile) \ TaxCustLen

  IF RecNo& > 0 THEN
    GET TAXFile, RecNo&, TAXCust(1)
  ELSEIF RecNo& = 0 THEN
    RecNo& = NumOfRecs& + 1
    TAXCust(1).Acct = RecNo&
  END IF

  'Copy Cust info from Form$(0,0) to the Cust Rec Structure
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(TAXCust(1)), VARPTR(TAXCust(1)), FormLen, 0

  'stamp this record with the file version number
  TAXCust(1).FileVer = CustRecVerNO

  PUT TAXFile, RecNo&, TAXCust(1)
  CLOSE TAXFile

  ERASE TAXCust

END SUB

SUB TAXCustomerMenu
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 9)
  
  MChoice$(1) = " Add a New Customer       "
  MChoice$(2) = " Edit Existing Customer   "
  MChoice$(3) = " Delete Customer Account  "
  MChoice$(4) = " Exit to OS"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2) + 1
  Col = ((80 - MaxLen) \ 2) - 2
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0

    ClearBack
    'ClearBack
    
    TitleBox 2, Col, MaxLen + 3, "TAX Customer Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      AddEditTaxCust 0
      
    CASE 2
      ClearBack
      'RecNo& = 1
      LookUp RecNo&, "Customer", 2, True, False
      IF RecNo& > 0 THEN
        AddEditTaxCust RecNo&
      END IF
    CASE 3
    CASE 4
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
END SUB

