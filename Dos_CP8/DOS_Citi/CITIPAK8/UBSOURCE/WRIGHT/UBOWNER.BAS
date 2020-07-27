DEFINT A-Z

DECLARE FUNCTION Compare3% (BYVAL Seg1, BYVAL Adr1, BYVAL Seg2, BYVAL Adr2, NumBytes)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FmtBook$ (BOOK$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (Number$)

DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB EditCustomer ()
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB InsertT (SEG StartElement AS ANY, ElSize%, NumEls%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
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
DECLARE SUB UBLog (Text$)
DECLARE SUB WaitForAction ()
DECLARE SUB MPaintBox (TRow%, LCol%, BRow%, RCol%, PColor%)
DECLARE FUNCTION Unique$ (Path$)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  ''$INCLUDE: 'PageInfo.BI'                  'Form Page info
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  ''$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'newcust.bi'
  ''$INCLUDE: 'ubrate.BI'
  ''$INCLUDE: 'ubTrans.BI'
  '$INCLUDE: 'ubowner.BI'

  CONST False = 0, True = NOT False

SUB EditOwnerInfo (RecNo&, OTempName$)
  
  REDIM TempScr(0)
  SaveScrn TempScr()

  LibName$ = "UB"
  ScrnName$ = "UBOWNER"

  '--Get the total number of fields from all pages
  NumFlds = NumFlds + LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM OwnFrm(1) AS FormInfo
  REDIM OwnForm$(NumFlds, 2)
  REDIM OwnFld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, OwnFld(), OwnForm$(), ErrCode

  '--Increment StartEl to next

  FOR F = 1 TO NumFlds
    TFormLen = TFormLen + OwnFld(F).StorLen
  NEXT

  OwnForm$(0, 0) = SPACE$(TFormLen)

  IF RecNo& > 0 THEN
    GOSUB LoadOwnerInfo
    UnPackBuffer 0, 0, OwnForm$(), OwnFld()
  END IF

  Action = 1
  FirstTime = True

  MPaintBox 2, 5, 22, 75, 8

  DisplayUBScrn ScrnName$

  DO

    EditForm OwnForm$(), OwnFld(), OwnFrm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE OwnFrm(1).KeyCode
    CASE F10Key 'F10
      IF RecNo& > 0 THEN
        GOSUB SaveOwnerInfo
      ELSE
        'need to figure out how to deal with adding new customer
        'the aborting edit
      END IF
      ExitFlag = True
    CASE EscKey
      RestScrn TempScr()
      ExitFlag = True
    END SELECT

    '--check for mouse clicks on buttons not attached to the form
    IF OwnFrm(1).Presses THEN
      SELECT CASE OwnFrm(1).MRow
      CASE 19   'Look for the f10 or esc button
        SELECT CASE OwnFrm(1).MCol
        CASE 31 TO 40           '--F10 Save Button
          PressButton F10Key, 19, 31, 40
        CASE 42 TO 53           '--F10 Save Button
          PressButton F10Key, 19, 42, 53
        END SELECT
      END SELECT                'row
    END IF

'
'    '--Check screen page
  LOOP UNTIL ExitFlag

  RestScrn TempScr()

  ERASE TempScr, OwnFrm, OwnForm$, OwnFld


EXIT SUB

SaveOwnerInfo:
  REDIM UBOwnerRec(1) AS UBOwnerRecType
  UBOwnerRecLen = LEN(UBOwnerRec(1))

  BCopy SSEG(OwnForm$(0, 0)), SADD(OwnForm$(0, 0)), VARSEG(UBOwnerRec(1)), VARPTR(UBOwnerRec(1)), TFormLen, 0

  UBFile = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBFile LEN = UBOwnerRecLen
  PUT UBFile, RecNo&, UBOwnerRec(1)
  CLOSE


RETURN

LoadOwnerInfo:

  REDIM UBOwnerRec(1) AS UBOwnerRecType

  UBOwnerRecLen = LEN(UBOwnerRec(1))

  OwnForm$(0, 0) = SPACE$(TFormLen)

  UBFile = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBFile LEN = UBOwnerRecLen
  GET UBFile, RecNo&, UBOwnerRec(1)
  CLOSE

  BCopy VARSEG(UBOwnerRec(1)), VARPTR(UBOwnerRec(1)), SSEG(OwnForm$(0, 0)), SADD(OwnForm$(0, 0)), TFormLen, 0

  ERASE UBOwnerRec


RETURN

END SUB

