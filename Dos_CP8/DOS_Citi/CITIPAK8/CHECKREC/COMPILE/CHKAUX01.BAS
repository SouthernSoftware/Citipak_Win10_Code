DEFINT A-Z
DECLARE SUB OpenBankFile (BankRecLen%, BankFile%, NumBanks%)
DECLARE SUB DisplayCRScrn (ScrnName$)
DECLARE SUB MakeZipCodeIndex (IndexText$)
DECLARE FUNCTION GetZipEDigit$ (Zip$)
DECLARE FUNCTION IsDeleted% (AcctNo&)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB AddEditLocation (RecNo&, FromFlag%)
DECLARE SUB ShowSearchWheel (BYVAL Row%, BYVAL Col%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB BlockClear ()
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, X%, Y%)
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayScrn (BTmp%(), Element%, MonoCode%, WipeType%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB ExplodBox (UlRow%, UlCol%, BrRow%, BrCol%, Scr(), El%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB GetCursor (X, Y, Button)
DECLARE SUB HideCursor ()
DECLARE SUB BlockClear ()
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB MScrnRest (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnSave (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB Pause3 (MillaSecs%, ProcessorSpeed&)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
'DECLARE SUB MQPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB ShowCursor ()
DECLARE SUB SaveScrn (Array())
DECLARE SUB RestScrn (Array())
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)
  
DECLARE FUNCTION Chk4DupeBookSeqNum (Book$, SeqNum$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION QPValI% (Number$)
  
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION FmtBook$ (Book$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION FindRateTbl% (RATECODE$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION FileSize& (FileName$)
  
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION ConvDateStr$ (D$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FudgeFactor& ()
DECLARE FUNCTION NovellThere% ()
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (N#)
DECLARE SUB Search4Cust (SEARCH$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB Search4LNum (LNum$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB Search4Meter (MeterNum$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB Search4SAddr (SAddr$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)

  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  '$INCLUDE: 'DefCnf.bi'      'Defines a TYPE for monitor and color info..
  '$INCLUDE: 'FORMEDIT.bi'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'UBTRANS.bi'
  '$INCLUDE: 'UBSETUP.bi'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'setcnf.bi'
  
  CONST False = 0, True = NOT False
  CONST Wheel$ = "|/Ä\"
  
  DIM SHARED PctC(1) AS STRING * 3
  DIM SHARED DebugFlag AS INTEGER


  '$INCLUDE: 'GL.BI'                        'GL File Types

FUNCTION AskAbandonPrint%
  
  STATIC BeenHere, Escape$
  
  IF NOT BeenHere THEN
    BeenHere = True
    Escape$ = CHR$(27)
  END IF
  
  Ky$ = INKEY$  'ButNo,DnFlg,NoPresses,col,row
  ButtonPress 1, N, MooseButton, MCol, MRow     ' ----- Check Mouse activity
  IF N AND 2 THEN               'if it was the right moose button and
    DO          'if they are still holding it down then
      GetCursor 0, 0, NewStatus 'wait till they let go of the button
    LOOP WHILE NewStatus
    ButtonPress 0, 0, 0, 0, 0   'this has the effect of clearing all
    Ky$ = Escape$               'moose buttons.
  END IF
  
  IF Ky$ <> Escape$ THEN
    EXIT FUNCTION
  END IF
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  ActMRow = 14
  BlockClear
'findme
  'DisplayUBScrn "ABORTJOB"
  
  DO
    
    Get.Moose.OR.Key Ky$, MooseButton, MRow, MCol
    
    IF MooseButton THEN
      MRow = (MRow \ 8) + 1     'Convert MCol, MRow to Row and Col.
      MCol = (MCol \ 8) + 1
      IF MRow = ActMRow THEN
        SELECT CASE MCol
        CASE 28 TO 39           'Cancel printing
          PressButton EscKey, ActMRow, 28, 39
        CASE 42 TO 55           'let it rip
          PressButton 13, ActMRow, 42, 55
        END SELECT
      END IF
    END IF
    
    IF LEN(Ky$) THEN
      IF Ky$ = Escape$ THEN
        AskAbandonPrint% = True
      ELSE
        AskAbandonPrint% = False
      END IF
      EXIT DO
    END IF
  LOOP
  
  RestScrn TempScrn()
  ERASE TempScrn
  
END FUNCTION

SUB CursorOff STATIC
  LOCATE , , 0
END SUB

SUB GetBankList (BankList$(), NumBanks, EntryFlag)

  REDIM Bank(1) AS BankRecType
  BankRecLen = LEN(Bank(1))

  'OpenBankFile BankRecLen, BankFile, NumBanks

  BankFile = FREEFILE
  OPEN "GLBANK.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS BankFile LEN = BankRecLen
  NumBanks = LOF(BankFile) \ BankRecLen
  
  IF NumBanks = 0 THEN
    CLOSE BankFile
    EXIT SUB
  ELSE
    IF EntryFlag THEN
      REDIM BankList$(1 TO NumBanks + 2)
      BankList$(1) = "All Banks "
      BankList$(2) = " 0 No Bank Code"
      DidCnt = 2
    ELSE
      REDIM BankList$(1 TO NumBanks + 1)
      BankList$(1) = " 0 No Bank Code"
      DidCnt = 1
    END IF
  END IF
  
  BankNum$ = SPACE$(2)

  FOR Cnt = 1 TO NumBanks
    GET BankFile, Cnt, Bank(1)
    IF NOT Bank(1).Deleted THEN
      DidCnt = DidCnt + 1
      RSET BankNum$ = QPTrim$(STR$(Bank(1).BankNum))
      BankList$(DidCnt) = BankNum$ + " " + QPTrim$(Bank(1).BankName)
      LSET BankNum$ = ""
    ELSE

    END IF
  NEXT
  NumBanks = DidCnt
  CLOSE BankFile
  
END SUB

FUNCTION PromptSaveData%
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  LibName$ = "CHECKREC"
  SaveFlag = 2
  
  FormName$ = "SAVE1ST"
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
'findme

  DisplayCRScrn FormName$
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE F0Key
      SaveFlag = True
    CASE EscKey
      SaveFlag = 1
    CASE 88, 120                'X Key
      SaveFlag = False
    END SELECT
    
  LOOP WHILE SaveFlag = 2       'proper key not set
  
  PromptSaveData = SaveFlag
  CursorOff
  
  RestScrn TempScrn()
  
  ERASE TempScrn, Form$, Fld, Frm
  
END FUNCTION

SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%) STATIC
  RSET PctC(1) = QPStrI$(INT((RecNo / NumOfRecs) * 100))
  'HideCursor
  QPrintRC PctC(1), 13, 40, Cnf.HiLite
  'ShowCursor
  '  QPrintRC STR$(FRE("")), 25, 1, Cnf.HiLite
END SUB

SUB ShowProcessingScrn (RptTitle$)

  TitleRow = 9
  TitleCol = 40 - (LEN(RptTitle$) \ 2) + 1

  'CursorOff
  'BlockClear
  DisplayCRScrn "MENUBAK"
  DisplayCRScrn "PRORPT"

  HideCursor
  QPrintRC RptTitle$, TitleRow, TitleCol, 126
  QPrintRC "Processing:    % Completed.", 13, 28, Cnf.HiLite
  ShowCursor
END SUB

