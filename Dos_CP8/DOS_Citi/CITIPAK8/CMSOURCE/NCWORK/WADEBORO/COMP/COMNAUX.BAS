DEFINT A-Z
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
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
DECLARE SUB DisplayCMScrn (ScrnName$)
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
DECLARE SUB ShowCursor ()
DECLARE SUB SaveScrn (array())
DECLARE SUB RestScrn (array())
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetA (Handle%, SEG Dest AS ANY, NumBytes%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)
  
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION QPValI% (Number$)
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION FmtBook$ (Book$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION FindRateTbl% (RATECODE$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION AskSavePayment% (ReceiptFlag$)
DECLARE SUB UBLog (Text$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION ConvDateStr$ (d$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (N#)
DECLARE SUB Search4Cust (SEARCH$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB Search4LNum (LNum$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB Search4Meter (MeterNum$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB Search4SAddr (SAddr$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB KillFile (File2Kill$)
DECLARE FUNCTION GetCustMeterType% (UBCustRec() AS ANY, ThisMeter%)
DECLARE FUNCTION GetDEFPort% ()

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
  '$INCLUDE: 'CMFILES.BI'

  CONST False = 0, True = NOT False
  CONST Wheel$ = "|/Ä\"
  
  DIM SHARED PctC(1) AS STRING * 3
  DIM SHARED DebugFlag AS INTEGER

FUNCTION AskSavePayment% (DefaultFlag$)
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  LibName$ = "CM"
  ScrnName$ = "ASKR2PST"
  
  '--Initialize the form name array
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  DisplayCMScrn ScrnName$
  
  ShowCursor
  
  Action = 1
  FirstTime = True
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = NOT FirstTime
      LSET Form$(1, 0) = DefaultFlag$
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE EscKey
      AskSavePayment% = False
      EXIT DO
    CASE F10Key
      Receipt$ = Form$(1, 0)
      IF Receipt$ = "Y" THEN
        AskSavePayment% = 1
      ELSE
        AskSavePayment% = True
      END IF
      EXIT DO
    END SELECT
    
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 14
        SELECT CASE Frm(1).MCol
        CASE 27 TO 40           'ESC Cancel button
          PressButton 27, 14, 27, 40
        CASE 41 TO 56           'F10 Save Button
          PressButton -68, 14, 41, 56
        END SELECT
      END SELECT
    END IF
  LOOP
  
  RestScrn TempScrn()
  
END FUNCTION

SUB ClearScrn STATIC
  WazzWind 1, 1, 25, 80, 7, 0, 0
END SUB

SUB CursorOff STATIC
  LOCATE , , 0
END SUB

SUB CustMessageSystem (RecNo&)
  
  CustRec& = RecNo&
  
  REDIM ScrnArray(0)
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBMessRec(1) AS UBMessRecType
  UBMessRecLen = LEN(UBMessRec(1))
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, CustRec&, UBCustRec(1)
  CLOSE
  
  LibName$ = "UB"
  ScrnName$ = "UBCUSMES"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True
  
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  FirstTime = True
  
  Action = 1
  
  DisplayUBScrn ScrnName$
  QPrintRC STR$(CustRec&), 3, 20, -1
  QPrintRC UBCustRec(1).CustName, 4, 20, -1
  QPrintRC UBCustRec(1).Status, 3, 67, -1
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      GOSUB LoadMessageInfo
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F3Key
      GOSUB ClearRecord
      GOSUB ClearForm
      Action = 1
    CASE F10Key
      SaveScrn ScrnArray()
      DisplayUBScrn "UPDATDSK"
      GOSUB SaveRecord
      RestScrn ScrnArray()
      DisplayUBScrn "UPDATEOK"
      WaitForAction
      ExitFlag = True
      RestScrn ScrnArray()
      Done = True
    CASE ESC
      EXIT SUB
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done
  
ExitMessageInquiry:
  EXIT SUB
  '***************
  
LoadMessageInfo:
  MessageRecord = UBCustRec(1).MessageRec
  IF MessageRecord > 0 THEN
    UBMess = FREEFILE
    OPEN "UBMESAGE.DAT" FOR RANDOM SHARED AS UBMess LEN = UBMessRecLen
    GET UBMess, MessageRecord, UBMessRec(1)
    CLOSE
    BCopy VARSEG(UBMessRec(1)), VARPTR(UBMessRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), UBMessRecLen, 0
    CALL UnPackBuffer(0, 0, Form$(), Fld())
  END IF
RETURN
  
SaveRecord:
  UBMess = FREEFILE
  OPEN "UBMESAGE.DAT" FOR RANDOM SHARED AS UBMess LEN = UBMessRecLen
  IF MessageRecord = 0 THEN
    MessageRecord = LOF(UBMess) / LEN(UBMessRec(1)) + 1
  END IF
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBMessRec(1)), VARPTR(UBMessRec(1)), UBMessRecLen, 0
  PUT UBMess, MessageRecord, UBMessRec(1)
  CLOSE
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, CustRec&, UBCustRec(1)
  UBCustRec(1).MessageRec = MessageRecord
  PUT UBCust, CustRec&, UBCustRec(1)
  CLOSE
RETURN
  
ClearRecord:
  IF MessageRecord > 0 THEN
    REDIM UBMessRec(1) AS UBMessRecType
    UBMess = FREEFILE
    OPEN "UBMESAGE.DAT" FOR RANDOM SHARED AS UBMess LEN = UBMessRecLen
    PUT UBMess, MessageRecord, UBMessRec(1)
    CLOSE
  END IF
RETURN
  
ClearForm:
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
RETURN
  
  
END SUB

FUNCTION FmtBook$ (Book$)
  Book$ = QPTrim$(Book$)
  BookLen = LEN(Book$)
  
  SELECT CASE BookLen
  CASE 0
    FmtBook$ = "00"
  CASE 1
    FmtBook$ = "0" + Book$
  CASE ELSE
    FmtBook$ = Book$
  END SELECT
  
END FUNCTION

FUNCTION FmtSeqN$ (SeqN$)
  
  SeqN$ = QPTrim$(SeqN$)
  SeqNLen = LEN(SeqN$)
  
  SELECT CASE SeqNLen
  CASE 0
    FmtSeqN$ = "000000"
  CASE 1 TO 5
    FmtSeqN$ = "000000"
    MID$(FmtSeqN$, (6 - SeqNLen) + 1) = SeqN$
  CASE ELSE
    FmtSeqN$ = SeqN$
  END SELECT
  
END FUNCTION

FUNCTION GetCustMeterType (UBCustRec() AS NewUBCustRecType, ThisMeter)
  
  'Meter Types
  'CONST MtrWaterOnly = 1
  'CONST MtrSewerOnly = 2
  'CONST MtrCombined = 3
  'CONST MtrElectric = 4
  'CONST MtrDemand = 5
  'CONST MtrGas = 6
  'CONST MtrTouchRead = 7
  
  LMtrType$ = QPTrim$(UBCustRec(1).LocMeters(ThisMeter).MTRType)
  LMtrTypeLen = LEN(LMtrType$)
  IF LMtrTypeLen > 0 THEN
    SELECT CASE LMtrType$
    CASE "W"
      LThisMeter = MtrWaterOnly
    CASE "S"
      LThisMeter = MtrSewerOnly
    CASE "C"
      LThisMeter = MtrCombined
    CASE "E"
      LThisMeter = MtrElectric
    CASE "D"
      LThisMeter = MtrDemand
    CASE "G"
      LThisMeter = MtrGas
    CASE "T"
      LThisMeter = MtrTouchRead
    CASE ELSE
      LThisMeter = True
    END SELECT
    GetCustMeterType = LThisMeter
  ELSE
    GetCustMeterType = 0
  END IF
  
END FUNCTION

FUNCTION GetDEFPort%
  REDIM zz(1) AS STRING * 1
  IF Exist("CMSETUP.DAT") THEN
    CMFile = FREEFILE
    OPEN "CMSETUP.DAT" FOR RANDOM AS CMFile LEN = 1
    IF LOF(CMFile) > 0 THEN
      GET CMFile, 1, zz(1)
    ELSE
      zz(1) = "1"
    END IF
    CLOSE CMFile
  ELSE
    zz(1) = "1"
  END IF
  
  GetDEFPort% = QPValI(zz(1))
  
  ERASE zz
  
END FUNCTION

'This function returns the number of customer records
FUNCTION GetNumOfCust
  REDIM TCustRec(1) AS NewUBCustRecType
  RecLen = LEN(TCustRec(1))
  CFileSize& = FileSize("UBCUST.DAT")
  GetNumOfCust = CFileSize& \ RecLen
  ERASE TCustRec
END FUNCTION

'Returns TRUE if this is a deleted account
'
FUNCTION IsDeleted% (AcctNo&)
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  FOpenS "UBCUST.DAT", C1Handle 'open data file
  FGetRTA C1Handle, UBCustRec(1), AcctNo&, UBCustRecLen
  FClose C1Handle
  IF UBCustRec(1).DelFlag <> 0 THEN
    IsDeleted% = True
  ELSE
    IsDeleted% = False
  END IF
  ERASE UBCustRec
END FUNCTION

SUB LoadUBSetUpFile (UBSetUpRec() AS UBSetupRecType, UBSetupLen)
  
  UBSetupLen = -1               'use the length as an error flag
  'assume the file is not there, or 0 bytes.
  IF Exist("UBSETUP.DAT") THEN
    FOpenS "UBSETUP.DAT", Handle                'open data file
    IF FLof&(Handle) > 0 THEN
      UBSetupLen = LEN(UBSetUpRec(1))
      FGetRTA Handle, UBSetUpRec(1), 1&, UBSetupLen
    END IF
    FClose Handle
  END IF
  
END SUB

SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
  
  STATIC SName$, AcctNum&, MeterNum$, SAddr$, LNum$
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  
  SName$ = ""
  AcctNum& = 0
  MeterNum$ = ""
  SAddr$ = ""
  LNum$ = ""

  IF INSTR(COMMAND$, "DEBUG") THEN
    DebugFlag = True
  END IF
  
  SELECT CASE QPValI(UBSetUpRec(1).DefLook)
  CASE 1
    LScrn = 1
  CASE 2
    LScrn = 2
  CASE 3
    LScrn = 3
  CASE 4
    LScrn = 4
  CASE 5
    LScrn = 5
  CASE ELSE
    LScrn = DefaultLook
  END SELECT
  
  CursorOff
  
  REDIM ScrnArray(0)
  REDIM ScrnArray2(0)
  
  SaveScrn ScrnArray()
  
  REDIM LText$(5)
  
  MScrn = 5
  
  LText$(1) = " Account Number:"
  LText$(2) = "    Search Name:"
  LText$(3) = "   Meter Number:"
  LText$(4) = "Service Address:"
  LText$(5) = "Location Number:"
  
  LibName$ = "UB"
  ScrnName$ = "LUPACCT"
  
  '--Initialize the form name array
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  Text$ = Text$ + " Look-Up"
  TextLen = LEN(Text$)
  TCol = ((80 - TextLen) \ 2)
  DisplayUBScrn ScrnName$
  
  QPrintRC Text$, 8, TCol, -1
  
  GOSUB DisplayLookupText
  
  ShowCursor
  
  Action = 1
  FirstTime = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = NOT FirstTime
      SELECT CASE LScrn
      CASE 1
        IF AcctNum& > 0 THEN
          Form$(1, 0) = QPStrL$(AcctNum&)
        END IF
      CASE 2
        Form$(1, 0) = SName$
      CASE 3
        Form$(1, 0) = MeterNum$
      CASE 4
        Form$(1, 0) = SAddr$
      CASE 5
        Form$(1, 0) = LNum$
      END SELECT
      Action = 1
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68, 13                'F10Key    Proceed with look up
      CursorOff
      SELECT CASE LScrn
      CASE 1    'account lookup
        AcctNum& = QPValL(Form$(1, 0))
        IF AcctNum& < 1 OR AcctNum& > GetNumOfCust THEN
          Ok = MsgBox%("UB.QSL", "BADACCTN")
        ELSE
          IF IsDeleted(AcctNum&) THEN
            Ok = MsgBox%("UB.QSL", "DELACCTN")
          ELSEIF ActiveOnly THEN
            CHand = FREEFILE
            OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHand LEN = UBCustRecLen
            GET #CHand, AcctNum&, UBCustRec(1)
            CLOSE CHand
            IF UBCustRec(1).Status = "A" THEN
              RecNo& = AcctNum&
              OKFlag = True
            ELSE
              Ok = MsgBox%("UB.QSL", "INACTACT")
            END IF
          ELSE
            RecNo& = AcctNum&
            OKFlag = True
          END IF
        END IF
        
      CASE 2    'Name lookup
        SName$ = LEFT$(QPTrim$(Form$(0, 0)), 10)
        IF LEN(SName$) = 0 THEN
          SName$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        Search4Cust SName$, RecNo&, CLSFlag, ActiveOnly
        IF RecNo& > 0 THEN
          OKFlag = True
        ELSEIF RecNo& = 0 THEN
          Ok = MsgBox%("UB.QSL", "NOMATCH")
        END IF
        RestScrn ScrnArray2()
        Action = 1
        'END IF
      CASE 3    'meter number
        MeterNum$ = QPTrim$(Form$(0, 0))
        IF LEN(MeterNum$) = 0 THEN
          Ok = MsgBox%("UB.QSL", "BADSEARH")
          Action = 1
          Frm(1).FldNo = 1
        ELSE
          SaveScrn ScrnArray2()
          RestScrn ScrnArray()
          Search4Meter MeterNum$, RecNo&, CLSFlag, ActiveOnly
          IF RecNo& > 0 THEN
            OKFlag = True
          ELSEIF RecNo& = 0 THEN
            Ok = MsgBox%("UB.QSL", "NOMATCH")
          END IF
          RestScrn ScrnArray2()
          Action = 1
        END IF
      CASE 4    'service address
        SAddr$ = QPTrim$(Form$(0, 0))
        IF LEN(SAddr$) = 0 THEN
          Ok = MsgBox%("UB.QSL", "BADSEARH")
          Action = 1
          Frm(1).FldNo = 1
        ELSE
          SaveScrn ScrnArray2()
          RestScrn ScrnArray()
          Search4SAddr SAddr$, RecNo&, CLSFlag, ActiveOnly
          IF RecNo& > 0 THEN
            OKFlag = True
          ELSEIF RecNo& = 0 THEN
            Ok = MsgBox%("UB.QSL", "NOMATCH")
          END IF
          RestScrn ScrnArray2()
          Action = 1
        END IF
        
      CASE 5    'Location lookup
        OK2Search = False
        LNum$ = QPTrim$(Form$(0, 0))
        DashPos = INSTR(LNum$, "-")
        IF LEN(LNum$) < 2 THEN  'OR DashPos <= 0 THEN
          Ok = MsgBox%("UB.QSL", "BADACCTN")
          Action = 1
          Frm(1).FldNo = 1
        ELSEIF DashPos > 1 THEN
          Book$ = FmtBook$(LEFT$(LNum$, DashPos - 1))
          SeqN$ = FmtSeqN$(MID$(LNum$, DashPos + 1))
          LNum$ = Book$ + "-" + SeqN$
          LSET Form$(1, 0) = LNum$
          SaveField 0, Form$(), Fld(), BadField
          Action = 1
          OK2Search = True
        ELSE
          Book$ = FmtBook$(LEFT$(LNum$, 2))
          SeqN$ = FmtSeqN$(MID$(LNum$, 3))
          LNum$ = Book$ + "-" + SeqN$
          LSET Form$(1, 0) = LNum$
          SaveField 0, Form$(), Fld(), BadField
          Action = 1
          OK2Search = True
        END IF
        IF OK2Search THEN
          SaveScrn ScrnArray2()
          RestScrn ScrnArray()
          Search4LNum LNum$, RecNo&, CLSFlag, ActiveOnly
          IF RecNo& > 0 THEN
            OKFlag = True
          ELSEIF RecNo& = 0 THEN
            Ok = MsgBox%("UB.QSL", "NOMATCH")
          END IF
          RestScrn ScrnArray2()
          Action = 1
        END IF
      END SELECT
      
    CASE -65    'F7Key
      IF LScrn < MScrn THEN
        LScrn = LScrn + 1
      ELSE
        LScrn = 1
      END IF
      LSET Form$(1, 0) = ""
      Action = 1
      FirstTime = True
      SaveField 0, Form$(), Fld(), BadField
      GOSUB DisplayLookupText
    CASE 27
      RecNo& = 0
      ExitFlag = True
    END SELECT
    
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 16
        SELECT CASE Frm(1).MCol
        CASE 22 TO 33           'ESC Cancel button
          PressButton 27, 16, 22, 33
        CASE 35 TO 45           'F7 Toggle Choice
          PressButton -65, 16, 35, 45
        CASE 47 TO 59           'F10 Save Button
          PressButton -68, 16, 47, 59
        END SELECT
      END SELECT
    END IF
    
  LOOP UNTIL ExitFlag OR OKFlag
  RestScrn ScrnArray()

  ERASE Frm, Form$, Fld, UBSetUpRec
  ERASE ScrnArray, ScrnArray2, UBCustRec
  ERASE LText$
  
  EXIT SUB
  
DisplayLookupText:
  QPrintRC LText$(LScrn), 12, 15, -1
RETURN
  
  
END SUB

SUB OpenMiscCodeFile (NumOfMiscRecs, MCFile)
  REDIM MiscCodeRec(1) AS MiscCodeRecType
  MiscCodeRecLen = LEN(MiscCodeRec(1))
  MCFile = FREEFILE
  OPEN "CMMISCCD.DAT" FOR RANDOM SHARED AS MCFile LEN = MiscCodeRecLen
  NumOfMiscRecs = LOF(MCFile) \ MiscCodeRecLen
  
END SUB

FUNCTION PromptSaveData%
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  LibName$ = "UB"
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
  
  DisplayUBScrn FormName$
  
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

'****************************************************************************
'Rounds a double precision value to nearest hundreth
'****************************************************************************
FUNCTION Round# (N#)
  Round# = INT(N# * 100 + .5000001) / 100
END FUNCTION

SUB Search4Cust (SEARCH$, RecNo&, CLSFlag%, ActiveOnly%)
  'LocationFlag%
  STATIC Choice, LastSEARCH$
  
  IF LastSEARCH$ <> SEARCH$ THEN
    LastSEARCH$ = SEARCH$
    Choice = 1
  END IF
  
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  WPos = 1
  
  DisplayUBScrn "SHOWSCRH"
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  IdxRecLen = 4 'we are using a integer
  IdxFileSize& = FileSize("UBCUSTNM.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  
  FGetAH "UBCUSTNM.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  SearchLen = LEN(SEARCH$)
  
  Match = False
  FirstRec = 1
  LastRec = IdxNumOfRecs
  
  BotOffSet = 1
  TopOffSet = IdxNumOfRecs
  
  FOpenS "UBCUST.DAT", C1Handle 'open data file
  
  'Find matching record
  
  MidRec = (LastRec + FirstRec) \ 2
  
  DO
    
    FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(MidRec).RecNum), UBCustRecLen
    UBSearchN$ = LEFT$(UBCustRec(1).SEARCH, SearchLen)
    
    ShowSearchWheel 12, 44
    'IF SEARCH$ = UBSearchN$ THEN
    IF (SEARCH$ = UBSearchN$) AND (UBCustRec(1).DelFlag = 0) THEN
      FirstMatchRec = MidRec
      IF MidRec - BotOffSet > 1 THEN
        MidRec = MidRec - 1
      ELSE
        EXIT DO
      END IF
      
    ELSEIF SEARCH$ < UBSearchN$ THEN            'lower
      TopOffSet = MidRec
      MidRec = TopOffSet - (((TopOffSet - BotOffSet) \ 2) + 1)
    ELSE        'higher
      BotOffSet = MidRec
      MidRec = BotOffSet + ((TopOffSet - BotOffSet) \ 2) + 1
      IF LMidRec = MidRec THEN
        IF FirstMatchRec > 0 THEN
          FirstMatchRec = MidRec
          TotalMatch = 1
        ELSE
          MidRec = MidRec - 1
          FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(MidRec).RecNum), UBCustRecLen
          UBSearchN$ = LEFT$(UBCustRec(1).SEARCH, SearchLen)
          'IF SEARCH$ = UBSearchN$ THEN
          IF (SEARCH$ = UBSearchN$) AND (UBCustRec(1).DelFlag = 0) THEN
            FirstMatchRec = MidRec
            TotalMatch = 1
          END IF
        END IF
        EXIT DO
      ELSE
        LMidRec = MidRec
      END IF
      IF MidRec = IdxNumOfRecs + 1 THEN
        EXIT DO
      END IF
    END IF
    IF TopOffSet = BotOffSet THEN
      EXIT DO
    END IF
  LOOP
  
  IF FirstMatchRec > 0 THEN
    MidRec = FirstMatchRec
    LastMatchRec = IdxNumOfRecs
    FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(FirstMatchRec).RecNum), UBCustRecLen
    MidRec = MidRec + 1
    DO UNTIL MidRec > IdxNumOfRecs
      FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(MidRec).RecNum), UBCustRecLen
      UBSearchN$ = LEFT$(UBCustRec(1).SEARCH, SearchLen)
      IF SEARCH$ <> UBSearchN$ THEN
        LastMatchRec = MidRec - 1
        EXIT DO
      END IF
      ShowSearchWheel 12, 44
      MidRec = MidRec + 1
    LOOP
    
    MidRec = FirstMatchRec
    FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(FirstMatchRec).RecNum), UBCustRecLen
    MidRec = MidRec - 1
    DO UNTIL MidRec = 0
      FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(MidRec).RecNum), UBCustRecLen
      UBSearchN$ = LEFT$(UBCustRec(1).SEARCH, SearchLen)
      IF SEARCH$ <> UBSearchN$ THEN
        FirstMatchRec = MidRec + 1
        EXIT DO
      ELSE
        FirstMatchRec = MidRec
      END IF
      ShowSearchWheel 12, 44
      MidRec = MidRec - 1
    LOOP
    
  ELSE
    RecNo& = 0
    FClose C1Handle
    GOTO ExitSearch
  END IF
  
  IF TotalMatch <> 1 THEN
    TotalMatch = (LastMatchRec - FirstMatchRec) + 1
  END IF
  IF NOT DebugFlag THEN
    FreeMem& = FRE(-1)
    IF FreeMem& >= 65536 THEN
      FreeMem& = 65536
    END IF
    MemNeeded& = TotalMatch * 64&
    IF MemNeeded& > FreeMem& THEN
      FClose C1Handle
      QPrintRC "Matched: " + QPStrI(TotalMatch), 24, 1, 15
      QPrintRC " Needed: " + QPStrL(MemNeeded&) + "  Free: " + QPStrL(FreeMem&), 25, 1, 15
      RecNo& = -1
      WaitForAction
      GOTO ExitSearch
    END IF
  END IF
  
  IF TotalMatch >= 1 THEN
    FOR Cnt = FirstMatchRec TO LastMatchRec
      FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(Cnt).RecNum), UBCustRecLen
      IF UBCustRec(1).DelFlag <> 0 GOTO DelSkip1
      IF NOT ActiveOnly OR (ActiveOnly AND (UBCustRec(1).Status = "A")) THEN
        DCnt = DCnt + 1
        StartSortOffSet = DCnt
        EndSortOffSet = DCnt
        REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
        LSET MChoice(DCnt).V = LEFT$(QPTrim$(UBCustRec(1).CUSTNAME), 20)
        MID$(MChoice(DCnt).V, 22, 22) = LEFT$(QPTrim$(UBCustRec(1).SERVADDR), 25)
        MID$(MChoice(DCnt).V, 48, 9) = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
        MID$(MChoice(DCnt).V, 59, 1) = UBCustRec(1).Status
        MID$(MChoice(DCnt).V, 61) = MKL$(CLNG(IdxBuff(Cnt).RecNum))
        IF StartSortOffSet < EndSortOffSet THEN
          ElmCnt = (EndSortOffSet - StartSortOffSet) + 1
          SortT MChoice(StartSortOffSet), ElmCnt, 0, 64, 47, 9
          '      (Element(1), NumElm, Dir, SSize%, MOffSet, MSize%)
        END IF
      END IF
DelSkip1:
      ShowSearchWheel 12, 44
    NEXT
    FClose C1Handle
    
    IF DCnt = 0 THEN
      RecNo& = 0
      GOTO ExitSearch
    END IF
    
    'FClose L1Handle
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Customer/Owner       Service Address        Location No.  S"
    '--Find max menu width
    
    '--Center Menu within Screen
    
    Row = 4
    Col = ((80 - 60) \ 2) - 1
    
    IF CLSFlag THEN
      Row = 4
      BoxBot = 17               'limit the box length
      BlockClear
    ELSE
      Row = 6
      BoxBot = 14               'limit the box length to go no lower than line 20
      RestScrn ScrnArray()
    END IF
    
    LOCATE Row, Col, 0
    
    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Matched:" + STR$(DCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
    LOOP UNTIL RecNo& > 0
  ELSE
    RecNo& = IdxBuff(FirstMatchRec).RecNum
    FClose C1Handle
    'FClose L1Handle
  END IF
  
ExitSearch:
  'cls
  'Shell
  RestScrn ScrnArray()
  
  ERASE ScrnArray, UBCustRec, IdxBuff, MChoice
  
END SUB

'SUB Search4Cust (SEARCH$, RecNo&, CLSFlag%, ActiveOnly%)
'
'  STATIC Choice, LastSEARCH$
'
'  IF LastSEARCH$ <> SEARCH$ THEN
'    LastSEARCH$ = SEARCH$
'    Choice = 1
'  END IF
'
'  REDIM ScrnArray(0)
'  SaveScrn ScrnArray()
'
'  DisplayUBScrn "SHOWSCRH"
'
'  REDIM MChoice(1 TO 1) AS FLen2
'  REDIM UBCustRec(1) AS NewUBCustRecType
'
'  MaxBlockCnt = 1024            'Max Buff size: 16384   (1024*16)
'  REDIM UBCustSN(1 TO MaxBlockCnt) AS nUBCustReIndexRecType
'
'  UBCustRecLen = LEN(UBCustRec(1))
'  UBCustSNLen = LEN(UBCustSN(1))
'
'  SearchLen = LEN(SEARCH$)
'  Match = False
'
''This search reads 1024 recs at a pop into a search buffer
''Basics file i/o CAN NOT read in this way.
''For a Cust file of "5000" recs
''Basic will do "5000" disk accesses.
''Our search ONLY NEEDS 5!
''5000 Recs, BlockSize 1024.
'' 4reads * 1024 = 4096 Recs
''   5000 - 4096 = 904 Odd Recs
'' 1read  * 904 Recs
'' 4 + 1 = 5 Reads!
''This is many times faster than Basic can do.
'
'  FOpenS "UBCUSTSN.DAT", C1Handle               'open data file
'  FOpenS "UBCUST.DAT", R1Handle 'open data file
'
'  NumOfCust& = FLof&(C1Handle) / UBCustSNLen
'
'  '************************************
'  NumChunks& = NumOfCust& \ MaxBlockCnt
'  '****DO NOT CHANGE THE DIVISION HERE!
'  OddRecs = NumOfCust& MOD MaxBlockCnt
'
'  IF NumChunks& = 0 THEN        'if the actual cust count is less than
'    'the work buffer
'    MaxBlockCnt = OddRecs
'    NumChunks& = 1
'    OddRecs = 0
'  END IF
'
'  BlockSize = UBCustSNLen * MaxBlockCnt
'
'  'Find matching record
'
'  FOR CCnt& = 1 TO NumChunks&
'    FGetRTA C1Handle, UBCustSN(1), CCnt&, BlockSize
'    FOR RecCnt = 1 TO MaxBlockCnt
'      UBSearchN$ = LEFT$(UBCustSN(RecCnt).SearchName, SearchLen)
'      IF (SEARCH$ = UBSearchN$) THEN
'        IF LEN(QPTrim$(UBCustSN(RecCnt).DelFlag)) GOTO DelSkip2
'        IF (ActiveOnly = 0) OR ((ActiveOnly = True) AND ((UBCustSN(RecCnt).Status = "A"))) THEN
'          WhatRec& = UBCustSN(RecCnt).RecNum
'          GOSUB CustLoadEM2
'        ELSEIF (ActiveOnly = 1) AND (UBCustSN(RecCnt).Status = "I") THEN
'          WhatRec& = UBCustSN(RecCnt).RecNum
'          GOSUB CustLoadEM2
'        END IF
'      END IF
'DelSkip2:
'    NEXT
'    ShowSearchWheel 12, 44
'  NEXT
'
'
'  IF OddRecs > 0 THEN                  'this is always less than max (1024)
'    BlockSize = UBCustSNLen * OddRecs  'Adj Block size to get last chunk
'    FGetA C1Handle, UBCustSN(1), BlockSize
'    FOR RecCnt = 1 TO OddRecs   'search'em
'      UBSearchN$ = LEFT$(UBCustSN(RecCnt).SearchName, SearchLen)
'      IF (SEARCH$ = UBSearchN$) THEN
'        IF LEN(QPTrim$(UBCustSN(RecCnt).DelFlag)) GOTO DelSkip3
'        IF (ActiveOnly = 0) OR ((ActiveOnly = True) AND ((UBCustSN(RecCnt).Status = "A"))) THEN
'          WhatRec& = UBCustSN(RecCnt).RecNum
'          GOSUB CustLoadEM2
'        ELSEIF (ActiveOnly = 1) AND (UBCustSN(RecCnt).Status = "I") THEN
'          WhatRec& = UBCustSN(RecCnt).RecNum
'          GOSUB CustLoadEM2
'        END IF
'      END IF
'DelSkip3:
'    NEXT
'    ShowSearchWheel 12, 44
'  END IF
'
'  FClose C1Handle               'close files
'  FClose R1Handle
'
'  IF DCnt = 0 THEN
'    RecNo& = 0
'    GOTO ExitSearch2
'  ELSE
'
'    MaxLen = 59 'Set menu width to zero
'    Action = 0  '0 means stay in the menu until they select something
'
'    IF Choice < 1 THEN
'      Choice = 1                'Pre-load choice to highlight
'    END IF
'
'    Title$ = SPACE$(MaxLen + 4)
'    LSET Title$ = "  Customer/Owner       Service Address        Location No.  S"
'    '--Find max menu width
'
'    '--Center Menu within Screen
'
'    Row = 4
'    Col = ((80 - 60) \ 2) - 1
'
'    IF CLSFlag THEN
'      Row = 4
'      BoxBot = 17               'limit the box length
'      BlockClear
'    ELSE
'      Row = 6
'      BoxBot = 14               'limit the box length to go no lower than line 20
'      RestScrn ScrnArray()
'    END IF
'
'    LOCATE Row, Col, 0
'
'    DO
'      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
'      QPrintRC "Matched:" + STR$(DCnt), BoxBot + 4, Col + 2, 15
'      QPrintRC Title$, Row - 1, Col, 112
'      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
'      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
'      IF Ky$ = CHR$(27) THEN
'        RecNo& = -1
'        EXIT DO 'choice = 0
'      END IF
'      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
'    LOOP UNTIL RecNo& > 0
'  END IF
'
'ExitSearch2:
'  RestScrn ScrnArray()
'
'  ERASE ScrnArray, UBCustRec, MChoice, UBCustSN
'
'  EXIT SUB
'
'CustLoadEM2:
'
'  FGetRTA R1Handle, UBCustRec(1), WhatRec&, UBCustRecLen
'
'  DCnt = DCnt + 1
'  'IF FRE(-1) < 5000 THEN STOP
'  REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
'  LSET MChoice(DCnt).V = LEFT$(QPTrim$(UBCustRec(1).CustName), 20)
'  MID$(MChoice(DCnt).V, 22, 22) = LEFT$(QPTrim$(UBCustRec(1).SERVADDR), 25)
'  MID$(MChoice(DCnt).V, 48, 9) = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
'  MID$(MChoice(DCnt).V, 59, 1) = UBCustRec(1).Status
'  MID$(MChoice(DCnt).V, 61) = MKL$(WhatRec&)
'
'RETURN
'
'END SUB

SUB Search4LNum (LocNum$, RecNo&, CLSFlag%, ActiveOnly%)
  
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  DisplayUBScrn "SHOWSCRH"
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  IdxRecLen = 4 'we are using a integer
  IdxFileSize& = FileSize("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  SearchLen = LEN(LocNum$)
  
  Match = False
  FirstRec = 1
  LastRec = IdxNumOfRecs
  
  BotOffSet = 0
  TopOffSet = IdxNumOfRecs
  
  FOpenS "UBCUST.DAT", C1Handle 'open data file
  'Find matching record
  MidRec = (LastRec + FirstRec) \ 2
  
  DO
    IF LastSRec = MidRec THEN EXIT DO
    LastSRec = MidRec
    FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(MidRec).RecNum), UBCustRecLen
    UBSearchN$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
    ShowSearchWheel 12, 44
    
    IF (LocNum$ = UBSearchN$) AND (UBCustRec(1).DelFlag = 0) THEN
      IF MidRec - BotOffSet > 1 THEN
        MidRec = MidRec - 1
      ELSE
        FirstMatchRec = MidRec
      END IF
    ELSEIF LocNum$ < UBSearchN$ THEN            'lower
      TopOffSet = MidRec
      MidRec = TopOffSet - ((TopOffSet - BotOffSet) \ 2)
    ELSE        'higher
      BotOffSet = MidRec
      MidRec = BotOffSet + ((TopOffSet - BotOffSet) \ 2) + 1
      IF MidRec = IdxNumOfRecs + 1 THEN
        EXIT DO
      END IF
    END IF
    IF TopOffSet = BotOffSet THEN EXIT DO
    'Look into this
    'IF TopOffSet = BotOffSet THEN EXIT DO
  LOOP UNTIL FirstMatchRec
  
  FClose C1Handle
  
  IF FirstMatchRec = 0 THEN
    RecNo& = 0
  ELSE
    RecNo& = IdxBuff(FirstMatchRec).RecNum
  END IF
  
  IF ActiveOnly AND UBCustRec(1).Status <> "A" THEN
    RecNo& = 0
  ELSEIF (ActiveOnly = 1) AND (UBCustRec(1).Status <> "I") THEN
    RecNo& = 0
  END IF
ExitLSearch:
  
  'cls
  'Shell
  RestScrn ScrnArray()
  ERASE ScrnArray, UBCustRec, IdxBuff
END SUB

SUB Search4Meter (MeterNum$, RecNo&, CLSFlag%, ActiveOnly%)
  
  STATIC Choice, OMeterNum$
  
  IF OMeterNum$ <> MeterNum$ THEN
    Choice = 1
    OMeterNum$ = MeterNum$
  END IF
  
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  WPos = 1
  
  DisplayUBScrn "SHOWSCRH"
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  CustFileSize& = FileSize("UBCUST.DAT")
  NumCustRecs = CustFileSize& \ UBCustRecLen
  
  'REDIM MChoice(1 TO 1) AS FLen2
  
  REDIM RecBuff(1 TO 1) AS LONG
  
  FOpenS "UBCUST.DAT", C1Handle 'open data file
  'Find matching record
  
  MatchCnt = 0
  FOR Cnt = 1 TO NumCustRecs
    ShowSearchWheel 12, 44
    FGetRTA C1Handle, UBCustRec(1), CLNG(Cnt), UBCustRecLen
    IF NOT UBCustRec(1).DelFlag THEN
      'IF NOT ActiveOnly OR (ActiveOnly AND (UBCustRec(1).Status = "A")) THEN
      IF (ActiveOnly = 0) OR ((ActiveOnly = True) AND ((UBCustRec(1).Status = "A"))) THEN
        GOSUB CheckEM2
      ELSEIF (ActiveOnly = 1) AND (UBCustRec(1).Status = "I") THEN
        GOSUB CheckEM2
      END IF
    END IF
  NEXT
  
  IF MatchCnt = 0 THEN
    RecNo& = 0
    FClose C1Handle
    GOTO ExitMeterSearch
  END IF
  IF NOT DebugFlag THEN
    FreeMem& = FRE(-1)
    IF FreeMem& >= 65536 THEN
      FreeMem& = 65536
    END IF
    MemNeeded& = MatchCnt * 64&
    IF MemNeeded& > FreeMem& THEN
      FClose C1Handle
      QPrintRC "Matched:>" + QPStrI(MatchCnt), 24, 1, 15
      QPrintRC " Needed: " + QPStrL(MemNeeded&) + "  Free: " + QPStrL(FreeMem&), 25, 1, 15
      RecNo& = -1
      WaitForAction
      GOTO ExitMeterSearch
    END IF
  END IF
  REDIM MChoice(1 TO MatchCnt) AS FLen2
  
  FOR Cnt = 1 TO MatchCnt
    ShowSearchWheel 12, 44
    FGetRTA C1Handle, UBCustRec(1), RecBuff(Cnt), UBCustRecLen
    TCustName$ = LEFT$(QPTrim$(UBCustRec(1).CustName), 30)
    Cnt = Cnt - 1
    FOR MeterCnt = 1 TO 7
      IF INSTR(UBCustRec(1).LocMeters(MeterCnt).MTRNUM, MeterNum$) > 0 THEN
        Cnt = Cnt + 1
        LSET MChoice(Cnt).V = TCustName$
        Book$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
        MID$(MChoice(Cnt).V, 32, 12) = UBCustRec(1).LocMeters(MeterCnt).MTRNUM
        MID$(MChoice(Cnt).V, 50, 9) = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
        MID$(MChoice(Cnt).V, 61) = MKL$(RecBuff(Cnt))
      END IF
    NEXT
  NEXT
  
  FClose C1Handle
  
  IF MatchCnt = 0 THEN
    RecNo& = 0
  ELSEIF MatchCnt > 1 THEN
    QPrintRC "Sorting. . .  ", 12, 32, -1
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    IF Choice = 0 THEN
      Choice = 1                'Pre-load choice to highlight
    ELSEIF Choice > MatchCnt THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Customer/Owner               Meter No.         Location No."
    '--Find max menu width
    
    '--Center Menu within Screen
    
    Row = 4
    Col = ((80 - 60) \ 2) - 1
    
    IF CLSFlag THEN
      Row = 4
      BoxBot = 17               'limit the box length
      BlockClear
    ELSE
      Row = 6
      BoxBot = 14               'limit the box length to go no lower than line 20
      RestScrn ScrnArray()
    END IF
    
    LOCATE Row, Col, 0
    
    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Matched:" + STR$(MatchCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
    LOOP UNTIL RecNo& > 0
  ELSE
    RecNo& = CVL(MID$(MChoice(1).V, 61, 4))
  END IF
  
ExitMeterSearch:
  
  RestScrn ScrnArray()
  ERASE ScrnArray, UBCustRec, RecBuff, MChoice
  EXIT SUB
  
CheckEM2:
  FOR MeterCnt = 1 TO 7
    IF INSTR(UBCustRec(1).LocMeters(MeterCnt).MTRNUM, MeterNum$) > 0 THEN
      MatchCnt = MatchCnt + 1
      REDIM PRESERVE RecBuff(1 TO MatchCnt) AS LONG
      RecBuff(MatchCnt) = Cnt
    END IF
  NEXT
RETURN
  
END SUB

SUB Search4SAddr (SAddr$, RecNo&, CLSFlag%, ActiveOnly%)
  
  STATIC Choice
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  WPos = 1
  
  DisplayUBScrn "SHOWSCRH"
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  CustFileSize& = FileSize("UBCUST.DAT")
  NumCustRecs = CustFileSize& \ UBCustRecLen
  
  'REDIM MChoice(1 TO 1) AS FLen2
  
  REDIM RecBuff(1 TO 1) AS LONG
  
  FOpenS "UBCUST.DAT", C1Handle 'open data file
  
  MatchCnt = 0
  FOR Cnt = 1 TO NumCustRecs
    ShowSearchWheel 12, 44
    FGetRTA C1Handle, UBCustRec(1), CLNG(Cnt), UBCustRecLen
    IF NOT UBCustRec(1).DelFlag THEN
      IF (ActiveOnly = 0) OR ((ActiveOnly = True) AND ((UBCustRec(1).Status = "A"))) THEN
        GOSUB CheckLoadEM2
      ELSEIF (ActiveOnly = 1) AND (UBCustRec(1).Status = "I") THEN
        GOSUB CheckLoadEM2
      END IF
    END IF
  NEXT
  
  FClose C1Handle
  
  IF NOT DebugFlag THEN
    FreeMem& = FRE(-1)
    IF FreeMem& >= 65536 THEN
      FreeMem& = 65536
    END IF
    MemNeeded& = MatchCnt * 64&
    IF MemNeeded& > FreeMem& THEN
      QPrintRC "Matched: " + QPStrI(MatchCnt), 24, 1, 15
      QPrintRC " Needed: " + QPStrL(MemNeeded&) + "  Over: " + QPStrL(MemNeeded& - FreeMem&), 25, 1, 15
      RecNo& = -1
      WaitForAction
      GOTO ExitSAddrSearch
    END IF
  END IF
  
  IF MatchCnt = 0 THEN
    GOTO ExitSAddrSearch
    RecNo& = -1
  ELSEIF MatchCnt > 1 THEN
    REDIM MChoice(1 TO MatchCnt) AS FLen2
    FOpenS "UBCUST.DAT", C1Handle               'open data file
    FOR Cnt = 1 TO MatchCnt
      FGetRTA C1Handle, UBCustRec(1), CLNG(RecBuff(Cnt)), UBCustRecLen
      Addr$ = LEFT$(QPTrim$(UBCustRec(1).SERVADDR), 25)
      Book$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
      LSET MChoice(Cnt).V = LEFT$(QPTrim$(UBCustRec(1).CustName), 20)
      MID$(MChoice(Cnt).V, 22, 25) = Addr$
      MID$(MChoice(Cnt).V, 50, 9) = Book$
      MID$(MChoice(Cnt).V, 61) = MKL$(RecBuff(Cnt))
      ShowSearchWheel 12, 44
    NEXT
    FClose C1Handle
    
    IF DCnt = 0 THEN
      RecNo& = 0
      GOTO ExitSAddrSearch
    END IF
    
    'FClose L1Handle
    
    QPrintRC "Sorting. . .   ", 12, 32, -1
    SortT MChoice(1), MatchCnt, 0, 64, 21, 25
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    IF Choice = 0 THEN
      Choice = 1                'Pre-load choice to highlight
    ELSEIF Choice > MatchCnt THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Customer                      Address"
    '--Find max menu width
    
    '--Center Menu within Screen
    
    Row = 4
    Col = ((80 - 60) \ 2) - 1
    
    IF CLSFlag THEN
      Row = 4
      BoxBot = 17               'limit the box length
      BlockClear
    ELSE
      Row = 6
      BoxBot = 14               'limit the box length to go no lower than line 20
      RestScrn ScrnArray()
    END IF
    
    LOCATE Row, Col, 0
    
    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Matched:" + STR$(MatchCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
    LOOP UNTIL RecNo& > 0
  ELSE
    RecNo& = RecBuff(MatchCnt)
  END IF
  
ExitSAddrSearch:
  
  'cls
  'Shell
  RestScrn ScrnArray()
  ERASE ScrnArray, UBCustRec, MChoice
  EXIT SUB
  
CheckLoadEM2:
  Cnted = Cnted + 1
  IF INSTR(UBCustRec(1).SERVADDR, SAddr$) > 0 THEN
    DCnt = DCnt + 1
    MatchCnt = MatchCnt + 1
    REDIM PRESERVE RecBuff(1 TO MatchCnt) AS LONG
    RecBuff(MatchCnt) = Cnt
  END IF
RETURN
  
END SUB

SUB ShowCustConsHist (CustRec&)
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  REDIM Metered(1 TO 15)
  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
  IF INSTR(UBSetUpRec(1).UtilName, "TROY") > 0 THEN
    TroyFlag = True
  END IF
  NumOfRevs = MaxRevsCnt
  FOR RevCnt = 1 TO 15
    RLen = LEN(QPTrim$(LEFT$(UBSetUpRec(1).Revenues(RevCnt).RevName, 14)))
    IF RLen >= 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    END IF
    IF UBSetUpRec(1).Revenues(RevCnt).UseMtr = "Y" THEN
      Metered(RevCnt) = True
    END IF
  NEXT
  
  REDIM MChoice(1 TO 1) AS FLen2
  REDIM UBTranRec(1) AS UBTransRecType
  REDIM UBCustRec(1) AS NewUBCustRecType
  
  UBCustRecLen = LEN(UBCustRec(1))
  UBTranRecLen = LEN(UBTranRec(1))
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  GET UBFile, CustRec&, UBCustRec(1)
  CLOSE UBFile
  
  CurBal# = UBCustRec(1).CurrBalance
  PreBal# = UBCustRec(1).PrevBalance
  
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  
  PrevTranRec& = UBCustRec(1).LastTrans
  
  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
      GET UBTran, PrevTranRec&, UBTranRec(1)
      IF UBTranRec(1).TransType = TranUtilityBill OR UBTranRec(1).TransType = TranUtilityBill + 100 THEN
        FOR MtrCnt = 1 TO 7
          IF UBTranRec(1).MtrTypes(MtrCnt) > 0 THEN
            DCnt = DCnt + 1
            REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
            'LSET MChoice(DCnt).V = Num2Date(UBTranRec(1).ReadDate)
            LSET MChoice(DCnt).V = Num2Date(UBTranRec(1).TransDate)
            SELECT CASE UBTranRec(1).MtrTypes(MtrCnt)
            CASE MtrWaterOnly
              MeterType$ = "Water"
            CASE MtrSewerOnly
              MeterType$ = "Sewer"
            CASE MtrCombined
              MeterType$ = "Combined"
            CASE MtrElectric
              MeterType$ = "Electric"
            CASE MtrDemand
              MeterType$ = "D Electric"
            CASE MtrGas
              MeterType$ = "Gas Meter"
            CASE MtrTouchRead
              MeterType$ = "Touch Read"
            CASE MtrLightsService
              MeterType$ = "L Service"
            END SELECT
            MID$(MChoice(DCnt).V, 13) = MeterType$
            MID$(MChoice(DCnt).V, 26) = FUsing$(STR$(UBTranRec(1).CurRead(MtrCnt)), "##########")
            MID$(MChoice(DCnt).V, 38) = FUsing$(STR$(UBTranRec(1).PrevRead(MtrCnt)), "##########")
            MeterConsp& = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
            IF MeterConsp& < 0 THEN
              MaxMeterAmt& = 10& ^ (LEN(STR$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
              MeterConsp& = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
            END IF
'working here
            MtrMulti# = 0
            FOR MCnt = 1 TO 7
              IF UBTranRec(1).MtrTypes(MtrCnt) = GetCustMeterType%(UBCustRec(), MCnt) THEN
                MtrMulti# = UBCustRec(1).LocMeters(MCnt).MtrMulti
                IF UBCustRec(1).LocMeters(MCnt).MTRUnit = "C" THEN
                  MeterConsp& = MeterConsp& * 7.481
                  EXIT FOR
                END IF
              END IF
            NEXT
            IF MtrMulti# = 0 THEN
              IF TroyFlag THEN
                MtrMulti# = 100
              ELSE
                MtrMulti# = 1
              END IF
            END IF
            MID$(MChoice(DCnt).V, 52) = FUsing$(STR$(MtrMulti# * MeterConsp&), "##########")
          END IF
        NEXT
      END IF
      PrevTranRec& = UBTranRec(1).PrevTrans
    LOOP
    
    CLOSE UBTran
    RestScrn TempScrn()
    MPaintBox 3, 5, 22, 75, 8
    
    MaxLen = 62 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    Balance$ = Title$
    LSET Title$ = " Trans Date   Meter Type      Current    Previous    Consumption"
    
    '--Find max menu width
    '--Center Menu within Screen
    
    Row = 4
    Col = 8
    Row = 6
    BoxBot = 17 'limit the box length to go no lower than line 20
    
    TitleBox BoxBot + 3, Col, MaxLen + 3, "Press <ESC> to continue.", Cnf
    
    QPrintRC Title$, Row - 1, Col, 112
    MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
    
    DO
      LOCATE Row, Col, 0
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RestScrn TempScrn()
        EXIT DO
      END IF
    LOOP
  ELSE
    CLOSE UBTran
    Ok = MsgBox%("UB.QSL", "NOCTRANS")
    RestScrn TempScrn()
  END IF
  
  RestScrn TempScrn()
  ERASE Metered, UBSetUpRec, MChoice
  ERASE TempScrn, UBTranRec, UBCustRec
  
  EXIT SUB
  
  
END SUB

SUB ShowCustHistory (CustRec&)
  
  u$ = CHR$(24)
  d$ = CHR$(25)
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  DisplayUBScrn "UBCUHIST"
  
  REDIM RevText$(1 TO MaxRevsCnt)
  REDIM Metered(1 TO 15)
  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
  NumOfRevs = MaxRevsCnt
  FOR RevCnt = 1 TO 15
    RevText$(RevCnt) = LEFT$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName), 14)
    IF LEN(RevText$(RevCnt)) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    END IF
    IF UBSetUpRec(1).Revenues(RevCnt).UseMtr = "Y" THEN
      Metered(RevCnt) = True
    END IF
  NEXT
  
  REDIM MChoice(1 TO 1) AS FLen2
  
  REDIM UBTranRec(1) AS UBTransRecType
  REDIM UBCustRec(1) AS NewUBCustRecType
  
  UBCustRecLen = LEN(UBCustRec(1))
  UBTranRecLen = LEN(UBTranRec(1))
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  GET UBFile, CustRec&, UBCustRec(1)
  CLOSE UBFile
  
  CurBal# = UBCustRec(1).CurrBalance
  PreBal# = UBCustRec(1).PrevBalance
  
Top:
  
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  
  PrevTranRec& = UBCustRec(1).LastTrans
  
  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
      DCnt = DCnt + 1
      REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
      GET UBTran, PrevTranRec&, UBTranRec(1)
      LSET MChoice(DCnt).V = Num2Date(UBTranRec(1).TransDate)
      'MID$(MChoice(DCnt).V, 15) = UBTranRec(1).TransDesc
      GOSUB GetTransType
      MID$(MChoice(DCnt).V, 13) = TType$
      MID$(MChoice(DCnt).V, 41) = FUsing(STR$(UBTranRec(1).TransAmt), "#####.##")
      'this will show th actual trans number in the list
      'MID$(MChoice(DCnt).V, 50) = FUsing(STR$(PrevTranRec&), "######")
      MID$(MChoice(DCnt).V, 52) = FUsing(STR$(UBTranRec(1).RunBalance), "#####.##")
      MID$(MChoice(DCnt).V, 61) = MKL$(PrevTranRec&)
      PrevTranRec& = UBTranRec(1).PrevTrans
    LOOP
    
    CLOSE UBTran
    
    RestScrn TempScrn()
    MPaintBox 3, 5, 22, 75, 8
    REDIM TempScrn2(0)
    SaveScrn TempScrn2()
    
HistTop:
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    Balance$ = Title$
    LSET Title$ = "  Trans Date       Description           Trans Amt    Balance  "
    LSET Balance$ = " Balance:" + FUsing(STR$(CurBal# + PreBal#), ",#####.##") + "   Cur:" + FUsing(STR$(CurBal#), ",#####.##") + "  Prev:" + FUsing(STR$(PreBal#), ",#####.##")
    
    '--Find max menu width
    '--Center Menu within Screen
    
    Row = 4
    Col = ((80 - 60) \ 2) - 1
    
    Row = 6
    BoxBot = 17 'limit the box length to go no lower than line 20
    
    'TitleBox BoxBot + 3, Col, MaxLen + 3, "       Press <ESC> to continue.", Cnf
    
    WazzWind BoxBot + 2, Col, BoxBot + 5, MaxLen + 3 + Col, 10, 4, True
    QPrintRC "  Use:  " + u$ + "-" + d$ + " to select.", BoxBot + 3, Col + 3, 15
    QPrintRC u$, BoxBot + 3, Col + 11, 14
    QPrintRC d$, BoxBot + 3, Col + 13, 14
    
    QPrintRC "Total: " + STR$(DCnt), BoxBot + 4, Col + 3, 15
    QPrintRC "Press:   [ESC] to continue.", BoxBot + 3, Col + 33, 15
    QPrintRC "        [ENTER] for detail.", BoxBot + 4, Col + 33, 15
    QPrintRC "ESC", BoxBot + 3, Col + 43, 14
    QPrintRC "ENTER", BoxBot + 4, Col + 42, 14
    
    QPrintRC Balance$, Row - 2, Col, 112
    QPrintRC Title$, Row - 1, Col, 112
    MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
    'FirstTime = True
    
    'SLEEP
    
    DO
      LOCATE Row, Col, 0
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RestScrn TempScrn()
        EXIT DO 'choice = 0
      ELSEIF Ky$ = CHR$(13) THEN
        RestScrn TempScrn2()
        GOTO ShowTransDetail
      END IF
    LOOP        'UNTIL EditLocRec& > 0
  ELSE
    CLOSE UBTran
    Ok = MsgBox%("UB.QSL", "NOCTRANS")
    RestScrn TempScrn()
  END IF
  
  RestScrn TempScrn()
  ERASE RevText$, Metered, UBSetUpRec, MChoice
  ERASE TempScrn, UBTranRec, UBCustRec
  
  EXIT SUB
  
ShowTransDetail:
  CursorOff
  TransRecNum& = CVL(RIGHT$(MChoice(Choice).V, 4))
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  GET UBTran, TransRecNum&, UBTranRec(1)
  CLOSE UBTran
  
  DisplayUBScrn "TRDETAIL"
  
  QPrintRC Num2Date(UBTranRec(1).TransDate), 3, 23, 15
  
  GOSUB GetTransType
  
  QPrintRC FUsing$(STR$(UBTranRec(1).TransAmt), "#####.##"), 4, 25, 15
  
  QPrintRC TType$, 4, 50, 15
  QPrintRC UBTranRec(1).TransDesc, 3, 50, 15
  
  FOR RevCnt = 1 TO NumOfRevs
    QPrintRC RevText$(RevCnt), RevCnt + 6, 8, 15
    QPrintRC FUsing$(STR$(UBTranRec(1).RevAmt(RevCnt)), "#####.##"), RevCnt + 6, 25, 15
    QPrintRC FUsing$(STR$(UBTranRec(1).TaxAmt(RevCnt)), "###.##"), RevCnt + 6, 36, 15
    '(Number$, Image$)
  NEXT
  
  FOR Cnt = 1 TO 7
    IF Metered(Cnt) THEN
      QPrintRC FUsing$(STR$(UBTranRec(1).CurRead(Cnt)), "#########"), Cnt + 6, 42, 15
      QPrintRC FUsing$(STR$(UBTranRec(1).PrevRead(Cnt)), "#########"), Cnt + 6, 53, 15
      IF QPTrim$(UBTranRec(1).ESTREAD(Cnt)) = "" THEN
        QPrintRC "N", Cnt + 6, 70, 15
      ELSE
        QPrintRC "Y", Cnt + 6, 70, 15
      END IF
    END IF
  NEXT
  
  WaitForAction
  RestScrn TempScrn2()
  GOTO HistTop
  
GetTransType:
  
  SELECT CASE UBTranRec(1).TransType
  CASE TranUtilityBill, TranUtilityBill + 100
    TType$ = "Utility Bill "
  CASE TranLateCharge, TranReconnectFee, TranLateCharge + 100, TranReconnectFee + 100
    TType$ = "Penalty, Reconnect Fee"
  CASE TranBillPayment, TranBillPayment + 100
    TDesc$ = QPTrim$(UBTranRec(1).TransDesc)
    IF INSTR(UBTranRec(1).TransDesc, "PAYMENT") = 0 AND LEN(TDesc$) > 0 THEN
      TType$ = "Utility Payment " + LEFT$(QPTrim$(UBTranRec(1).TransDesc), 10)
    ELSE
      TType$ = "Utility Payment"
    END IF
  CASE TranPenaltyPayment
    TType$ = "Penalty Payment"
  CASE TranPenaltyCharge
    TType$ = "Penalty/Late Fee"
  CASE TranAppliedDeposit
    TType$ = "Applied Deposit"
  CASE TranDepositPayment, TranDepositPayment + 100
    TDesc$ = QPTrim$(UBTranRec(1).TransDesc)
    IF INSTR(UBTranRec(1).TransDesc, "DEPOSIT") = 0 AND LEN(TDesc$) > 0 THEN
      TType$ = "Deposit Payment " + LEFT$(QPTrim$(UBTranRec(1).TransDesc), 10)
    ELSE
      TType$ = "Deposit Payment"
    END IF
  CASE TranDraftPayment
    TType$ = "Draft Payment"
  CASE TranBeginBalance, TranBeginBalance + 100
    TType$ = "Beginning Balance"
  CASE 9
    TType$ = "Deposit Refund"
  CASE TranUpwardAdjustment
    TType$ = "Upward Adjustment"
  CASE TranDownwardAdjustment
    TType$ = "Downward Adjustment"
  CASE ELSE
    TType$ = STR$(UBTranRec(1).TransType) + " ???"
  END SELECT
  
RETURN
  
END SUB

SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%) STATIC
  RSET PctC(1) = QPStrI$(INT((RecNo / NumOfRecs) * 100))
  'HideCursor
  QPrintRC PctC(1), 13, 40, Cnf.HiLite
  'ShowCursor
  '  QPrintRC STR$(FRE("")), 25, 1, Cnf.HiLite
END SUB

SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&) STATIC
  RSET PctC(1) = QPStrL$(INT((RecNo& / NumOfRecs&) * 100))
  'HideCursor
  QPrintRC PctC(1), 13, 40, Cnf.HiLite
  'ShowCursor
  '  QPrintRC STR$(FRE("")), 25, 1, Cnf.HiLite
END SUB

SUB ShowProcessingScrn (RptTitle$)
  TitleRow = 9
  TitleCol = 40 - (LEN(RptTitle$) \ 2) + 1
  CursorOff
  BlockClear
  DisplayUBScrn "PRORPT"
  HideCursor
  QPrintRC RptTitle$, TitleRow, TitleCol, 126
  QPrintRC "Processing:    % Completed.", 13, 28, Cnf.HiLite
  ShowCursor
END SUB

SUB ShowSearchWheel (BYVAL Row%, BYVAL Col%) STATIC
  WheelPos = WheelPos + 1
  IF WheelPos > 4 THEN WheelPos = 1
  HideCursor
  QPrintRC MID$(Wheel$, WheelPos, 1), Row, Col, -1
  ShowCursor
END SUB

