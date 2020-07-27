DEFINT A-Z
DECLARE FUNCTION ASCII% (Strng$)
DECLARE FUNCTION CheckFldEmpty% (FldData$)
DECLARE FUNCTION Compare3% (BYVAL Seg1, BYVAL Adr1, BYVAL Seg2, BYVAL Adr2, NumBytes)
DECLARE FUNCTION DOSError% ()
DECLARE FUNCTION EMPCheckEMPNumber% (EMPNumFld$, Frm AS ANY)
DECLARE FUNCTION EMPNumLeaveRec% ()
DECLARE FUNCTION EMPNumNull% (Form$())
DECLARE FUNCTION EmpAsk2Delete% ()
DECLARE FUNCTION EmpNumOfRecs% ()
DECLARE FUNCTION EmpNumOfRecs% ()
DECLARE FUNCTION EndOfForms (Fld() AS ANY)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FEof% (Handle%)
DECLARE FUNCTION FLInput$ (Handle%, Buff$)
DECLARE FUNCTION FLoc& (Handle%)
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FldNum (FldName$, Fld() AS ANY)
DECLARE FUNCTION LibNumberOfFields (LibName$, FrmName$)
DECLARE FUNCTION LibScrName$ (Scr, ScrnLib())
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION PDQTimer& ()
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPRTrim$ (Text$)
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPStrL$ (LongNum&)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI (NumString$)
DECLARE FUNCTION Round# (DoubleNum#)
DECLARE FUNCTION ScrnLibSize (LibName$)
DECLARE FUNCTION Value# (e$, ErCode%)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION WhichError% ()

DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BlockClear ()
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, X%, y%)
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DeleteT (SEG StartElement AS ANY, ElSize%, NumEls%)
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB EMPAddShowName (First$, Last$)
DECLARE SUB EMPChooseList (MatchCnt%, LookUpType%)
DECLARE SUB EMPEditRec (RecNo%)
DECLARE SUB EMPLoadPrintMask (FileMaskName$)
DECLARE SUB EMPLoadRFldData (ReqFld%(), ReqNumOfFlds%)
DECLARE SUB EMPShowTransHist (RecNo%)
DECLARE SUB EMPShowYTDTotals (RecNo%)
DECLARE SUB EMPrintEmpData ()
DECLARE SUB EMPrintEmpList ()
DECLARE SUB EMPrintTermEmpList ()
DECLARE SUB EditForm (Form$(), Fld() AS ANY, Frm AS ANY, Cnf AS ANY, Action)
DECLARE SUB EmpLookUpEdit ()
DECLARE SUB EmpSaveNewRec (RecNo%, Form$())
DECLARE SUB EmpSaveOldRec (RecNo%, Form$())
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FFlush (Handle%)
DECLARE SUB FGetA (Handle%, SEG Element AS ANY, NumOfBytes&)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRT (Handle%, Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetT (Handle%, Dest AS ANY, NumBytes%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPut (Handle%, ToDisk$)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutRT (Handle%, Source AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Source AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutT (Handle%, Source AS ANY, NumBytes%)
DECLARE SUB FSeek (Handle%, Location&)
DECLARE SUB GetCursor (X, y, Button)
DECLARE SUB HideCursor ()
DECLARE SUB InitEmpMulti (Fld() AS ANY)
DECLARE SUB InsertT (SEG StartElement AS ANY, ElSize%, NumEls%)
DECLARE SUB Lib2Scrn (NameInLib$, ScrnLib(), MonoCode, Attribute, ErrorCode)
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB LibGetFldDef (LibName$, FrmName$, StartEl, Fld() AS ANY, Form$(), ErrCode)
DECLARE SUB LoadScrnLib (LibName$, ScrnLib(), ErrorCode)
DECLARE SUB MPaintBox (ULRow, ULCol, LRRow, LRCol, Colr)
DECLARE SUB MQPrint (X$, Colr)
DECLARE SUB MScrnRest (ULRow, ULCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnSave (ULRow, ULCol, LRRow, LRCol, SEG Address)
DECLARE SUB Pause (Ticks)
DECLARE SUB PayMenu (BYVAL MenuNum%, Choice%, NumOfItems%)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB PrintArray (FirstFld, LastFld, Form$(), Fld() AS ANY)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB RestScrn (ScrnArray%())
DECLARE SUB SaveField (FldNo%, Form$(), Fld() AS ANY, BadFld%)
DECLARE SUB SaveScrn (ScrnArray%())
DECLARE SUB SearchDataFile (LookUpType%, Form$(), MatchCnt%)
DECLARE SUB SearchT2 (SEG Array AS ANY, NumEls%, Match, Found%, Dir%, Code%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB Sort (Array$(), Direction%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB TextCursor (Fg, Bg)
DECLARE SUB UnPackBuffer (FirstFld, LastFld, Form$(), Fld() AS ANY)
DECLARE SUB VertMenu (Item$(), Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB WazzWind (BYVAL TopRow, BYVAL LeftCol, BYVAL BotRow, BYVAL RghtCol, BYVAL FrameColor, BYVAL FrameType, BYVAL Shadow)

'$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.

'$INCLUDE: 'FieldInf.bi'                '        "        field information
'$INCLUDE: 'FormEdit.bi'                '        "        form editing info.
'$INCLUDE: 'PRFiles.bi'
'$INCLUDE: 'PREmpRec.biw'
'$INCLUDE: 'PageInfo.bi'    '        "          Page information
'$INCLUDE: 'Retire.Bi'
'$INCLUDE: 'DedCodes.Bi'
'$INCLUDE: 'ERNCodes.Bi'
'$INCLUDE: 'PRTRANS.Biw'
'$INCLUDE: 'PRCHECK.Bi'
'$INCLUDE: 'PRUNIT.Bi'
'$INCLUDE: 'LEAVE.Bi'

TYPE FLen
   V AS STRING * 43
END TYPE


CONST False = 0
CONST True = NOT False

CONST NameLookUp = True
CONST NumbLookUp = False

CONST ELUBtnRow = 25

  
  DIM SHARED TempData1Rec  AS EmpData1Type
  DIM SHARED EmpData1Rec   AS EmpData1Type

  DIM SHARED EmpData2Rec   AS EmpData2Type
  DIM SHARED EmpData3Rec   AS EmpData3Type

  DIM SHARED TransRec      AS TransRecType

  REDIM SHARED Items(0) AS FLen

SUB EMPAddShowName (First$, Last$)
  HideCursor
  EmpName$ = SPACE$(40)
  LSET EmpName$ = (QPTrim$(First$) + " " + QPTrim$(Last$))
  QPrintRC EmpName$, 1, 26, 112
  EmpName$ = ""
  ShowCursor
END SUB

FUNCTION EmpAsk2Delete

  StartEl = 0                         'Load first form at array start
  FormName$ = "WARNDEL"
  NumFlds = LibNumberOfFields(MiscQLib, FormName$)

  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)             'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo     'DIM the field information array
  LibGetFldDef MiscQLib, FormName$, StartEl, Fld(), Form$(), ErrCode

  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1

  '----- Setup TYPE for setting and reading form editing information.

  Frm(1).FldNo = 1                 'Start editing on field #1
  Frm(1).InsStat = False           'Set insert state (True = Insert on)
  Frm(1).StartEl = 0               'Set form starting element to 0 and

  DisplayMiscScrn FormName$

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
      CASE 89, 121           'yes
        OKFlag = True
      CASE EscKey, 78, 110   'no
        ExitFlag = True
    END SELECT

  LOOP UNTIL OKFlag OR ExitFlag            'proper key not set

  CursorOff
  IF OKFlag THEN
    EmpAsk2Delete = True
  ELSE
    EmpAsk2Delete = False
  END IF
  ERASE Frm, Form$, Fld


END FUNCTION

FUNCTION EMPCheckEMPNumber (EmpNum$, Frm AS FormInfo)

  REDIM TempScrn(0)
  Temp$ = EmpNum$
  RSET EmpNum$ = QPTrim$(Temp$)

  DupeFlag = False
  NumOfRecs = EmpNumOfRecs

'  NumOfRecs = 10

  IF NumOfRecs > 0 THEN
    REDIM EmpNumbers(1 TO NumOfRecs)  AS EmpNumType

    FGetAH EMPNumFileName, EmpNumbers(1), 10, NumOfRecs

    FOR Cnt = 1 TO NumOfRecs
      IF EmpNumbers(Cnt).EmpNum = EmpNum$ THEN
        DupeFlag = True
        EXIT FOR
      END IF
    NEXT
  END IF

  IF DupeFlag THEN
    SaveScrn TempScrn()
    DisplayMiscScrn EMPDupeNum
    WaitForAction
    RestScrn TempScrn()
    Frm.FldNo = 1
    Frm.TxtPos = 1
    Frm.Presses = False
    Frm.KeyCode = 0
    Frm.StartEl = 0

  END IF
  EmpNum$ = Temp$
  EMPCheckEMPNumber = DupeFlag

END FUNCTION

SUB EMPChooseList (MatchCnt%, LookUpType%)

  VAction = 1
  BoxBot = 17
  BoxTop = 6
  MaxLen = 40
  LeftCol = ((80 - MaxLen) \ 2) - 1


  DO
    HideCursor
    LOCATE BoxTop, LeftCol
    BlockClear
    SELECT CASE LookUpType
      CASE NameLookUp
        QPrintRC "  Last Name        First Name     Number    ", BoxTop - 1, LeftCol, Cnf.ActivCH
      CASE NumbLookUp
        QPrintRC "   Number    Last Name         First Name   ", BoxTop - 1, LeftCol, Cnf.ActivCH
    END SELECT

    MPaintBox BoxTop, LeftCol + MaxLen + 4, BoxTop, LeftCol + MaxLen + 5, 8
    WazzWind 21, LeftCol, 23, LeftCol + MaxLen + 3, Cnf.MenBox, 2, True
    QPrintRC " Use (" + CHR$(24) + "-" + CHR$(25) + ") to select", 22, LeftCol + (MaxLen \ 2) - 8, Cnf.ActivCH

    ShowCursor

    DO

    VertMenuT Items(), Choice, MaxLen, BoxBot, Ky$, VAction, Cnf

    SELECT CASE ASCII(Ky$)
      CASE EscKey
        Choice = EscKey
        ExitFlag = True
        EXIT DO
      CASE 13

        EdRecNo = CVI(RIGHT$(Items(Choice).V, 2))
        EMPEditRec EdRecNo

        IF EdRecNo = -100 THEN
          ExitFlag = True
        ELSEIF EdRecNo > 0 THEN 'update list only if rec was edited
          SELECT CASE LookUpType
            CASE NameLookUp
              ENoOffSet = 30
            CASE NumbLookUp
              ENoOffSet = 1
          END SELECT
          IF LEN(QPTrim$(EmpData1Rec.EmpNo)) THEN
            MID$(Items(Choice).V, ENoOffSet) = RIGHT$(EmpData1Rec.EmpNo, 8)
          END IF
        END IF
        EXIT DO
      CASE ELSE
    END SELECT
    LOOP
  LOOP UNTIL ExitFlag
  ERASE Items
  CursorOff

END SUB

SUB EMPEditRec (RecNo%)

  SHARED Choice$()

  REDIM Frm(1)    AS FormInfo

  REDIM ReqFld(0)

  CursorOff

  REDIM TempScrn(0)

  NumOfScrns = 3

  REDIM FormName$(1 TO NumOfScrns)

  FOR Scr = 1 TO NumOfScrns
    FormName$(Scr) = "EMPMA" + LTRIM$(STR$(Scr))
  NEXT

  NumFlds = -1

  FOR Scr = 1 TO NumOfScrns           '-- Get total number of fields
    NumFlds = NumFlds + LibNumberOfFields(EmployeeQLib, FormName$(Scr)) + 1
  NEXT

  REDIM Form$(NumFlds, 2)           'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo   'DIM the field information array
  REDIM PageInfo(1 TO NumOfScrns) AS PageInfoType

  StartEl = 0                         'Load first form at array start
  FOR Scr = 1 TO NumOfScrns           'Get field defs
    LibGetFldDef EmployeeQLib, FormName$(Scr), StartEl, Fld(), Form$(), ErrCode
    PageInfo(Scr).PageNo = Scr
    PageInfo(Scr).FirstFld = StartEl + 1
    StartEl = StartEl + Fld(StartEl).Fields + 1 'Inc StartEl to next
    PageInfo(Scr).LastFld = StartEl - 1
  NEXT

  FOR n = 1 TO NumFlds             'For all fields on both pages
    LSET Form$(n, 0) = ""        'Clear each field
  NEXT

  InitEmpMulti Fld()

  EMPLoadRFldData ReqFld(), ReqNumOfFlds

  LeaveTblFld = FldNum("LEAVETBL", Fld())

  EMPFirstFld = FldNum("EMPFNAME", Fld())
  EMPLastFld = FldNum("EMPLNAME", Fld())

  EmpFedO2F = FldNum("EMPFEDO2", Fld())
  EmpStaO2F = FldNum("EMPSTAO2", Fld())

  EmpRetNoF = FldNum("EMPRETNO", Fld())
  EmpRetTPF = FldNum("EMPRETTP", Fld())

  EmpHPF = FldNum("EMPHP", Fld())
  EmpPTypeF = FldNum("EMPPTYPE", Fld())

  'RateFld = FldNum("EMPPRATE", Fld())
  'ORateFld = FldNum("EMPORATE", Fld())
  'FreqFld = FldNum("EMPPFREQ", Fld())

  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1                 'Start editing on field #1
  Frm(1).InsStat = False           'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0               'Set form starting element to 0 and

  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1
                      
  Scr = 1        'Set screen number to one
  ClearScrn
  
  LibFile2Scrn EmployeeQLib, FormName$(Scr), MonoCode, -1, ErrCode   'display screen

  IF RecNo THEN
'    Fld(1).Protected = True
    Data2RecLen = LEN(EmpData2Rec)
    FOpenS EmpData2Name, D2Handle
    FGetRTA D2Handle, EmpData2Rec, CLNG(RecNo), Data2RecLen
    FClose D2Handle
    OrgNum$ = QPTrim$(EmpData2Rec.EmpNo)
    Form$(0, 0) = SPACE$(Data2RecLen)
    BCopy VARSEG(EmpData2Rec), VARPTR(EmpData2Rec), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), Data2RecLen, 0
    UnPackBuffer 0, 0, Form$(), Fld()
    LSET Form$(1, 0) = LTRIM$(Form$(1, 0))
    EMPAddShowName Form$(EMPFirstFld, 0), Form$(EMPLastFld, 0)
    IF QPTrim$(Form$(EmpRetNoF, 0)) = "" THEN
      Form$(EmpRetTPF, 0) = ""
      Fld(EmpRetTPF).Protected = True
    END IF
  ELSE
    TLastFld = NumFlds - 1
    TFirstFld = TLastFld - 16
    FOR Cnt2 = TLastFld TO TFirstFld STEP -1
      Form$(Cnt2, 0) = "0.00"
    NEXT
    'SaveField 0, Form$(), Fld(), BadFld
  END IF

'---- set benefit percentage field to default to 100%
  IF RecNo = 0 THEN
    'FOR Cnt = 1 TO NumFlds
    '  SELECT CASE Fld(Cnt).FType
    '    CASE 3, 4, 5, 10, 18   'if a numaric type field
    '      Form$(Cnt, 0) = "0"
    '  END SELECT
    'NEXT
    Form$(FldNum("EMPBCODE", Fld()), 0) = "100.00"
  END IF

  IF QPValI(Form$(LeaveTblFld, 0)) < 1 THEN
     Form$(LeaveTblFld, 0) = " 1"
    'FirstTimeFlag
  END IF

  Fld(LeaveTblFld).HiRange = EMPNumLeaveRec%
'***************************************************

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    'IF Frm(1).KeyCode <> 0 THEN STOP

'---- check for mandatory fields
    IF (Frm(1).KeyCode <> 0 AND Frm(1).PrevFld <> Frm(1).FldNo) OR (Frm(1).KeyCode = F0Key) THEN
      FldEmpty = CheckFldEmpty((Form$(Frm(1).PrevFld, 0)))
      IF FldEmpty THEN
        FOR Cnt = 1 TO ReqNumOfFlds
          IF ReqFld(Cnt) = Frm(1).PrevFld THEN
            Frm(1).FldNo = Frm(1).PrevFld
            Frm(1).KeyCode = 0
            Frm(1).StartEl = 0
            SaveScrn TempScrn()
            DisplayMiscScrn NoBlank
            WaitForAction
            RestScrn TempScrn()
          END IF
        NEXT
      END IF
    END IF

' make include overtime fields protected if an amount is selected
   IF Frm(1).PrevFld <> Frm(1).FldNo THEN
     FOR Cnt = 1 TO 12
       IF Choice$(Cnt, 7) = "" THEN EXIT FOR
       ThisFldNo = FldNum("EMPPCT" + LTRIM$(STR$(Cnt)), Fld())
       TempFld$ = LTRIM$(RTRIM$(Form$(ThisFldNo, 0)))
       IF LEN(TempFld$) THEN
         IF INSTR("Amount", TempFld$) = 0 THEN
           Fld(ThisFldNo + 1).Protected = False
           Fld(ThisFldNo + 2).Protected = False
         ELSE
           Fld(ThisFldNo + 1).Protected = False
           Fld(ThisFldNo + 2).Protected = True
         END IF
       ELSE
         Fld(ThisFldNo + 1).Protected = True
         Fld(ThisFldNo + 2).Protected = True
       END IF
     NEXT
   END IF

'---- skip retire type field if retire numb field is blank
    IF Frm(1).PrevFld = EmpRetNoF AND Frm(1).FldNo <> EmpRetNoF THEN
      IF QPTrim$(Form$(EmpRetNoF, 0)) = "" THEN
        Form$(EmpRetTPF, 0) = ""
        Action = 2
        Fld(EmpRetTPF).Protected = True
      ELSE
        Fld(EmpRetTPF).Protected = False
      END IF
    END IF


'---- adjust distribute via hour/percent field to correct type
    IF Frm(1).PrevFld = EmpPTypeF AND Frm(1).FldNo <> EmpPTypeF THEN
      IF Form$(EmpHPF, 0) = " " THEN  'if it blank then
        SELECT CASE ASCII(Form$(EmpPTypeF, 0))
          CASE 72     'hourly
            Form$(EmpHPF, 0) = "H"
          CASE 83     'salaried
            Form$(EmpHPF, 0) = "P"
        END SELECT
      END IF
    END IF

'---- adjust HiRange in fed/sta amount fields according to type selected
'---- (i.e. amount or percent)
    IF (Frm(1).PrevFld = EmpFedO2F AND Frm(1).FldNo <> EmpFedO2F) OR (Frm(1).PrevFld = EmpStaO2F AND Frm(1).FldNo <> EmpStaO2F) THEN
      IF Frm(1).FldEdited THEN
        LSET Form$(Frm(1).PrevFld + 1, 0) = ""
        PrintArray Frm(1).PrevFld + 1, Frm(1).PrevFld + 1, Form$(), Fld()
      END IF
      SELECT CASE ASCII(Form$(Frm(1).PrevFld, 0))
        CASE 65    'as AMOUNT
           Fld(Frm(1).PrevFld + 1).HiRange = 99999.99
        CASE 80    'as PERCENTAGE
           Fld(Frm(1).PrevFld + 1).HiRange = 100
      END SELECT
    END IF


'---- if not editing an existing employee then check for dupe employee number
    'IF NOT Fld(1).Protected THEN
      IF (Frm(1).PrevFld = 1) AND (Frm(1).FldNo <> 1) THEN
        IF RecNo = 0 THEN
          Oh.Shit! = EMPCheckEMPNumber(Form$(1, 0), Frm(1))
        ELSEIF LEN(OrgNum$) AND QPTrim$(Form$(1, 0)) <> OrgNum$ THEN
          Oh.Shit! = EMPCheckEMPNumber(Form$(1, 0), Frm(1))
        END IF

'---- make sure the employee number field is NOT BLANK
      IF EMPNumNull(Form$()) THEN
        SELECT CASE Frm(1).KeyCode
          CASE F0Key, EscKey
            Frm(1).Edited = False
          CASE ELSE
            'Action = 0
            Frm(1).StartEl = 0
            Frm(1).FldNo = 1
            Frm(1).TxtPos = 1
        END SELECT
      END IF
    END IF
 
'--------------------------------------------

    IF Frm(1).Presses AND Frm(1).MRow = 25 THEN

      SELECT CASE Frm(1).MCol
        CASE 3 TO 11           'F1
          PressButton F1Key, 25, 3, 12
        CASE 13 TO 24           'F3 Delete
          PressButton F3Key, 25, 13, 24
        CASE 25 TO 34            'F4 History
            PressButton F4Key, 25, 25, 34
        CASE 35 TO 42            'F7 YTD Totals
            PressButton F7Key, 25, 35, 42
        CASE 43 TO 49           'PageDown
          IF Scr < NumOfScrns THEN
            PressButton PgDnKey, 25, 43, 49
          END IF
        CASE 50 TO 56           'PageUp
          IF Scr > 1 THEN
            PressButton PgUpKey, 25, 50, 56
          END IF
        CASE 58 TO 67           'F10
          PressButton F0Key, 25, 58, 67

        CASE 69 TO 78           'Escape
          PressButton EscKey, 25, 69, 78
      END SELECT
    END IF
    
'*************************************************************************
'-------------------------------------------------------------------------
    SELECT CASE Frm(1).KeyCode
      CASE F0Key, -93, EscKey

        SELECT CASE Frm(1).KeyCode
          CASE EscKey
            IF Frm(1).Edited THEN
              SaveFlag = PromptSaveData
              IF SaveFlag = True THEN
                IF (LEN(OrgNum$) AND QPTrim$(Form$(1, 0)) <> OrgNum$) OR RecNo = 0 THEN
                  IF EMPCheckEMPNumber(Form$(1, 0), Frm(1)) THEN
                    SaveFlag = 1
                  END IF
                END IF
              END IF
            ELSE
              SaveFlag = False
            END IF

          CASE F0Key, -93
            'IF Frm(1).KeyCode = -93 THEN
            '  FOpenS EmpData2Name, D2Handle
            '  FGetRTA D2Handle, EmpData2Rec, CLNG(RecNo), Data2RecLen
            '  'EmpData2Rec.PreNoteFlag = 0
            '  FPutRTA D2Handle, EmpData2Rec, CLNG(RecNo), Data2RecLen
            '  FClose D2Handle
            'END IF
            GOSUB CheckAllRFlds
            IF FldEmpty THEN
              SaveFlag = 1
              GOTO JumpHere
            END IF
            IF Frm(1).Edited THEN
              IF LEN(OrgNum$) AND QPTrim$(Form$(1, 0)) <> OrgNum$ THEN
                IF EMPCheckEMPNumber(Form$(1, 0), Frm(1)) THEN
                  SaveFlag = 1
                ELSE
                  SaveFlag = True
                END IF
              ELSE
                SaveFlag = True
              END IF
            ELSE
              SaveFlag = False
            END IF
        END SELECT
JumpHere:

        SELECT CASE SaveFlag
          CASE True                   'yep save data to disk
            SaveScrn TempScrn()
            DisplayMiscScrn UpdatingDisk
            EndDelay& = PDQTimer& + 19
            IF RecNo = 0 THEN
              EmpSaveNewRec RecNo, Form$()
            ELSE
              EmpSaveOldRec RecNo, Form$()
            END IF
            DO: LOOP UNTIL PDQTimer& > EndDelay&
            DisplayMiscScrn UpdatedOK
            WaitForAction
            RestScrn TempScrn()
            ERASE TempScrn
            ExitFlag = True
          CASE False                  'nope don't save abandon edit
            ExitFlag = True
          CASE 1                      'oops continue editing
            Action = 2
            ExitFlag = False
        END SELECT
      CASE F3Key
         IF RecNo > 0 THEN
           SaveScrn TempScrn()
           OK2Delete = EmpAsk2Delete
           RestScrn TempScrn()

           IF OK2Delete THEN
             FOpenS EmpData2Name, E2Handle
             FGetRTA E2Handle, EmpData2Rec, CLNG(RecNo), Data2RecLen
             EmpData2Rec.Deleted = True

             LSET EmpData2Rec.EmpNo = ""

             FPutRTA E2Handle, EmpData2Rec, CLNG(RecNo), Data2RecLen
             FClose E2Handle
             EmpData2Rec.Deleted = False

             FOpenS EmpData1Name, E1Handle
             FGetRTA E1Handle, EmpData1Rec, CLNG(RecNo), LEN(EmpData1Rec)
             EmpData1Rec.Deleted = True

             LSET EmpData1Rec.EmpNo = ""

             FPutRTA E1Handle, EmpData1Rec, CLNG(RecNo), LEN(EmpData1Rec)
             FClose E1Handle
             EmpData1Rec.Deleted = False
'**********
             REDIM TempNumRec(1)  AS EmpNumType
             FOpenS EMPNumFileName, NHandle
             FOR Cnt = 1 TO IdxNumOfRecs
               FGetRTA D1Handle, TempData1Rec, CLNG(Cnt), Data1RecLen
               IF NOT TempData1Rec.Deleted THEN
                 TempNumRec(1).EmpNum = TempData1Rec.EmpNo
                 FPutRTA NHandle, TempNumRec(1), CLNG(Cnt), 10
               END IF
             NEXT
             FClose NHandle
             ERASE TempNumRec
             FClose D1Handle
             RecNo = -100
             DisplayMiscScrn UpdatedOK
             WaitForAction
             ExitFlag = True
           ELSE
             Action = 2
           END IF
         END IF
      CASE F4Key
        IF RecNo > 0 THEN
          EMPShowTransHist RecNo
        END IF
      CASE F7Key
        IF RecNo > 0 THEN
          EMPShowYTDTotals RecNo
        END IF
      CASE ELSE

    END SELECT

'---- adjust to next/previous page if needed
    IF Frm(1).FldNo > PageInfo(Scr).LastFld OR Frm(1).FldNo < PageInfo(Scr).FirstFld THEN
      SELECT CASE Frm(1).FldNo
        CASE PageInfo(1).FirstFld TO PageInfo(1).LastFld
          Scr = 1
        CASE PageInfo(2).FirstFld TO PageInfo(2).LastFld
          Scr = 2
        CASE PageInfo(3).FirstFld TO PageInfo(3).LastFld
          Scr = 3
      END SELECT
      Action = 2
      
      LibFile2Scrn EmployeeQLib, FormName$(Scr), MonoCode, -2, ErrCode
      'Lib2Scrn FormName$(Scr), ScrnLib(), MonoCode, -2, ErrCode
      SELECT CASE Scr
        CASE 2
          FOR Cnt = 1 TO 12
            QPrintRC Choice$(Cnt, 7), Cnt + 11, 7, Cnf.PulBar
          NEXT
        CASE 3
          FOR Cnt = 1 TO 3
            QPrintRC Choice$(Cnt, 8), Cnt + 3, 7, Cnf.PulBar
          NEXT
      END SELECT
      EMPAddShowName Form$(EMPFirstFld, 0), Form$(EMPLastFld, 0)
    END IF
  '*************************************************************************
  'Keep editing until the user presses
  LOOP UNTIL ExitFlag = True
  CursorOff

'09-07-95 ***
'This was added to kill a bug in the picklist when a record was
'viewed after a rec was edited. (picking up the prior emp number)
  IF NOT SaveFlag THEN
    RecNo = 0
  END IF

  ERASE Frm, ReqFld, TempScrn, FormName$
  ERASE Form$, Fld, PageInfo, Choice$

  EXIT SUB

CheckAllRFlds:
  FOR Cnt2 = 1 TO NumFlds
    FOR Cnt = 1 TO ReqNumOfFlds
      IF ReqFld(Cnt) = Cnt2 THEN
        FldEmpty = CheckFldEmpty((Form$(Cnt2, 0)))
        IF FldEmpty THEN
          Frm(1).FldNo = Cnt2
          Frm(1).PrevFld = Cnt2
          Frm(1).Presses = 0
          Frm(1).KeyCode = 0
          Frm(1).StartEl = 0
          SaveScrn TempScrn()
          DisplayMiscScrn NoBlank
          WaitForAction
          RestScrn TempScrn()
          ExitLoopFlag = True
        END IF
      END IF
      IF ExitLoopFlag THEN EXIT FOR
    NEXT
    IF ExitLoopFlag THEN EXIT FOR
  NEXT
  ExitLoopFlag = False
  RETURN

END SUB

SUB EMPLoadRFldData (ReqFld(), ReqNumOfFlds)
  RecLen = 2
  NumOfRecs = FileSize(ReqFldFileName) \ RecLen
  REDIM ReqFld(1 TO NumOfRecs)
  FGetAH ReqFldFileName, ReqFld(1), 2, NumOfRecs
  ReqNumOfFlds = NumOfRecs
END SUB

SUB EmpLookUpEdit

  REDIM Frm(1) AS FormInfo

  'REDIM GoodRecs$(0)

  REDIM TempScrn(0)

  FormName$ = "ELUMAIN"
  NumFlds = LibNumberOfFields(EmployeeQLib, FormName$)

  REDIM Form$(NumFlds, 2)             'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo     'DIM the field information array
  LookUpType = NameLookUp
  StartEl = 0                         'Load first form at array start
  LibGetFldDef EmployeeQLib, FormName$, StartEl, Fld(), Form$(), ErrCode

  '----- Setup TYPE for setting and reading form editing information.
  Frm(1).FldNo = 1                 'Start editing on field #1
  Frm(1).InsStat = False           'Set the insert state (-1 = Insert on)
  Frm(1).StartEl = 0               'Set form starting element to 0 and

  '----- Set the "Action" flag to force the editor to initialize itself and
  '      display the data on the form.
  Action = 1


  BlockClear

  LibFile2Scrn EmployeeQLib, FormName$, MonoCode, -1, ErrCode

  GOSUB SetUpLookUp

  DO                 'Pole the editing procedure
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    SELECT CASE Frm(1).KeyCode
      CASE F5Key
        CursorOff
        SearchDataFile LookUpType, Form$(), MatchCnt
        SaveScrn TempScrn()
        SELECT CASE MatchCnt
          CASE -1     'blank search spec
            DisplayMiscScrn EMPBlankLookUp
            WaitForAction
          CASE 0
            DisplayMiscScrn EMPNotFound
            WaitForAction
          CASE 1
            '06-16-94
            'added to go directly to editing the rec if only ONE record
            'matched the search criteria
            EdRecNo = CVI(RIGHT$(Items(1).V, 2))
            EMPEditRec EdRecNo
            '03-30-95 added to cure a bug
            Action = 2
          CASE ELSE        'matched more than 1, give them the list
            EMPChooseList MatchCnt, LookUpType
            Action = 2
            'REDIM GoodRecs$(0)
        END SELECT
        RestScrn TempScrn()
      CASE F7Key                         'F2 = Save
        LookUpType = NOT LookUpType
        GOSUB SetUpLookUp
      CASE ELSE

    END SELECT

    IF Frm(1).Presses AND Frm(1).MRow = ELUBtnRow THEN

      SELECT CASE Frm(1).MCol
        CASE 2 TO 11           'F1
          PressButton F1Key, ELUBtnRow, 2, 11

        CASE 13 TO 23           'F5
          PressButton F5Key, ELUBtnRow, 13, 23

        CASE 35 TO 45           'F7
          PressButton F7Key, ELUBtnRow, 35, 45

        CASE 69 TO 79           'escape
          PressButton EscKey, ELUBtnRow, 69, 79
      END SELECT

    END IF

    '*************************************************************************
    'Keep editing until the user presses
  LOOP UNTIL Frm(1).KeyCode = EscKey   '  the Escape key.

  CursorOff

  ERASE Frm, Fld, Form$, TempScrn

  EXIT SUB

SetUpLookUp:

  SELECT CASE LookUpType
    CASE True
      Fld(3).Protected = False
      Fld(1).Protected = False
      Fld(2).Protected = True
      LookTypeScrn$ = "LOOKNAME"
      F7Type$ = "Name  "
    CASE ELSE
      Fld(3).Protected = True
      Fld(1).Protected = True
      Fld(2).Protected = False
      LookTypeScrn$ = "LOOKNUMB"
      F7Type$ = "Number"
  END SELECT
'******************************************************
  LibFile2Scrn EmployeeQLib, LookTypeScrn$, MonoCode, -2, ErrCode
  QPrintRC F7Type$, 25, 39, Cnf.HiLite
  PrintArray 1, 3, Form$(), Fld()
RETURN

END SUB

FUNCTION EMPNumLeaveRec%
  REDIM TLeaveRec(1) AS LeaveRecType
  LLen = LEN(TLeaveRec(1))
  EMPNumLeaveRec = FileSize(LeaveFileName) \ LLen
  ERASE TLeaveRec
END FUNCTION

FUNCTION EMPNumNull% (Form$())
  IF LTRIM$(RTRIM$(Form$(1, 0))) = "" THEN
    EMPNumNull = True
  ELSE
    EMPNumNull = False
  END IF
END FUNCTION

FUNCTION EmpNumOfRecs
  EmpNumOfRecs = FileSize(EmpIdxLName) \ 2
END FUNCTION

SUB EmpSaveNewRec (RecNo%, Form$())
  
  REDIM TempIDXRec(1)    AS EmployeeIndexType

  'DIM TempIdxNRec    AS EmployeeIndexType

  REDIM TempNumRec(1)  AS EmpNumType

  REDIM Check(1)     AS PRCheckRecType

  IdxRecLen = LEN(TempIDXRec(1))
  IdxFileSize& = FileSize(EmpIdxLName)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

  Data1RecLen = LEN(TempData1Rec)
  Data2RecLen = LEN(EmpData2Rec)
  Data3RecLen = LEN(EmpData3Rec)
  TTRecLen = LEN(TransRec)
  CheckRecLen = LEN(Check(1))

  DataFileSize& = FileSize(EmpData1Name)
  DataNumOfRecs& = DataFileSize& \ Data1RecLen

  NextFreeRec& = DataNumOfRecs& + 1

  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(EmpData2Rec), VARPTR(EmpData2Rec), LEN(Form$(0, 0)), 0
  'copy the employee's combined data to our type variable

  FGetAH EMPPinFileName, LastPin, 2, 1
  NextPin = LastPin + 1
  EmpData2Rec.EmpPin = NextPin
  FPutAH EMPPinFileName, NextPin, 2, 1

  RSET EmpData2Rec.EmpNo = QPTrim$(EmpData2Rec.EmpNo)

'09-29-94 Fix to make sure the pointer on new employee is set to ZERO
'cured abscure bug in trans history when adding new after edit old emp

  EmpData2Rec.LastTransRec = 0       '****************

  EmpData1Rec.EmpNo = EmpData2Rec.EmpNo
  EmpData1Rec.EMPLNAME = EmpData2Rec.EMPLNAME
  EmpData1Rec.EMPFNAME = EmpData2Rec.EMPFNAME

  EmpData1Rec.Data1RecNum = NextFreeRec&
  EmpData1Rec.TransRecNum = NextFreeRec&

  EmpData3Rec.Data1RecNum = NextFreeRec&

  TransRec.EmpPin = EmpData2Rec.EmpPin

  TempNumRec(1).EmpNum = EmpData2Rec.EmpNo

  IF IdxNumOfRecs THEN                        'if there are any recs
    FOpenS EmpData2Name, D2Handle
    FPutRTA D2Handle, EmpData2Rec, NextFreeRec&, Data2RecLen
    FClose D2Handle

    FOpenS TransWorkFileName, THandle   'open it
    FPutRTA THandle, TransRec, NextFreeRec&, TTRecLen
    FClose THandle

' 02/22/94   added emp data file 3
    FOpenS EmpData3Name, D3Handle
    FPutRTA D3Handle, EmpData3Rec, NextFreeRec&, Data3RecLen
    FClose D3Handle

    FOpenS EMPNumFileName, Handle
    FPutRTA Handle, TempNumRec(1), NextFreeRec&, 10
    FClose Handle

    FOpenS ChecksFileName, Handle
    FPutRTA Handle, Check(1), NextFreeRec&, CheckRecLen
    FClose Handle

    REDIM IdxBuff(1 TO IdxNumOfRecs + 1) AS EmployeeIndexType
    'dim array to num of recs +1
    FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, IdxNumOfRecs
    'load it

    FOpenS EmpData1Name, D1Handle             'open data file

    LowOffSet = 0                          '
    TopOffSet = IdxNumOfRecs + 1
    RecOffSet = IdxNumOfRecs \ 2

    IF RecOffSet < 1 THEN RecOffSet = 1

    DO

      FGetRTA D1Handle, TempData1Rec, CLNG(IdxBuff(RecOffSet).DataRecNum), Data1RecLen
      'get the rec from the data file based off the index rec num

      SELECT CASE Compare3%(VARSEG(EmpData1Rec), VARPTR(EmpData1Rec.EMPLNAME), VARSEG(TempData1Rec), VARPTR(TempData1Rec.EMPLNAME), 24)
        CASE -1          'Lower
          TopOffSet = RecOffSet
          RecOffSet = RecOffSet \ 2
          IF RecOffSet < LowOffSet THEN
            RecOffSet = LowOffSet + ((TopOffSet - LowOffSet) \ 2)
          END IF
        CASE 1           'Higher
          LowOffSet = RecOffSet
          RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
        CASE 0           'Equal
          SELECT CASE Compare3%(VARSEG(EmpData1Rec), VARPTR(EmpData1Rec.EMPFNAME), VARSEG(TempData1Rec), VARPTR(TempData1Rec.EMPFNAME), 24)
            CASE -1          'Lower
              TopOffSet = RecOffSet
              RecOffSet = RecOffSet \ 2
              IF RecOffSet < LowOffSet THEN
                RecOffSet = LowOffSet + ((TopOffSet - LowOffSet) \ 2)
              END IF
            CASE 1           'Higher
              LowOffSet = RecOffSet
              RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
            CASE 0
              LowOffSet = RecOffSet
              RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
        END SELECT
      END SELECT
    LOOP WHILE TopOffSet - LowOffSet > 1

    InsertRecAt = TopOffSet

    AmtToMove = (IdxNumOfRecs + 1) - TopOffSet

    IF AmtToMove THEN
      InsertT IdxBuff(InsertRecAt), IdxRecLen, AmtToMove
    END IF
    IdxBuff(InsertRecAt).DataRecNum = NextFreeRec&
    FPutAH EmpIdxLName, IdxBuff(1), IdxRecLen, IdxNumOfRecs + 1

'--------------------------------------------
'-sort via employee number

    FGetAH EmpIdxNName, IdxBuff(1), IdxRecLen, IdxNumOfRecs
    'load the Employee number index

    LowOffSet = 0                          '
    TopOffSet = IdxNumOfRecs + 1
    RecOffSet = IdxNumOfRecs \ 2

    IF RecOffSet < 1 THEN RecOffSet = 1

    DO

      FGetRTA D1Handle, TempData1Rec, CLNG(IdxBuff(RecOffSet).DataRecNum), Data1RecLen
      'get the rec from the data file based off the index rec num

      SELECT CASE Compare3%(VARSEG(EmpData1Rec), VARPTR(EmpData1Rec.EmpNo), VARSEG(TempData1Rec), VARPTR(TempData1Rec.EmpNo), 10)
      CASE -1          'Lower
        TopOffSet = RecOffSet
        RecOffSet = RecOffSet \ 2
        IF RecOffSet < LowOffSet THEN
          RecOffSet = LowOffSet + ((TopOffSet - LowOffSet) \ 2)
        END IF
      CASE 1           'Higher
        LowOffSet = RecOffSet
        RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
      CASE 0  'this case should never be true
        TopOffSet = RecOffSet
        RecOffSet = RecOffSet \ 2
        IF RecOffSet < LowOffSet THEN
          RecOffSet = LowOffSet + ((TopOffSet - LowOffSet) \ 2)
        END IF
      END SELECT
    LOOP WHILE TopOffSet - LowOffSet > 1

    InsertRecAt = TopOffSet

    AmtToMove = (IdxNumOfRecs + 1) - TopOffSet

    IF AmtToMove THEN
      InsertT IdxBuff(InsertRecAt), IdxRecLen, AmtToMove
    END IF
    IdxBuff(InsertRecAt).DataRecNum = NextFreeRec&
    FPutAH EmpIdxNName, IdxBuff(1), IdxRecLen, IdxNumOfRecs + 1
    FPutRTA D1Handle, EmpData1Rec, NextFreeRec&, Data1RecLen


'-============================================================
  ELSE
    InsertRecAt = 1
    TempIDXRec(1).DataRecNum = 1
    FPutAH EMPNumFileName, TempNumRec(1), 10, 1&
    FPutAH EmpIdxLName, TempIDXRec(1), IdxRecLen, 1&
    FPutAH EmpIdxNName, TempIDXRec(1), IdxRecLen, 1&

    FOpenS EmpData2Name, D2Handle
    FPutRTA D2Handle, EmpData2Rec, 1&, Data2RecLen
    FClose D2Handle

'add a blank trans work/temp rec
    FOpenS TransWorkFileName, THandle   'open it
    FPutRTA THandle, TransRec, 1&, TTRecLen
    FClose THandle

' 02/22/94   added emp data file 3
    FOpenS EmpData3Name, D3Handle
    FPutRTA D3Handle, EmpData3Rec, 1&, Data3RecLen
    FClose D3Handle

'add a blank check rec
    FOpenS ChecksFileName, Handle
    FPutRTA Handle, Check(1), 1&, CheckRecLen
    FClose Handle

    FOpenS EmpData1Name, D1Handle             'open data file
    FPutRTA D1Handle, EmpData1Rec, 1&, Data1RecLen
  END IF

  FClose D1Handle



END SUB

SUB EmpSaveOldRec (RecNo%, Form$())

  IdxRecLen = 2 'LEN(TempIdxLRec)
  IdxFileSize& = FileSize(EmpIdxLName)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

  Data1RecLen = LEN(TempData1Rec)
  Data2RecLen = LEN(EmpData2Rec)

  DataFileSize& = FileSize(EmpData1Name)
  DataNumOfRecs& = DataFileSize& \ Data1RecLen

  'NextFreeRec& = DataNumOfRecs& + 1

  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(EmpData2Rec), VARPTR(EmpData2Rec), LEN(Form$(0, 0)), 0
  'copy the employee's combined data to our type variable

  RSET EmpData2Rec.EmpNo = QPTrim$(EmpData2Rec.EmpNo)

  EmpData1Rec.EmpNo = EmpData2Rec.EmpNo
  EmpData1Rec.EMPLNAME = EmpData2Rec.EMPLNAME
  EmpData1Rec.EMPFNAME = EmpData2Rec.EMPFNAME
  EmpData1Rec.Data1RecNum = RecNo

  FOpenS EmpData2Name, D2Handle
  FPutRTA D2Handle, EmpData2Rec, CLNG(RecNo), Data2RecLen
  FClose D2Handle

  FOpenS EmpData1Name, D1Handle

  FGetRTA D1Handle, TempData1Rec, CLNG(RecNo), Data1RecLen

  IF EmpData1Rec.EMPLNAME <> TempData1Rec.EMPLNAME OR EmpData1Rec.EMPFNAME <> TempData1Rec.EMPFNAME THEN
    NameSort = True
  END IF

  IF EmpData1Rec.EmpNo <> TempData1Rec.EmpNo THEN
    NumbSort = True
  END IF

  IF NameSort OR NumbSort THEN
    FPutRTA D1Handle, EmpData1Rec, CLNG(RecNo), Data1RecLen
    'dim array to num of recs
  ELSE
    FClose D1Handle
    EXIT SUB
  END IF

'both data files are saved by here

  REDIM IdxBuff(1 TO IdxNumOfRecs) AS EmployeeIndexType
  IF NameSort THEN
    FGetAH EmpIdxLName, IdxBuff(1), IdxRecLen, IdxNumOfRecs
    'load name index
    FOR Cnt = 1 TO IdxNumOfRecs
      IF IdxBuff(Cnt).DataRecNum = RecNo THEN
        DeleteT IdxBuff(Cnt), IdxRecLen, (IdxNumOfRecs - Cnt)
        IdxBuff(IdxNumOfRecs).DataRecNum = 0
        EXIT FOR
      END IF
    NEXT

    LowOffSet = 0                          '
    TopOffSet = IdxNumOfRecs
    RecOffSet = IdxNumOfRecs \ 2

    IF RecOffSet < 1 THEN RecOffSet = 1

    DO

      FGetRTA D1Handle, TempData1Rec, CLNG(IdxBuff(RecOffSet).DataRecNum), Data1RecLen
      'get the rec from the data file based off the index rec num

      SELECT CASE Compare3%(VARSEG(EmpData1Rec), VARPTR(EmpData1Rec.EMPLNAME), VARSEG(TempData1Rec), VARPTR(TempData1Rec.EMPLNAME), 24)
        CASE -1          'Lower
          TopOffSet = RecOffSet
          RecOffSet = RecOffSet \ 2
          IF RecOffSet < LowOffSet THEN
            RecOffSet = LowOffSet + ((TopOffSet - LowOffSet) \ 2)
          END IF
        CASE 1           'Higher
          LowOffSet = RecOffSet
          RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
        CASE 0           'Equal
          SELECT CASE Compare3%(VARSEG(EmpData1Rec), VARPTR(EmpData1Rec.EMPFNAME), VARSEG(TempData1Rec), VARPTR(TempData1Rec.EMPFNAME), 24)
            CASE -1          'Lower
              TopOffSet = RecOffSet
              RecOffSet = RecOffSet \ 2
              IF RecOffSet < LowOffSet THEN
                RecOffSet = LowOffSet + ((TopOffSet - LowOffSet) \ 2)
              END IF
            CASE 1           'Higher
              LowOffSet = RecOffSet
              RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
            CASE 0
              LowOffSet = RecOffSet
              RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
        END SELECT
      END SELECT
    LOOP WHILE TopOffSet - LowOffSet > 1

    InsertRecAt = TopOffSet

    AmtToMove = (IdxNumOfRecs + 1) - TopOffSet

    IF AmtToMove THEN
      InsertT IdxBuff(InsertRecAt), IdxRecLen, AmtToMove
    END IF
    IdxBuff(InsertRecAt).DataRecNum = RecNo

    FPutAH EmpIdxLName, IdxBuff(1), IdxRecLen, IdxNumOfRecs
  END IF
'--------------------------------------------
'-sort via employee number
  IF NumbSort THEN
    FGetAH EmpIdxNName, IdxBuff(1), IdxRecLen, IdxNumOfRecs
    FOR Cnt = 1 TO IdxNumOfRecs
      IF IdxBuff(Cnt).DataRecNum = RecNo THEN
        DeleteT IdxBuff(Cnt), IdxRecLen, (IdxNumOfRecs - Cnt)
        EXIT FOR
      END IF
    NEXT

    'load the Employee number index

    LowOffSet = 0                          '
    TopOffSet = IdxNumOfRecs
    RecOffSet = IdxNumOfRecs \ 2

    IF RecOffSet < 1 THEN RecOffSet = 1

    DO

      FGetRTA D1Handle, TempData1Rec, CLNG(IdxBuff(RecOffSet).DataRecNum), Data1RecLen
      'get the rec from the data file based off the index rec num

      SELECT CASE Compare3%(VARSEG(EmpData1Rec), VARPTR(EmpData1Rec.EmpNo), VARSEG(TempData1Rec), VARPTR(TempData1Rec.EmpNo), 10)
      CASE -1          'Lower
        TopOffSet = RecOffSet
        RecOffSet = RecOffSet \ 2
        IF RecOffSet < LowOffSet THEN
          RecOffSet = LowOffSet + ((TopOffSet - LowOffSet) \ 2)
        END IF
      CASE 1           'Higher
        LowOffSet = RecOffSet
        RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
      CASE 0
        TopOffSet = RecOffSet
        RecOffSet = RecOffSet \ 2
        IF RecOffSet < LowOffSet THEN
          RecOffSet = LowOffSet + ((TopOffSet - LowOffSet) \ 2)
        END IF
      END SELECT
    LOOP WHILE TopOffSet - LowOffSet > 1

    InsertRecAt = TopOffSet

    AmtToMove = (IdxNumOfRecs + 1) - TopOffSet

    IF AmtToMove THEN
      InsertT IdxBuff(InsertRecAt), IdxRecLen, AmtToMove
    END IF
    IdxBuff(InsertRecAt).DataRecNum = RecNo
    FPutAH EmpIdxNName, IdxBuff(1), IdxRecLen, IdxNumOfRecs
    'FPutRTA D1Handle, EmpData1Rec, RecNo, Data1RecLen

  END IF
'-============================================================

  IF NumbSort THEN
    REDIM TempNumRec(1)  AS EmpNumType
    FOpenS EMPNumFileName, NHandle
    FOR Cnt = 1 TO IdxNumOfRecs
      FGetRTA D1Handle, TempData1Rec, CLNG(Cnt), Data1RecLen
      TempNumRec(1).EmpNum = TempData1Rec.EmpNo
      FPutRTA NHandle, TempNumRec(1), CLNG(Cnt), 10
    NEXT
    FClose NHandle
    ERASE TempNumRec
  END IF

  FClose D1Handle
END SUB

SUB EMPShowTransHist (RecNo%)

  CursorOff
  REDIM TempScrn(0)
  SaveScrn TempScrn()

  'REDIM TItems(1)

  REDIM TEmp2Rec(1)   AS EmpData2Type
  REDIM TTransRec(1)  AS TransRecType


  REDIM TempData(1)   AS STRING * 45
  REDIM Temp1(1)  AS STRING * 6
  REDIM Temp2(1)  AS STRING * 10

  REDIM TItems(1 TO 150) AS FLen

  Emp2RecLen = LEN(TEmp2Rec(1))
  TRecSize = LEN(TTransRec(1))

  FOpenS EmpData2Name, D2Handle
  FGetRTA D2Handle, TEmp2Rec(1), CLNG(RecNo), Emp2RecLen
  FClose D2Handle

  IF TEmp2Rec(1).LastTransRec > 0 THEN
    GOSUB LoadTransHist
  ELSE
    DisplayMiscScrn NoTransHist
    WaitForAction
    GOTO HistExit
  END IF

  VAction = 0

  BoxBot = 17
  BoxTop = 8

  MaxLen = 45
  LeftCol = ((80 - MaxLen) \ 2) - 1

'    OMenBox = Cnf.MenBox
'    OActivCh = Cnf.ActivCh
'    Cnf.MenBox = 30
'    Cnf.ActivCh = 31

    HideCursor
    LOCATE BoxTop, LeftCol
    QPrintRC " Chk No.   Check Date   Period End    Net Pay    ", BoxTop - 1, LeftCol, Cnf.HiLite
    MPaintBox BoxTop, LeftCol + MaxLen + 4, BoxTop, LeftCol + MaxLen + 5, 8

    WazzWind 20, LeftCol, 22, LeftCol + MaxLen + 3, Cnf.HiLite, 2, True
    QPrintRC "Press ESC to continue.", 21, LeftCol + (MaxLen \ 2) - 8, Cnf.HiLite

    ShowCursor

    VertMenuT TItems(), Choice, MaxLen, BoxBot, Ky$, VAction, Cnf
    'VertMenu GoodRecs$(), Choice, MaxLen, BoxBot, Ky$, VAction, Cnf

'    Cnf.MenBox = OMenBox
'    Cnf.ActivCh = OActivCh
    
'    CursorOff

HistExit:
  RestScrn TempScrn()
  ERASE TempScrn, TEmp2Rec, TTransRec, TempData, Temp1, Temp2, TItems
  'Pause 3

EXIT SUB

LoadTransHist:

  PrevTransRecNum& = TEmp2Rec(1).LastTransRec
  FOpenS TransHistFileName, THandle

  DO
    MatchCnt = MatchCnt + 1
    FGetRTA THandle, TTransRec(1), PrevTransRecNum&, TRecSize
'    PRINT TTransRec(1).EmpPin

    RSET Temp1(1) = QPStrL$(TTransRec(1).CheckNum)
    LSET TempData(1) = Temp1(1)
    MID$(TempData(1), 10) = Num2Date(TTransRec(1).CheckDate)
    MID$(TempData(1), 23) = Num2Date(TTransRec(1).PayPdEnd)
    RSET Temp2(1) = FUsing$(STR$(TTransRec(1).NetPay), "##,###.##")
    'RSET Temp2(1) = FUsing$(STR$(12345.55), "##,###.##")
    MID$(TempData(1), 34) = Temp2(1)
    TItems(MatchCnt).V = TempData(1)
    PrevTransRecNum& = TTransRec(1).PrevTransRec
  LOOP WHILE PrevTransRecNum& AND MatchCnt < 150
  FClose THandle
  IF MatchCnt > 1 THEN
    REDIM PRESERVE TItems(1 TO MatchCnt) AS FLen
'    SortT TItems(1), MatchCnt, 0, LEN(TItems(1).V), 0, 8
    'Sort GoodRecs$(), 1
  END IF

RETURN

END SUB

SUB EMPShowYTDTotals (RecNo)

  SHARED Choice$()

  REDIM TempScrn(0)
  SaveScrn TempScrn()

  REDIM TEmp3Rec(1) AS EmpData3Type
  Emp3RecLen = LEN(TEmp3Rec(1))
  FOpenS EmpData3Name, D3Handle
  FGetRTA D3Handle, TEmp3Rec(1), CLNG(RecNo), Emp3RecLen
  FClose D3Handle

  FormName$ = "EMPYTD"
  NumFlds = LibNumberOfFields(EmployeeQLib, FormName$)
  REDIM Form$(NumFlds, 2)             'DIM the form data array
  REDIM Fld(NumFlds) AS FieldInfo     'DIM the field information array
  StartEl = 0                         'Load first form at array start
  LibGetFldDef EmployeeQLib, FormName$, StartEl, Fld(), Form$(), ErrCode

  Form$(0, 0) = SPACE$(Emp3RecLen)

  BCopy VARSEG(TEmp3Rec(1)), VARPTR(TEmp3Rec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), Emp3RecLen, 0
  UnPackBuffer 0, 0, Form$(), Fld()

  WazzWind 3, 3, 22, 77, Cnf.MenBox, 1, 1'True
  LibFile2Scrn EmployeeQLib, FormName$, MonoCode, -1, ErrCode

  Temp$ = SPACE$(10)

  StartCol = 5
  FOR Cnt = 1 TO 6
    RSET Temp$ = Choice$(Cnt, 7)
    QPrintRC Temp$, 16, StartCol, Cnf.ActivCH
    RSET Temp$ = Choice$(Cnt + 6, 7)
    QPrintRC Temp$, 18, StartCol, Cnf.ActivCH
    StartCol = StartCol + 12
  NEXT

  StartCol = 36
  FOR Cnt = 1 TO 3
    RSET Temp$ = Choice$(Cnt, 8)
    QPrintRC Temp$, 6, StartCol, Cnf.ActivCH
    StartCol = StartCol + 15
  NEXT

  PrintArray 1, NumFlds, Form$(), Fld()
  WaitForAction
  RestScrn TempScrn()

  ERASE TempScrn, Form$, Fld, TEmp3Rec

END SUB

SUB InitEmpMulti (Fld() AS FieldInfo)

  SHARED Choice$(), TaxText$()

  REDIM ERNCodeRec(1 TO 3)   AS ErnCodeRecType
  REDIM RetireRec(1 TO 6)    AS RetireRecType
  REDIM DedCodeRec(1 TO 12)  AS DedCodeRecType
  REDIM Choice$(14, 11)                      '

  Choice$(0, 0) = STR$(FldNum("EMPGENDR", Fld()))
  Choice$(0, 1) = STR$(FldNum("EMPRACE", Fld()))
  Choice$(0, 2) = STR$(FldNum("EMPSTATS", Fld()))
  Choice$(0, 3) = STR$(FldNum("EMPPTYPE", Fld()))
  Choice$(0, 4) = STR$(FldNum("EMPPFREQ", Fld()))
  Choice$(0, 5) = STR$(FldNum("EMPEIC", Fld()))
  Choice$(0, 6) = STR$(FldNum("EMPRETTP", Fld()))
  Choice$(0, 9) = STR$(FldNum("EMPSTAS", Fld()))
  Choice$(0, 11) = STR$(FldNum("EMPFEDS", Fld()))

'----------------------------
  Choice$(1, 0) = "Male"
  Choice$(2, 0) = "Female"

'  Choice$(1, 1) = "Asian"
'  Choice$(2, 1) = "Black"
'  Choice$(3, 1) = "Hispanic"
'  Choice$(4, 1) = "Native American"
'  Choice$(5, 1) = "White"

  Choice$(1, 2) = "Full-Time"
  Choice$(2, 2) = "Part-Time"
  Choice$(3, 2) = "Temporary"
  Choice$(4, 2) = "Seasonal "

  Choice$(1, 3) = "Salaried"
  Choice$(2, 3) = "Hourly"

  Choice$(1, 4) = "Weekly"
  Choice$(2, 4) = "Bi-Weekly"
  Choice$(3, 4) = "Semi-Monthly"
  Choice$(4, 4) = "Monthly"
  Choice$(5, 4) = "Quarterly"
  Choice$(6, 4) = "Semi-Annually"
  Choice$(7, 4) = "Annually"

  Choice$(1, 5) = "0 - No Certificate"
  Choice$(2, 5) = "1 - Employee Only"
  Choice$(3, 5) = "2 - Employee & Spouse"

  Choice$(1, 9) = TaxText$(1)  '"Single"
  Choice$(2, 9) = TaxText$(2)  '"Married"
  Choice$(3, 9) = TaxText$(3)  '"Head of Household"
  FOR Cnt2 = 1 TO 3
    IF Choice$(Cnt2, 9) = "N/A" THEN Choice$(Cnt2, 9) = ""
  NEXT

  Choice$(1, 11) = "Single"
  Choice$(2, 11) = "Married"
  Choice$(3, 11) = "Head of Household"

  Choice$(1, 10) = "Amount"
  Choice$(2, 10) = "Percent"
  
  FGetAH RetireFileName, RetireRec(1), LEN(RetireRec(1)), 6
  FOR Cnt = 1 TO 6
    Choice$(Cnt, 6) = QPTrim$(RetireRec(Cnt).TYPEDES1)
  NEXT

  FGetAH DedCodeFileName, DedCodeRec(1), LEN(DedCodeRec(1)), 12

  Choice$(0, 7) = ""
  FOR Cnt = 1 TO 12
    Choice$(Cnt, 7) = QPTrim$(DedCodeRec(Cnt).DCDESC1)
    IF Choice$(Cnt, 7) = "" THEN
      ErnFldOffSet = FldNum("EMPPCT" + LTRIM$(STR$(Cnt)), Fld())
      FOR zz = ErnFldOffSet TO ErnFldOffSet + 3
        Fld(zz).Protected = True
      NEXT
 '     Choice$(Cnt, 7) = "NOT-USED"
    END IF
  NEXT

  FGetAH ErnCodeFileName, ERNCodeRec(1), LEN(ERNCodeRec(1)), 3
  FOR Cnt = 1 TO 3
    Choice$(Cnt, 8) = QPTrim$(ERNCodeRec(Cnt).ERNCODE1)
    IF Choice$(Cnt, 8) = "" THEN
      ErnFldOffSet = FldNum("EMPEACT" + LTRIM$(STR$(Cnt)), Fld())
      FOR zz = ErnFldOffSet TO ErnFldOffSet + 2
        Fld(zz).Protected = True
      NEXT
'      Choice$(Cnt, 8) = "NOT-USED"
    END IF
  NEXT
  Fld(zz).Protected = False

  FOR Cnt = 1 TO 12
    Choice$(0, 10) = Choice$(0, 10) + STR$(FldNum("EMPPCT" + LTRIM$(STR$(Cnt)), Fld())) + ","
  NEXT

  FOR Cnt = 1 TO 3
    Choice$(0, 10) = Choice$(0, 10) + STR$(FldNum("EMPEPCT" + LTRIM$(STR$(Cnt)), Fld())) + ","
  NEXT

  ERASE ERNCodeRec, RetireRec, DedCodeRec

END SUB

SUB PREmpMenu

  Choice = 1

  DO
    PayMenu EmpMaintMenu, Choice, 5

    SELECT CASE Choice
      CASE 1
'        IF EmpNumOfRecs < 3 THEN
          EMPEditRec 0
'        ELSE
'          DisplayMiscScrn "DEMOONLY"
'          WaitForAction
'          'show demo version screen
'        END IF

      CASE 2 TO 6
        IF EmpNumOfRecs = 0 THEN
          DisplayMiscScrn EMPNoRecs
          WaitForAction
        ELSE
          SELECT CASE Choice
            CASE 2
              EmpLookUpEdit
            CASE 3
              'EMPrintEmpData            'unrem
            CASE 4
              'EMPrintEmpList            'unrem
            CASE 5
              'EMPrintTermEmpList        'unrem
          END SELECT
        END IF
    END SELECT
  LOOP UNTIL Choice = EscKey

END SUB

SUB SearchDataFile (LookUpType%, Form$(), MatchCnt%)

  REDIM TempScrn(0)
  SaveScrn TempScrn()

  MatchCnt = False

  REDIM TempData(1) AS STRING * 43
  
'  REDIM GoodRecs$(1 TO EmpNumOfRecs)

  REDIM Items(1 TO EmpNumOfRecs) AS FLen

  FirstName$ = UCASE$(QPTrim$(Form$(3, 0)))
  LastName$ = UCASE$(QPTrim$(Form$(1, 0)))
  EmpNumb$ = UCASE$(QPTrim$(Form$(2, 0)))

  SELECT CASE LookUpType
    CASE NameLookUp
      IF LEN(FirstName$) THEN
        FirstOK = True
        ParmsOK = True
      END IF
      IF LEN(LastName$) THEN
        LastOK = True
        ParmsOK = True
      END IF
    CASE NumbLookUp
      IF LEN(EmpNumb$) THEN
        ParmsOK = True
      END IF
  END SELECT

JumpType:
  SELECT CASE ParmsOK
    CASE True
      FOpenS EmpData1Name, D1Handle
      Data1RecLen = LEN(TempData1Rec)
      NumOfRecs = EmpNumOfRecs
    
      SELECT CASE LookUpType
        CASE NameLookUp
          FOR Cnt = 1 TO NumOfRecs
            FGetA D1Handle, TempData1Rec, CLNG(Data1RecLen)
            IF NOT TempData1Rec.Deleted = True THEN
              'put deleted test here
              IF FirstOK THEN
                IF LEFT$(UCASE$(TempData1Rec.EMPFNAME), LEN(FirstName$)) = FirstName$ THEN
                  MatchOK = True
                ELSE
                  MatchOK = False
                END IF
              END IF
              IF LastOK AND (FirstOK EQV MatchOK) THEN
                IF LEFT$(UCASE$(TempData1Rec.EMPLNAME), LEN(LastName$)) = LastName$ THEN
                  MatchOK = True
                ELSE
                  MatchOK = False
                END IF
              END IF
              IF MatchOK THEN GOSUB ParseToMenu
            END IF
          NEXT

        CASE NumbLookUp

          FOR Cnt = 1 TO NumOfRecs
            FGetA D1Handle, TempData1Rec, CLNG(Data1RecLen)
            IF NOT TempData1Rec.Deleted = True THEN
              IF INSTR(TempData1Rec.EmpNo, EmpNumb$) THEN MatchOK = True
              IF MatchOK THEN GOSUB ParseToMenu
            END IF
          NEXT
      END SELECT

    CASE False
      FOpenS EmpData1Name, D1Handle
      Data1RecLen = LEN(TempData1Rec)
      NumOfRecs = EmpNumOfRecs
      FOR Cnt = 1 TO NumOfRecs
        FGetA D1Handle, TempData1Rec, CLNG(Data1RecLen)
        IF NOT TempData1Rec.Deleted = True THEN
          GOSUB ParseToMenu
        END IF
      NEXT
  END SELECT
  FClose D1Handle

  IF MatchCnt > 1 THEN
    REDIM PRESERVE Items(1 TO MatchCnt) AS FLen
    SortT Items(1), MatchCnt, 0, LEN(Items(1).V), 0, 43
  '  Sort GoodRecs$(), 0
  END IF

  RestScrn TempScrn()

  ERASE TempScrn, TempData

  EXIT SUB

ParseToMenu:
  MatchOK = False
  MatchCnt = MatchCnt + 1
  'tems(MatchCnt). = TempData(1)
  SELECT CASE LookUpType
    CASE NameLookUp
      LSET Items(MatchCnt).V = TempData1Rec.EMPLNAME
      MID$(Items(MatchCnt).V, 18) = QPTrim$(TempData1Rec.EMPFNAME)
      MID$(Items(MatchCnt).V, 30) = RIGHT$(TempData1Rec.EmpNo, 8)
    CASE NumbLookUp
      LSET Items(MatchCnt).V = RIGHT$(TempData1Rec.EmpNo, 8)
      MID$(Items(MatchCnt).V, 12) = QPTrim$(TempData1Rec.EMPLNAME)
      MID$(Items(MatchCnt).V, 30) = QPTrim$(TempData1Rec.EMPFNAME)
  END SELECT
  MID$(Items(MatchCnt).V, 42) = MKI$(TempData1Rec.Data1RecNum)
  RETURN

END SUB

