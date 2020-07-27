DEFINT A-Z
DECLARE FUNCTION RecalcOK% ()
DECLARE SUB RecalcAvgUse ()
DECLARE SUB ZeroOffBalances ()
DECLARE SUB HHCodeMaint ()
DECLARE SUB ResetProrate ()
DECLARE SUB DeleteRateTbl ()
DECLARE FUNCTION LOGPassWordOK% ()
DECLARE FUNCTION RatePassWordOK% ()
DECLARE SUB MakeBKGround ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION nGetPassword$ (Row%, Col%, Max%, Abort%)
DECLARE SUB RelinkWorkOrders ()
DECLARE SUB ubGetAcctStruct (GLFundLen%, GLAcctLen%, GLDetLen%)
DECLARE SUB UBLog (Text$)
DECLARE SUB OpenAcctIdx (AcctIdxFileNum%, NumIdxRecs%)
DECLARE SUB OpenAcctFile (AcctFileNum%, NumAccts%)
DECLARE SUB ReLinkHistory ()
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB PrintRateListing ()
DECLARE SUB ReIndexSystem (PromptFlag%)
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB SmallPause ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB EnterNewBalance ()
DECLARE FUNCTION GetNumOfCust% ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB SetBeginningBalance ()
DECLARE SUB EnterEditBalance ()
DECLARE SUB PrintBalanceTrans ()
DECLARE SUB PostBalanceTrans ()
DECLARE SUB UBRateMaint ()
DECLARE SUB AddEditRateTbl (EDFlag%)
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB LoadRateTblRec (Form$(), TblNum%)
DECLARE SUB SaveRateTblRec (Form$(), RecNo%)
DECLARE SUB SelectRateTbl (RecNo%)
DECLARE SUB EditUBSetUpRec ()
DECLARE SUB LoadSetUpRec (Form$(), OKFlag)
DECLARE SUB SaveSetUpRec (Form$())
DECLARE SUB LoadSysDraftRec (Form$())
DECLARE SUB EditBankDraftRec ()
DECLARE SUB SaveSysDraftRec (Form$())
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB WaitForAction ()
DECLARE SUB EditSysBankDraft ()
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB KillFile (File2Kill$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB QPrintRC (Text$, Row%, Col%, Kolor)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE FUNCTION GetAcctTitle$ (AcctRecNum)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB SortStr2 (BYVAL Address%, Size%, Direction%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Num2Date$ (DateNum)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION ASCII (Character$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Number$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION QPStrL$ (Number&)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE FUNCTION GetNumRateRecs ()
DECLARE SUB ClearScrn ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE FUNCTION Valid% (FileName2Check$)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'ubtrans.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.BI'
 
  CONST False = 0, True = NOT False

  CRLF$ = CHR$(13) + CHR$(10)

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen%
  
  'GOTO StartHere
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 12)
  
  MChoice$(1) = "Utility System Setup"
  MChoice$(2) = "Rate Code Maintenance"
  MChoice$(3) = "Bank Draft Setup"
  MChoice$(4) = "Relink Transactions History"
  MChoice$(5) = "Relink Work Order History"
  MChoice$(6) = "Reindex Utility Files"
  MChoice$(7) = "Reset Prorate Percentages"
  MChoice$(8) = "View Utility Active Log File"
  MChoice$(9) = "Export Customer Information"
  MChoice$(10) = "Recalc Average Consumption"

  IF UBSetUpRec(1).HHDEVICE = "H" THEN
    SelectStop = 12
    MChoice$(11) = "HandHeld Code Maintenance"
    MChoice$(12) = "Exit to OS"
  ELSE
    SelectStop = 11
    MChoice$(11) = "Exit to OS"
  END IF
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 20
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
  Col = ((80 - MaxLen) \ 2)

  IF INSTR(COMMAND$, "REINDEX") > 0 THEN
    RelinkFlag = True
    GOTO RelinkEntry
  END IF

  IF INSTR(COMMAND$, "RELINK") > 0 THEN
    RelinkFlag = True
    GOTO RelinkHist
  END IF

  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 2, Col, MaxLen + 3, "Miscellaneous Functions ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN
      Choice = 0
      ExitFlag = True
      EXIT DO   'choice = 0
    END IF
    
    SELECT CASE Choice
    CASE 1
      RUN "ubsetup"
    CASE 2
'12-09-97 add password to Rate Code Area
      IF WEnvTest% THEN
        UBRateMaint
      ELSEIF RatePassWordOK THEN
        UBRateMaint
      END IF
    CASE 3
      EditBankDraftRec
    CASE 4
RelinkHist:
      IF RelinkFlag THEN
        ReLinkHistory
        GOTO ExitHere
      ELSE
        ReLinkHistory
      END IF
    CASE 5
      RelinkWorkOrders
    CASE 6
RelinkEntry:
      IF RelinkFlag THEN
        ReIndexSystem False
        GOTO ExitHere
      ELSE
        ReIndexSystem True
      END IF
    CASE 7
      ResetProrate
    CASE 8
      IF LOGPassWordOK THEN
        CursorOff
        UBLog " IN: Review LOG"
        PrintRptFile "", "UBLOG.DAT", 1, RetCode%, 2
        UBLog "OUT: Review LOG" + CRLF$
      END IF
    CASE 9
      RUN "UBEXPORT"
    CASE 10
      RecalcAvgUse
      'ZeroOffBalances

    CASE 11
      IF SelectStop = 12 THEN
'        HHCodeMaint
'unrem
      ELSE
        HideCursor
        CLS
        END
      END IF
    CASE 12
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF WEnvTest AND INSTR(COMMAND$, "TEST") THEN
    GOTO ExitHere
  ELSE
    RUN "ubmenu"
  END IF
  
ExitHere:
  HideCursor
  ClearScrn
  END

StartHere:

SUB AddEditRateTbl (EDFlag)

  REDIM UBRateTblRec(1) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))

  ChkFlag$ = CHR$(1)

  UBLog " IN: Add\Edit Rate Code"

  REDIM ScrnArray(0)
  
  LibName$ = "UB"
  ScrnName$ = "UBNRATE"
  NumScrns = 1
  
  '--define the multi-choice fields
  SHARED Choice$()
  REDIM Choice$(0 TO 4, 0)
  
  '--Initialize the form name array
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
    LSET Form$(F, 0) = ""
  NEXT
  
  IF EDFlag THEN
    SelectRateTbl RecNo
    IF RecNo > 0 THEN
      GOSUB LoadRateRec
      UBLog "EDITING RATE:" + STR$(RecNo) + "  " + Form$(1, 0)
      Fld(1).Protected = True
    ELSE
      GOTO ExitEdit
    END IF
  ELSE
    RecNo = GetNumRateRecs + 1
  END IF
  
  '--Set screen number to one and display screen
  Scr = 1
  BlockClear
  DisplayUBScrn ScrnName$
  
  ShowCursor
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      SaveScrn ScrnArray()
      GOSUB ValidateRateCode
      IF BadRate = False THEN
        GOSUB SaveRateRec
        DisplayUBScrn "UPDATEOK"
        WaitForAction
        UBLog "UPDATED RATE:" + STR$(RecNo) + "  " + Form$(1, 0)
        ExitFlag = True
      ELSE
        Action = 1
      END IF
      RestScrn ScrnArray()
    CASE EscKey
      IF EDFlag THEN
        UBLog "ABORTED EDIT: " + Form$(1, 0)
      ELSE
        UBLog "ABORTED ADDING RATE!"
      END IF
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 21   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 36 TO 47           '--Save Button
          PressButton -68, 21, 36, 47
        CASE 50 TO 63           '--cancel button
          PressButton 27, 21, 50, 63
        END SELECT
      END SELECT                'row
    END IF
    
    IF (Frm(1).FldNo > 1) AND (Frm(1).PrevFld = 1) AND NOT EDFlag THEN
      GOSUB Check4DupeRate
    END IF

    '--Check screen page
  LOOP UNTIL ExitFlag

ExitEdit:
  UBLog "OUT: Add\Edit Rate Code" + CRLF$
  HideCursor
EXIT SUB

ValidateRateCode:

  BadRate = False

  IF NOT EDFlag THEN
    GOSUB Check4DupeRate
  END IF

  IF BadRate THEN
    GOTO RateExit
  END IF

  RateCo$ = QPTrim$(Form$(1, 0))

  IF LEN(RateCo$) = 0 THEN
    GOSUB ShowErrorScrn
    Frm(1).FldNo = 1
    GOTO RateExit
  END IF

  MINAMT$ = QPTrim$(Form$(4, 0))
  IF LEN(MINAMT$) > 0 THEN
    MINAMT# = Value#(MINAMT$, ErrCode%)
  ELSE
    MINAMT# = -1000001
  END IF
  IF MINAMT# < -1000000 THEN
    GOSUB ShowErrorScrn
    Frm(1).FldNo = 3
    GOTO RateExit
  END IF

  MINAMT$ = QPTrim$(Form$(4, 0))
  IF LEN(MINAMT$) > 0 THEN
    MINAMT# = Value#(MINAMT$, ErrCode%)
  ELSE
    MINAMT# = -1000001
  END IF
  IF MINAMT# < -1000000 THEN
    GOSUB ShowErrorScrn
    Frm(1).FldNo = 4
    GOTO RateExit
  END IF

  MINAMT$ = QPTrim$(Form$(6, 0))
  IF LEN(MINAMT$) > 0 THEN
    MINAMT# = Value#(MINAMT$, ErrCode%)
  ELSE
    MINAMT# = -1
  END IF
  IF MINAMT# < 0 THEN
    GOSUB ShowErrorScrn
    Frm(1).FldNo = 6
    GOTO RateExit
  END IF

  MINAMT$ = QPTrim$(Form$(7, 0))
  IF LEN(MINAMT$) > 0 THEN
    MINAMT# = Value#(MINAMT$, ErrCode%)
  ELSE
    MINAMT# = -1
  END IF
  IF MINAMT# < 0 THEN
    GOSUB ShowErrorScrn
    Frm(1).FldNo = 7
    GOTO RateExit
  END IF

RateExit:

RETURN

ShowErrorScrn:
  HideCursor
  BadRate = True
  DisplayUBScrn "ERRSCRN1"
  QPrintRC "INVALID OR, BLANK FIELD ENTRY!", 11, 26, -1
  QPrintRC "Please Correct and Save Again.", 13, 26, -1
  ShowCursor
  WaitForAction
RETURN

LoadRateRec:

  Form$(0, 0) = SPACE$(UBRateTblRecLen)

  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  GET UBFile, RecNo, UBRateTblRec(1)
  CLOSE UBFile

  UBRateTblRec(1).ChkByte = ChkFlag$

  BCopy VARSEG(UBRateTblRec(1)), VARPTR(UBRateTblRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), UBRateTblRecLen, 0
  UnPackBuffer 0, 0, Form$(), Fld()

RETURN

SaveRateRec:

  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBRateTblRec(1)), VARPTR(UBRateTblRec(1)), UBRateTblRecLen, 0
  UBRateTblRec(1).ChkByte = ChkFlag$
  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  PUT UBFile, RecNo, UBRateTblRec(1)
  CLOSE UBFile
  'UBLog "ADDED RATE: " + Form$(1, 0) + " REC:" + STR$(RecNo)

RETURN

'062598 Added dupe check
Check4DupeRate:
  NewRate$ = QPTrim$(Form$(1, 0))
  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  NumOfRateRecs = LOF(UBFile) \ UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRateRecs
    GET UBFile, Cnt, UBRateTblRec(1)
    ThisRate$ = QPTrim$(UBRateTblRec(1).RATECODE)
    IF ThisRate$ = NewRate$ THEN
      UBLog "RATE DUPE: " + ThisRate$
      SaveScrn ScrnArray()
      DisplayUBScrn "ERRSCRN1"
      QPrintRC "Duplicate Rate Code", 10, 31, -1
      QPrintRC "Please Correct and Try Again.", 12, 26, -1
      WaitForAction
      RestScrn ScrnArray()
      Frm(1).FldNo = 1
      BadRate = True
      EXIT FOR
    END IF
  NEXT

  CLOSE

RETURN

END SUB

SUB DeleteRateTbl

  UBLog " IN: DELETE Rate Code"

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))      'Length of Cust Record Structure

  REDIM UBRateTblRec(1) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))

  LibName$ = "UB"
  ScrnName$ = "DELRATE"

  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  Action = 1
  FirstTime = True

  RecNo = -1             'tell SelectRate to display Delete dialog

  SelectRateTbl RecNo

  IF RecNo = 0 THEN      'if they didn't select one then exit
    GOTO ExitRateDelete
  END IF

  BlockClear
  ShowProcessingScrn " Scanning Customer Records"

  UBFile = FREEFILE      'open file and get the code for the selected rate
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  GET UBFile, RecNo, UBRateTblRec(1)
  CLOSE UBFile

  DelRate$ = QPTrim$(UBRateTblRec(1).RATECODE)
  IF LEN(DelRate$) = 0 THEN      'if it was a blank rate table, set blank flag
    BlankFlag = True             'And jump over cust search
    GOTO BlankEntry
  END IF

  UBFile = FREEFILE              'open cust file & prepare to search
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfCRecs& = LOF(UBFile) \ UBCustRecLen
  FOR Cnt& = 1 TO NumOfCRecs&          'look at all customers
    GET #UBFile, Cnt&, UBCustRec(1)
    IF NOT UBCustRec(1).DelFlag THEN   'not deleted ones
      FOR SCnt = 1 TO 15               'look at all 15 possiable revenues
        CustRate$ = QPTrim$(UBCustRec(1).Serv(SCnt).RATECODE)
        IF LEN(CustRate$) > 0 THEN     'if there is a rate code
          IF DelRate$ = CustRate$ THEN 'if they are the same then they
            NoDelFlag = True           'can't delete this rate code
            EXIT FOR                   'no need to continue looking
          END IF                       'at this customer
        END IF                         '
      NEXT                             '
      IF NoDelFlag THEN                'done with the search
        EXIT FOR
      END IF
    END IF
    ShowPctComp Cnt&, NumOfCRecs&
  NEXT
  CLOSE UBFile                         'close up

BlankEntry:
  IF NoDelFlag THEN                    'if they can't delete
    UBLog "ERROR: CAN'T DELETE RATE: " + DelRate$  'log it
    GOSUB NODeleteErr                              'show em the error
  ELSE                                 'they can delete this rate
    BlockClear
    DisplayUBScrn "DELRATE"
    DO
      EditForm Form$(), Fld(), Frm(1), Cnf, Action
      IF FirstTime THEN
        FirstTime = False
        IF BlankFlag THEN
          LSET Form$(1, 0) = "BLANK"
        ELSE
          LSET Form$(1, 0) = DelRate$
        END IF
        Action = 1
      END IF

      '--Check for Key presses
      SELECT CASE Frm(1).KeyCode       'ask em if they are sure
      CASE -68    'F10Key              'they said go ahead delete the rate
        GOSUB DeleteRateRecord         'so go do it!
        ExitFlag = True
      CASE EscKey                      'nope they changed there mind
        ExitFlag = True                'set exit flag
      END SELECT
    LOOP UNTIL ExitFlag
  END IF
                                          
ExitRateDelete:
  UBLog "OUT: DELETE Rate Code" + CRLF$           'log function exit
  ERASE UBCustRec, Frm, Form$, Fld, UBRateTblRec  'clean up

  EXIT SUB                                'exit delete function

NODeleteErr:           'Can't delete error display
  BlockClear
  HideCursor
  DisplayUBScrn "ERRSCRN1"
  QPrintRC "THERE ARE CUSTOMERS USING THAT RATE", 10, 23, -1
  QPrintRC "CAN NOT DELETE RATE CODE: " + DelRate$, 12, 26, -1
  QPrintRC DelRate$, 12, 52, 78
  ShowCursor
  WaitForAction
RETURN

DeleteRateRecord:

  KillFile "TRATETBL.TMP"              'make sure the temp files not there
  NAME "UBRATE.DAT" AS "TRATETBL.TMP"  'rename the rate file to the temp name

  UBFile1 = FREEFILE                   'open old rate file
  OPEN "TRATETBL.TMP" FOR RANDOM SHARED AS UBFile1 LEN = UBRateTblRecLen

  UBFile2 = FREEFILE                   'open new rate file
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile2 LEN = UBRateTblRecLen
  NumOfRate = LOF(UBFile1) \ UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRate                'step thru the file
    IF Cnt <> RecNo THEN                  'if this isn't rate to delete
      GET UBFile1, Cnt, UBRateTblRec(1)   'get the rec from old file
      PUT UBFile2, , UBRateTblRec(1)      'then write this rec to new file
    END IF                                '
  NEXT                                    'go till all are processed
  CLOSE                                   'close up
  KillFile "TRATETBL.TMP"                 'kill old rate file
  BlockClear                              'tell the user were done
  DisplayUBScrn "UPDATEOK"                '
  WaitForAction                           'wait for mOOse or key
  UBLog "DELETED RATE: " + DelRate$ + " REC:" + STR$(RecNo) 'log deleted rate

RETURN

END SUB

SUB EditBankDraftRec
  
  REDIM TempScrn(0)

  UBLog " IN: Edit SYS DRAFT."

  LibName$ = "UB"
  ScrnName$ = "UBDRAFT"
  NumScrns = 1
  
  '--define the multi-choice fields
  NumFlds = -1
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
    LSET Form$(F, 0) = ""
  NEXT
  
  LoadSysDraftRec Form$()
  CALL UnPackBuffer(0, 0, Form$(), Fld())
  
  '--Set screen number to one and display screen
  Scr = 1
  LibFile2Scrn LibName$, ScrnName$, MonoCode, Attribute%, ErrCode
  
  ShowCursor
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      GOSUB VerifyDraftInfo
      IF FileOK THEN
        SaveSysDraftRec Form$()
        UBLog "DFT: Updated SYS Draft INFORMATION."
        ExitFlag = True
      END IF
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 21   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 43 TO 54           '--Save Button
          PressButton -68, 21, 43, 54
        CASE 62 TO 72           '--cancel button
          PressButton 27, 21, 57, 70
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
    
  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag
  
  'HideCursor
  UBLog "OUT: Edit SYS DRAFT."
  EXIT SUB

VerifyDraftInfo:
  FileOK = True
  FileName$ = QPTrim$(Form$(8, 0))
  LSET Form$(8, 0) = FileName$
  SaveField 8, Form$(), Fld(), BadFld%

  PerPos = INSTR(FileName$, ".")

  IF PerPos > 9 OR NOT Valid(FileName$) THEN
    FileOK = False
    SaveScrn TempScrn()
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "Invalid DRAFT FILE NAME!", 10, 28, -1
    QPrintRC "Please correct and try again.", 12, 26, -1
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 8
    Action = 1
  END IF

RETURN

END SUB

SUB LoadSysDraftRec (Form$())
  
  REDIM UBSysDraftRec(1) AS UBDraftRecType
  UBSysDraftRecLen = LEN(UBSysDraftRec(1))
  Form$(0, 0) = SPACE$(UBSysDraftRecLen)
  
  UBFile = FREEFILE
  OPEN "UBSDRAFT.DAT" FOR RANDOM SHARED AS UBFile LEN = UBSysDraftRecLen
  GET UBFile, 1, UBSysDraftRec(1)
  CLOSE
  
  BCopy VARSEG(UBSysDraftRec(1)), VARPTR(UBSysDraftRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), UBSysDraftRecLen, 0
  
  ERASE UBSysDraftRec
  
END SUB

FUNCTION LOGPassWordOK

  UBLog " IN: LOG File Password"

  Pz$ = ""
  Z$ = "-:)6:(308"        'Log Password is: REVIEWLOG

  FOR Cnt = 1 TO LEN(Z$)  'decrypt the password
    Pz$ = Pz$ + CHR$(ASC(MID$(Z$, Cnt, 1)) XOR 127)
  NEXT

  DO
    BlockClear
    WazzWind 8, 15, 16, 65, 14, 1, 1

    QPrintRC "[ VIEW LOG FILE PASSWORD ]", 8, 18, 14
    QPrintRC "Enter password:", 12, 22, 15

    IF Try > 0 THEN
      QPrintRC "       Attempt:" + STR$(Try), 14, 22, 15
    END IF

    Pw$ = nGetPassword(12, 38, 10, Abort)
    
    IF Abort THEN
      UBLog "PASSWORD ABORTED: " + Pw$
      EXIT DO
    END IF

    IF Pw$ = Pz$ THEN
      UBLog "PASSWORD OK!"
      OKFlag = True
    ELSE
      UBLog "INVALID PASSWORD: " + Pw$
      QPrintRC "Invalid Password. Press any key to try again.", 14, 18, 31
      WaitForAction
    END IF

    Try = Try + 1

    IF Try = 4 THEN
      BlockClear
      WazzWind 8, 15, 16, 65, 15, 1, 1
      QPrintRC "Get the password from the system operator.", 11, 20, 14
      QPrintRC "Press any key to return to the Main Menu.", 13, 20, 14
      WaitForAction
      EXIT DO
    END IF

  LOOP UNTIL OKFlag

'VGetOut:
  IF OKFlag THEN
    LOGPassWordOK = True
  ELSE
    LOGPassWordOK = False
  END IF

END FUNCTION

SUB PrintRateListing
  
  REDIM UBRateTblRec(1) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))
  
  NumRateRecs = FileSize("UBRATE.DAT") \ UBRateTblRecLen
  
  IF NumRateRecs = 0 THEN
    GOTO ExitRateListing
  END IF
  
  BlockClear
  ShowProcessingScrn "Rate Code Listing Report."
  
  REDIM StepText(1 TO 10) AS STRING * 40
  REDIM UBSetUp(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUp(1))
  UBFile = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM SHARED AS UBFile LEN = UBSetupLen
  IF LOF(UBFile) = 0 THEN
    OKFlag = False
  ELSE
    OKFlag = True
    GET UBFile, 1, UBSetUp(1)
  END IF
  CLOSE
  
  TownName$ = UBSetUp(1).UTILNAME
  Dash80$ = STRING$(78, "-")
  NumPrinted = 0
  
  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  RPTFile = FREEFILE
  OPEN "RATELIST.RPT" FOR OUTPUT AS RPTFile
  GOSUB PrintRateHeader
  FOR Cnt = 1 TO NumRateRecs
    GET UBFile, Cnt, UBRateTblRec(1)
    IF NumPrinted = 3 THEN
      PRINT #RPTFile, Dash80$
      PRINT #RPTFile, CHR$(12)
      GOSUB PrintRateHeader
    END IF
    PRINT #RPTFile, "       Rate Code:  "; UBRateTblRec(1).RATECODE
    PRINT #RPTFile, "     Description:  "; UBRateTblRec(1).RATEDESC
    PRINT #RPTFile, "  Minimum Charge:"; FUsing$(STR$(UBRateTblRec(1).MINAMT), "#######.##")
    PRINT #RPTFile, "   Minimum Units:"; FUsing$(STR$(UBRateTblRec(1).MINUNITS), "##########")
    PRINT #RPTFile, "      Max Amount:"; FUsing$(STR$(UBRateTblRec(1).MaxAmt), "######.##")
    PRINT #RPTFile, "      [ Step ]        [ Beg Unit ]     [ Amount/Unit ]"
    FOR RCnt = 1 TO 10
      LSET StepText$(RCnt) = ""
      IF UBRateTblRec(1).TblBreaks(RCnt).UNITS >= 0 THEN
        MID$(StepText$(RCnt), 8) = FUsing$(STR$(UBRateTblRec(1).TblBreaks(RCnt).UNITS), "########")
      END IF
      IF UBRateTblRec(1).TblBreaks(RCnt).UNITAMT >= 0 THEN
        MID$(StepText$(RCnt), 25) = FUsing$(STR$(UBRateTblRec(1).TblBreaks(RCnt).UNITAMT), "####.######")
      END IF
    NEXT
    PRINT #RPTFile, "     First Break:"; StepText$(1)
    PRINT #RPTFile, "    Second Break:"; StepText$(2)
    PRINT #RPTFile, "     Third Break:"; StepText$(3)
    PRINT #RPTFile, "    Fourth Break:"; StepText$(4)
    PRINT #RPTFile, "     Fifth Break:"; StepText$(5)
    PRINT #RPTFile, "     Sixth Break:"; StepText$(6)
    PRINT #RPTFile, "   Seventh Break:"; StepText$(7)
    PRINT #RPTFile, "    Eighth Break:"; StepText$(8)
    PRINT #RPTFile, "     Ninth Break:"; StepText$(9)
    PRINT #RPTFile, "        All Over:"; StepText$(10)
    PRINT #RPTFile,
    NumPrinted = NumPrinted + 1
    ShowPctComp Cnt, NumRateRecs
  NEXT
  PRINT #RPTFile, Dash80$
  PRINT #RPTFile, CHR$(12)
  CLOSE
  
  ERASE UBSetUp, UBRateTblRec, StepText
  
  PrintRptFile "Rate Code Listing Report.", "RATELIST.RPT", 1, RetCode%, 1
  KillFile "RATELIST.RPT"
  GOTO ExitRateListing
  
PrintRateHeader:
  PageNo = PageNo + 1
  PRINT #RPTFile, "Utility Billing Rate Table Listing."
  PRINT #RPTFile, TownName$; TAB(70); "Page:"; PageNo
  PRINT #RPTFile, "Report Date: "; DATE$
  PRINT #RPTFile, Dash80$
  NumPrinted = 0
  RETURN
  
ExitRateListing:
END SUB

FUNCTION RatePassWordOK

  UBLog " IN: Rate Code Password"

  Pz$ = ""
  Z$ = "-<2>61+"    'RateCode Password is: RCMAINT
  FOR Cnt = 1 TO LEN(Z$)
    Pz$ = Pz$ + CHR$(ASC(MID$(Z$, Cnt, 1)) XOR 127)
  NEXT

  DO

    BlockClear
    WazzWind 8, 15, 16, 65, 14, 1, 1

    QPrintRC "[ RATE CODE MAINTENANCE ]", 8, 18, 14
    QPrintRC "Enter password:", 12, 22, 15

    IF Try > 0 THEN
      QPrintRC "       Attempt:" + STR$(Try), 14, 22, 15
    END IF

    Pw$ = nGetPassword(12, 38, 8, Abort)

    IF Abort THEN
      UBLog "PASSWORD ABORTED: " + Pw$
      EXIT DO
    END IF

    IF Pw$ = Pz$ THEN
      UBLog "PASSWORD OK!"
      OKFlag = True
    ELSE
      UBLog "INVALID PASSWORD: " + Pw$
      QPrintRC "Invalid Password. Press any key to try again.", 14, 18, 31
      WaitForAction
    END IF

    Try = Try + 1

    IF Try = 4 THEN
      BlockClear
      WazzWind 8, 15, 16, 65, 15, 1, 1
      QPrintRC "Get the password from the system operator.", 11, 20, 14
      QPrintRC "Press any key to return to the Main Menu.", 13, 20, 14
      WaitForAction
      EXIT DO
    END IF

  LOOP UNTIL OKFlag

GetOut:
  IF OKFlag THEN
    RatePassWordOK = True
  ELSE
    RatePassWordOK = False
  END IF

END FUNCTION

SUB RecalcAvgUse

  UBLog " IN: Recalc Average Use"
  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

'******************************************************
  LibName$ = "UBSETUP"
  ScrnName$ = "REAVGUSE"
  NumScrns = 1

  '--define the multi-choice fields
  NumFlds = -1
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
    LSET Form$(F, 0) = ""
  NEXT

  '--Set screen number to one and display screen
  Scr = 1
  BlockClear
  LibFile2Scrn LibName$, ScrnName$, MonoCode, Attribute%, ErrCode

  ShowCursor

  Action = 1

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      GOSUB CheckMonths
      IF OKFlag THEN
        EXIT DO
      END IF
    CASE EscKey
      ExitFlag = True
    END SELECT

  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag
  
  IF ExitFlag THEN
    ExitFlag = False
    UBLog "OUT: Recalc Average Use ABORTED!"
    GOTO RecalcExit
  END IF
'******************************************************

  ShowProcessingScrn "Recalculating Average Use."

  UBTran = FREEFILE
  OPEN "UBTrans.dat" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  UBFile = FREEFILE
  OPEN "UBCUST.dat" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  FOR Cnt& = 1 TO NumOfRecs&
    ShowPctComp Cnt&, NumOfRecs&
    GET UBFile, Cnt&, UBCustRec(1)
    REDIM TotalUse(1 TO 7) AS LONG
    REDIM UseCnt(1 TO 7) AS INTEGER
    DidCnt = 0
    LastTran& = UBCustRec(1).LastTrans
    DO WHILE LastTran& > 0
      GET UBTran, LastTran&, UBTranRec(1)
      IF UBTranRec(1).TransType = TranUtilityBill THEN
        FOR MCnt = 1 TO 7
          IF UBTranRec(1).CurRead(MCnt) > 0 THEN
            TotalUse(MCnt) = TotalUse(MCnt) + (UBTranRec(1).CurRead(MCnt) - UBTranRec(1).PrevRead(MCnt))
            UseCnt(MCnt) = UseCnt(MCnt) + 1
          END IF
        NEXT
        DidCnt = DidCnt + 1
        IF DidCnt >= AvgCnt THEN
          EXIT DO
        END IF
      END IF
      LastTran& = UBTranRec(1).PrevTrans
    LOOP

    FOR MCnt = 1 TO 7
      IF TotalUse(MCnt) > 0 THEN
        UBCustRec(1).LocMeters(MCnt).AvgUse = TotalUse(MCnt) / UseCnt(MCnt)
        UBCustRec(1).LocMeters(MCnt).UseCnt = UseCnt(MCnt)
      END IF
    NEXT
    PUT UBFile, Cnt&, UBCustRec(1)
  NEXT
  CLOSE
  UBLog "AVG: Recalc Average Use Complete."
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  UBLog "OUT: Recalc Average Use"
  EXIT SUB

CheckMonths:
  AvgCnt = Value(Form$(1, 0), ECode)
  IF AvgCnt < 1 OR AvgCnt > 24 THEN
    Frm(1).FldNo = 1
    Action = 2
  ELSE
    OKFlag = True
    UBLog "AVG: Recalc Using" + STR$(AvgCnt) + " Months"
  END IF
  IF OKFlag THEN
    IF NOT RecalcOK% THEN
      UBLog "AVG: Recalc Average Use ABORTED!"
      ExitFlag = True
    END IF
  END IF
  
RETURN
RecalcExit:
END SUB

FUNCTION RecalcOK%

  LibName$ = "UB"
  ScrnName$ = "MUSTEXIT"
  NumScrns = 1

  '--define the multi-choice fields
  NumFlds = -1
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
    LSET Form$(F, 0) = ""
  NEXT

  '--Set screen number to one and display screen
  Scr = 1
  BlockClear
  DisplayUBScrn ScrnName$

  FirstTime = True
  ShowCursor
  Action = 1

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      QPrintRC " READY TO RECALC AVERAGE USE?", 14, 25, 79
    END IF
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      OKFlag = True
      EXIT DO
    CASE EscKey
      ExitFlag = True
    END SELECT
  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag

  IF NOT OKFlag THEN
    RecalcOK = False
  ELSE
    RecalcOK = True
  END IF

END FUNCTION

SUB ReLinkHistory

  UBLog " IN: Relink Transaction History"
  BlockClear
  Ok = MsgBox%("UB", "MUSTEXIT")
  
  SELECT CASE Ok
  CASE 2
    UBLog "ABORTED: Relink Transaction History"
    GOTO ExitRelink
  END SELECT
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))              'Length of Cust Record Structure
  
  REDIM UBTransRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTransRec(1))             'Length of Tran Record Structure
  
  ShowProcessingScrn "Relinking Transaction History"
  
  QPrintRC "        Pass: 1 of 2         ", 11, 26, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfCRecs& = LOF(UBFile) \ UBCustRecLen
  
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  NumOfTRecs& = LOF(UBTran) \ UBTranRecLen

  UBLog "BEGIN: Pass 1 of 2"
  FOR CCnt& = 1 TO NumOfCRecs&
    GET UBFile, CCnt&, UBCustRec(1)
    UBCustRec(1).LastTrans = 0
    PUT UBFile, CCnt&, UBCustRec(1)
    ShowPctComp CCnt&, NumOfCRecs&
  NEXT

  UBLog "FINISH: Pass 1 of 2"

  QPrintRC "        Pass: 2 of 2         ", 11, 26, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1
  UBLog "BEGIN: Pass 2 of 2"
  FOR TCnt& = 1 TO NumOfTRecs&
    GET UBTran, TCnt&, UBTransRec(1)
    IF (UBTransRec(1).CustAcctNo > 0) AND (UBTransRec(1).CustAcctNo <= NumOfCRecs&) THEN
      GET UBFile, UBTransRec(1).CustAcctNo, UBCustRec(1)
      UBTransRec(1).PrevTrans = UBCustRec(1).LastTrans
      UBCustRec(1).LastTrans = TCnt&
      PUT UBFile, UBTransRec(1).CustAcctNo, UBCustRec(1)
      PUT UBTran, TCnt&, UBTransRec(1)
   ' ELSE
   '   BadCnt& = BadCnt& + 1
    END IF
    ShowPctCompL TCnt&, NumOfTRecs&
  NEXT
done:
  CLOSE
  'PRINT BadCnt&
  'WaitForAction
  UBLog "FINISH: Pass 2 of 2"
  UBLog "RELINK: Transactions Completed."
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
ExitRelink:
  UBLog "OUT: Relink Transaction History" + CRLF$

END SUB

SUB RelinkWorkOrders

  UBLog " IN: Relink Workorder History"

  BlockClear
  Ok = MsgBox%("UB", "MUSTEXIT")

  SELECT CASE Ok
  CASE 2
    UBLog "ABORTED: Relink Workorder History"
    GOTO ExitWORelink
  END SELECT

  'BlockClear
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))              'Length of Cust Record Structure

  REDIM WorkOrderRec(1) AS WorkOrderRecType
  WorkOrderRecLen = LEN(WorkOrderRec(1))

  ShowProcessingScrn "Relinking Work Order History"

  QPrintRC "        Pass: 1 of 2         ", 11, 26, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfCRecs& = LOF(UBFile) \ UBCustRecLen

  UBWrkOrd = FREEFILE
  OPEN "UBWRKORD.DAT" FOR RANDOM SHARED AS UBWrkOrd LEN = WorkOrderRecLen
  NumOfWORecs& = LOF(UBWrkOrd) \ WorkOrderRecLen
  UBLog "BEGIN: Pass 1 of 2"
  FOR CCnt& = 1 TO NumOfCRecs&
    GET UBFile, CCnt&, UBCustRec(1)
    UBCustRec(1).WOLastTrans = 0
    PUT UBFile, CCnt&, UBCustRec(1)
    ShowPctComp CCnt&, NumOfCRecs&
  NEXT
  UBLog "FINISH: Pass 1 of 2"

  QPrintRC "        Pass: 2 of 2         ", 11, 26, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1
  UBLog "BEGIN: Pass 2 of 2"
  FOR TCnt& = 1 TO NumOfWORecs&
    GET UBWrkOrd, TCnt&, WorkOrderRec(1)
    IF (WorkOrderRec(1).CustRec > 0) AND (WorkOrderRec(1).CustRec <= NumOfCRecs&) THEN
      GET UBFile, WorkOrderRec(1).CustRec, UBCustRec(1)
      WorkOrderRec(1).PrevTransRec = UBCustRec(1).WOLastTrans
      UBCustRec(1).WOLastTrans = TCnt&
      PUT UBFile, WorkOrderRec(1).CustRec, UBCustRec(1)
      PUT UBWrkOrd, TCnt&, WorkOrderRec(1)
    END IF
    ShowPctComp TCnt&, NumOfWORecs&
  NEXT
  CLOSE
  UBLog "FINISH: Pass 2 of 2"
  UBLog "RELINK: Workorders Completed."
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction

ExitWORelink:
  UBLog "OUT: Relink Workorder History" + CRLF$
  EXIT SUB

END SUB

SUB ResetProrate

  UBLog " IN: Reset Prorate Percentages"
  BlockClear
  Ok = MsgBox%("UB", "MUSTEXI2")

  SELECT CASE Ok
  CASE 2
    UBLog "ABORTED: Reset Prorate Percentages"
    GOTO ExitReset
  END SELECT

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))              'Length of Cust Record Structure

  ShowProcessingScrn "Reseting Prorate Percentages"

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfCRecs& = LOF(UBFile) \ UBCustRecLen

  UBLog "BEGIN: Reseting Percentages"
  FOR CCnt& = 1 TO NumOfCRecs&
    GET UBFile, CCnt&, UBCustRec(1)
    UBCustRec(1).ProRatePCT = 100
    PUT UBFile, CCnt&, UBCustRec(1)
    ShowPctComp CCnt&, NumOfCRecs&
  NEXT
  CLOSE

  UBLog " Reseting Percentages Complete"
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction

ExitReset:
  UBLog "OUT: Reset Prorate Percentages" + CRLF$

END SUB

SUB SaveSysDraftRec (Form$())
  
  REDIM UBSysDraftRec(1) AS UBDraftRecType
  UBSysDraftRecLen = LEN(UBSysDraftRec(1))
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBSysDraftRec(1)), VARPTR(UBSysDraftRec(1)), UBSysDraftRecLen, 0
  
  UBFile = FREEFILE
  OPEN "UBSDRAFT.DAT" FOR RANDOM SHARED AS UBFile LEN = UBSysDraftRecLen
  PUT UBFile, 1, UBSysDraftRec(1)
  CLOSE
  
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
  ERASE UBSysDraftRec
  
END SUB

SUB SelectRateTbl (RecNo)
  
  STATIC Choice
  IF Choice = 0 THEN Choice = 1

  IF RecNo = -1 THEN               'if the recno is -1 then this was
    DelFlag = True                 'called from the delete rate function
    RecNo = 0
  END IF

  REDIM UBRateTblRec(1) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))

  NumOfRateRecs = GetNumRateRecs   'get number of rates
  
  'add show a message
  IF NumOfRateRecs <= 0 THEN       'if they don't have any yet
    GOTO ExitRateSelect            'exit
  END IF
  
  REDIM MChoice$(1 TO NumOfRateRecs)  'setup array to number of rates
  UBFile = FREEFILE                   'open the rate code file
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRateRecs        'step thru all rate recs
    GET UBFile, Cnt, UBRateTblRec(1)
    MChoice$(Cnt) = SPACE$(50)        'setup space in the array element
    LSET MChoice$(Cnt) = UBRateTblRec(1).RATECODE      'load rate info
    MID$(MChoice$(Cnt), 6) = UBRateTblRec(1).RATEDESC  'into the array
    MID$(MChoice$(Cnt), 49) = MKI$(Cnt)                'set recno in array
  NEXT                                        'go till all processed
  CLOSE                                       'close rate file

  SortStr2 VARPTR(MChoice$(1)), Cnt - 1, 0    'sort the rate codes
  
  MaxLen = 47   'Set menu width to zero
  BoxBot = 16   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  'Choice = 1    'Pre-load choice to highlight
  
  TText$ = SPACE$(MaxLen + 4)
  LSET TText$ = " Code  Description"
  
  '--Center Menu within Screen

  IF DelFlag THEN
    Row = 8                                   'if called from delete
  ELSE                                        'leave room for the
    Row = 6                                   'top title box
  END IF

  Col = 15
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    ShowCursor
    IF DelFlag THEN
      TitleBox 2, Col, MaxLen + 3, "Delete Rate Table ", Cnf
    END IF
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select ", Cnf
    QPrintRC TText$, Row - 1, Col, 112
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      RecNo = 0
      ExitFlag = True
    ELSE
      RecNo = CVI(MID$(MChoice$(Choice), 49, 2))
      ExitFlag = True
    END IF
    
  LOOP UNTIL ExitFlag
  
ExitRateSelect:
  
END SUB

SUB UBRateMaint
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 4)
  
  MChoice$(1) = "Add a New Rate Table"
  MChoice$(2) = "Edit an Existing Rate Table"
  MChoice$(3) = "Delete an Existing Rate Table"
  MChoice$(4) = "Print Rate Table Listing"
  
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
  Col = ((80 - MaxLen) \ 2)
  Help$ = "Customer Maintenance Menu"
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    
    BlockClear
    TitleBox 2, Col, MaxLen + 3, "Rate Table Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      AddEditRateTbl False
    CASE 2
      AddEditRateTbl True
    CASE 3
      DeleteRateTbl
    CASE 4
      PrintRateListing
    END SELECT
  LOOP
  
END SUB

SUB ZeroOffBalances

'  UBLog " IN: Zero Offseting Balances"
'  BlockClear
'  Ok = MsgBox%("UBSETUP", "MUSTEXI3")
'
'  SELECT CASE Ok
'  CASE 2
'    UBLog "ABORTED: Zero Offseting Balances"
'    'GOTO ExitRelink
'  END SELECT
'
'  REDIM UBCustRec(1) AS NewUBCustRecType
'  UBCustRecLen = LEN(UBCustRec(1))              'Length of Cust Record Structure
'
'  ShowProcessingScrn "Zero Offseting Balances"
'
'  QPrintRC "        Pass: 1 of 2         ", 11, 26, -1
'  QPrintRC "Processing:    % Complete", 13, 28, -1
'
'  UBFile = FREEFILE
'  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
'  NumOfCRecs& = LOF(UBFile) \ UBCustRecLen
'
'  CLOSE

END SUB

