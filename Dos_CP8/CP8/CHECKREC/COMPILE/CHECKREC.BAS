DEFINT A-Z
DECLARE FUNCTION OpenLockCRFile% (CRHandle)
DECLARE SUB GetCursor (X, Y, Button)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)
DECLARE SUB AddChecks2File ()
DECLARE SUB SortCheckFile ()
DECLARE SUB SaveScrn (Array())
DECLARE SUB RestScrn (Array())
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, X%, Y%)
DECLARE SUB RemoveCanceledChks ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB CursorOff ()
DECLARE SUB PickList (Items() AS ANY, Picked%(), NPicked%, DspRow%, UpDateFlag, Cnf AS ANY)
DECLARE SUB SelectChks2Recon ()
DECLARE SUB GetBankList (BankList$(), NumBanks%, EntryFlag)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB DisplayCRScrn (ScrnName$)
DECLARE SUB CShowProcessingScrn (RptTitle$)
DECLARE SUB PrintCheckList (WhatKind%)
DECLARE FUNCTION nRound! (N!)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION MsgBox% (ScrnLib$, Scrn$)
DECLARE FUNCTION PromptSaveData% ()

DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB ClrBG ()
DECLARE FUNCTION WEnvTest ()
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE FUNCTION Monitor% ()
DECLARE SUB MPaintBox (TRow, TCol, Colr, BRow, RCol)
DECLARE SUB MnuTitle (Title$, Row, Col, Wdth, TxtColor)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
DECLARE SUB ClearEOL (Colr%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB FileView (FileName$, Ky, Action, FVI AS ANY, SEG Array)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE SUB MScrnSave (UlRow%, UlCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (UlRow%, UlCol%, LRRow%, LRCol%, SEG Element%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Date2Num% (Dat$)
  
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE SUB QPrint (X$, Colr%, Page%)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB BlockClear ()

  '$INCLUDE: 'GL.BI'
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'oschk.bi'
  '$INCLUDE: 'fileview.bi'
  '$INCLUDE: 'QSCR.bi'
  
  COMMON SHARED ToDay$, CRHandle
  
  '$INCLUDE: 'SetCnf.BI'
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  TYPE ChkSortType
    ChkNum AS SINGLE
    RecNo  AS LONG
  END TYPE
  
  CONST False = 0, True = NOT False
  
  ToDay$ = DATE$
  
  IF WEnvTest THEN
    FileExt$ = ".BAS"
  ELSE
    FileExt$ = ".EXE"
  END IF
  
  IF NOT OpenLockCRFile(CRHandle) THEN
    BlockClear
    DisplayCRScrn "LOCKERR"
    WaitForAction
    GOTO LockErrorExit
  END IF
  
  'Dim the choice array to the number of menu items
  REDIM MenuText$(1 TO 8)
  MenuText$(1) = "Print UnCanceled Checks List"
  MenuText$(2) = "Select Checks to Cancel"
  MenuText$(3) = "Print Cancelled Checks List"
  MenuText$(4) = "Remove Cancelled Checks"
  MenuText$(5) = "Add Outstanding Checks to File"
  MenuText$(6) = "Sort Outstanding Checks File"
  MenuText$(7) = "Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
  Action = 1    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  FOR Cnt = 1 TO UBOUND(MenuText$)              'Find max menu width
    TLen = LEN(MenuText$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  'Center Menu within Screen
  Row = ((25 - (UBOUND(MenuText$))) \ 2)
  Col = ((80 - MaxLen) \ 2)
  
  LOCATE 1, 1, 0
  
  'Help$ = "Budgetary Accounting"
  Help$ = "Check Reconciliation"
  Title$ = "Check Reconciliation  "
  
  ShowCursor
  
  DO
    
    ClrBG
    DisplayCRScrn "MENUBAK"
    
    MnuTitle Title$, Row, Col, MaxLen, 15
    
    QPrintRC User$, 1, 2, -1
    QPrintRC ToDay$, 1, 70, -1
    QPrintRC Help$, 25, 2, -1
    
    LOCATE Row, Col, 0          'set upper left corner of menu, turn off the cursor
    'OChoice = Choice
    'ST# = TIMER
    'ET# = ST# + 10

    DO
      CALL VertMenu(MenuText$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf)

      IF Ky$ = CHR$(27) THEN
        Choice = 7
        EXIT DO
      END IF

      IF LEN(Ky$) > 0 AND Choice > 0 AND Action <> 3 THEN
        EXIT DO          'choice = 0
      END IF

      'IF Choice <> OChoice THEN
      '  OChoice = Choice
      '  ST# = TIMER
      '  ET# = ST# + 10
      'END IF

      'IF TIMER > ET# THEN
      '  Choice = 7
      '  EXIT DO
      'END IF
    LOOP

    SELECT CASE Choice
    CASE 1
      'ClrBG
      PrintCheckList 0
    CASE 2
      'ClrBG
      SelectChks2Recon
    CASE 3
      'ClrBG
      PrintCheckList 1
    CASE 4
      RemoveCanceledChks
    CASE 5
      'ClrBG
      AddChecks2File
    CASE 6
      'ClrBG
      SortCheckFile
    CASE 7
      CLOSE CRHandle
      IF INSTR(COMMAND$, "TEST") = 0 THEN
        IF Exist("GLMENU" + FileExt$) THEN
          RUN "GLMENU"
        END IF
      END IF
      CLS
      HideCursor
      END
      
    END SELECT
    
  LOOP
  
  CLOSE CRHandle
  
LockErrorExit:
  
  ClrBG
  HideCursor
  
  IF INSTR(COMMAND$, "TEST") = 0 THEN
    IF Exist("GLMENU" + FileExt$) THEN
      RUN "GLMENU"
    END IF
  END IF
  
  CLS
  
  
  HideCursor
  END

SUB AddChecks2File
  Help$ = SPACE$(35)

  DidSome = False
  
  SHARED Choice$()
  
  REDIM BankList$(1)
  
  GetBankList BankList$(), NumBanks, 0
  
  IF NumBanks < 6 THEN
    REDIM Choice$(0 TO 6, 1)
  ELSE
    REDIM Choice$(0 TO NumBanks, 1)
  END IF
  
  Choice$(0, 0) = "6"
  Choice$(0, 1) = "1"
  
  Choice$(2, 0) = "Accounts Payable"
  Choice$(1, 0) = "Payroll"
  
  IF NumBanks < 1 THEN
    Choice$(1, 1) = " NO BANKS DEFINED!"
    PaintFlag = True
  ELSE
    FOR ZZ = 1 TO NumBanks
      Choice$(ZZ, 1) = BankList$(ZZ)
    NEXT
  END IF
  
  LibName$ = "CHECKREC"
  ScrnName$ = "ADDCHKS"
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
  
  FirstTime = True
  
  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  
  LSET Help$ = "Adding Checks"
  
  Scr = 1
  DisplayCRScrn ScrnName$
  QPrintRC Help$, 25, 2, -1
  
  IF PaintFlag THEN
    MPaintBox 13, 43, 13, 63, 0
  END IF
  ShowCursor
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      IF Frm(1).FldNo > 1 THEN
        Frm(1).FldNo = Frm(1).FldNo - 1
      END IF
      GOSUB CheckFields
      IF FieldsOK THEN
        GOSUB SaveNewCheck
        FOR F = 1 TO NumFlds
          'IF F <> 3 THEN
          LSET Form$(F, 0) = ""
          'END IF
        NEXT
        FieldsOK = False
        Frm(1).FldNo = 1
        Action = 1
      ELSE
        Action = 1
      END IF
      
    CASE 27
      AbortFlag = True
    END SELECT
    
    '--Check screen page
    
  LOOP UNTIL AbortFlag OR ExitFlag
  
  IF DidSome THEN
    SortCheckFile
  END IF
  
  GOTO ExitAdd
  
  
CheckFields:
  
  REDIM OSChk(1) AS OSChkRecType
  
  FieldsOK = True               'assume all is well
  
  BankCode$ = LEFT$(QPTrim$(Form$(1, 0)), 2)
  BankLen = LEN(BankCode$)

  OSChk(1).BankCode = Value#(BankCode$, ECode)
  IF OSChk(1).BankCode = 0 AND BankLen = 0 THEN
    FieldsOK = False
    Frm(1).FldNo = 1
    GOTO ErrorRet
  END IF
  
  OSChk(1).ChkNum = Value#(Form$(2, 0), ECode)
  IF OSChk(1).ChkNum = 0 THEN
    FieldsOK = False
    Frm(1).FldNo = 2
    GOTO ErrorRet
  END IF
  
  ChkDate$ = Form$(3, 0)
  ChkDat = Date2Num(ChkDate$)
  IF ChkDat <= 0 THEN
    FieldsOK = False
    Frm(1).FldNo = 3
    GOTO ErrorRet
  END IF
  ChkDate$ = LEFT$(ChkDate$, 2) + "/" + MID$(ChkDate$, 4, 2) + "/" + RIGHT$(ChkDate$, 2)
  OSChk(1).ChkDate = ChkDate$
  
  Desc$ = QPTrim$(Form$(4, 0))
  IF LEN(Desc$) = 0 THEN
    FieldsOK = False
    Frm(1).FldNo = 4
    GOTO ErrorRet
  END IF
  OSChk(1).Desc = Desc$
  
  ChkAmt! = Value#(Form$(5, 0), ECode)
  
  OSChk(1).Amt = ChkAmt!
  
  IF OSChk(1).Amt = 0 THEN
    FieldsOK = False
    Frm(1).FldNo = 5
    GOTO ErrorRet
  END IF
  
  Src$ = LEFT$(QPTrim$(Form$(6, 0)), 1)

  IF LEN(Src$) = 0 THEN
    FieldsOK = False
    Frm(1).FldNo = 6
    GOTO ErrorRet
  ELSE
    
    IF Src$ = "A" THEN
      OSChk(1).Src = 0
    ELSE
      OSChk(1).Src = 1
    END IF
  END IF
  
ErrorRet:
  IF NOT FieldsOK THEN
    REDIM TempScrn(1)
    SaveScrn TempScrn()
    DisplayCRScrn "BADDATA"
    WaitForAction
    RestScrn TempScrn()
  END IF
  RETURN
  
SaveNewCheck:
  
  
  ChkRecLen = LEN(OSChk(1))
  NextChkRecs = (LOF(CRHandle) \ ChkRecLen) + 1
  PUT CRHandle, NextChkRecs, OSChk(1)
  CLOSE CRHandle
  DidSome = True
  IF NOT OpenLockCRFile(CRHandle) THEN
    BlockClear
    DisplayCRScrn "LOCKERR"
    WaitForAction
    CLS
    HideCursor
    END
  END IF
  
  RETURN
  
  
ExitAdd:
  
END SUB

SUB ClrBG
  
  STATIC BeenHere, Block$
  
  IF NOT BeenHere THEN
    Block$ = STRING$(80, CHR$(177))
    BeenHere = True
  END IF
  
  LOCATE , , 0
  FOR Row = 2 TO 24
    QPrintRC Block$, Row, 1, 3
  NEXT
  QPrintRC SPACE$(80), 25, 1, 112
  QPrintRC "Wait", 25, 2, 112
  
END SUB

SUB DisplayCRScrn (ScrnName$)
  LibFile2Scrn "CHECKREC.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
END SUB

SUB MnuTitle (Title$, X, Y, Wdth, TitColor)
  
  TRow = X - 1
  W = Wdth + 4
  SCol = Y + W
  'T$ = "- " + Title$ + " -"
  
  TitleCol = (Y + ((W / 2) - (LEN(Title$) / 2)))
  IF NOT TitleCol MOD 2 THEN
    TitleCol = TitleCol + 1
  END IF
  'TitleCol = y + 1
  
  QPrintRC SPACE$(W), TRow, Y, TitColor
  QPrintRC Title$, TRow, TitleCol, TitColor
  MPaintBox X, SCol, X, SCol + 1, 8
  
  
END SUB

FUNCTION nRound! (N!)
  N! = INT(N! * 100 + .5) / 100
  nRound! = N!
END FUNCTION

SUB PrintCheckList (WhatKind%)

  SHARED Choice$()

  REDIM GLSetup(1) AS GLSetupRecType
  GLFile = FREEFILE
  OPEN "GLSETUP.DAT" FOR RANDOM AS GLFile LEN = LEN(GLSetup(1))
  GET GLFile, , GLSetup(1)
  CLOSE GLFile

  TownName$ = QPTrim$(GLSetup(1).UserName)

  Page = 0
  MaxLines = 55
  Dash$ = STRING$(80, "-")
  FF$ = CHR$(12)
  
  REDIM BankList$(1)
  
  GetBankList BankList$(), NumBanks, -1
  
  ReportFile$ = "RPT.PRN"
  IF WhatKind = 0 THEN
    RptTitle$ = "Outstanding Checks Report"
  ELSE
    RptTitle$ = " Canceled Checks Report  "
    UseChk = False
    UseBank = False

'This was changed to show date in canceled list
    UseDate = True
    ThruDate = 32767   'set to max possiable date
    GOTO ProcessReport
  END IF
  
  IF NumBanks < 3 THEN
    REDIM Choice$(0 TO 3, 1)
  ELSE
    REDIM Choice$(0 TO NumBanks, 1)
  END IF
  
  Choice$(0, 0) = "2"
  Choice$(1, 0) = "Both"
  Choice$(2, 0) = "Payroll"
  Choice$(3, 0) = "Accounts Payable"

  Choice$(0, 1) = "3"
  IF NumBanks = 1 THEN
    Choice$(1, 1) = " NO BANKS DEFINED!"
    PaintFlag = True
  ELSE
    FOR ZZ = 1 TO NumBanks
      Choice$(ZZ, 1) = BankList$(ZZ)
    NEXT
  END IF
  
  LibName$ = "CHECKREC"
  ScrnName$ = "CRPTINFO"
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
  
  FirstTime = True
  
  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  
  Scr = 1
  DisplayCRScrn ScrnName$
  IF PaintFlag THEN
    MPaintBox 13, 43, 13, 63, 0
  END IF
  ShowCursor
  
  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      LSET Form$(1, 0) = DATE$
      LSET Form$(2, 0) = Choice$(1, 0)
      LSET Form$(3, 0) = BankList$(1)
      FirstTime = False
      QPrintRC RptTitle$, 6, 23, -1
      Action = 2
    END IF

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key

      ThruDate = Date2Num(Form$(1, 0))
      IF ThruDate > 0 THEN
        UseDate = True
      END IF

      SELECT CASE LEFT$(Form$(2, 0), 1)
      CASE "B"
        UseChk = False
      CASE "P"
        ChkType = 1
        UseChk = True
      CASE "A"
        ChkType = 0
        UseChk = True
      END SELECT
      Bank$ = QPTrim$(Form$(3, 0))
      IF LEFT$(Bank$, 1) = "A" THEN
        UseBank = False
      ELSE
        BankNum = VAL(Bank$)
        UseBank = True
      END IF
      ExitFlag = True
    CASE 27
      AbortFlag = True
    END SELECT
    
  LOOP UNTIL AbortFlag OR ExitFlag
  
  IF AbortFlag THEN GOTO ChkListExit
  
ProcessReport:
  
  HideCursor
  
  'Report is sent to the following file which is passed to fileview for
  'screen output or printed using the BLPrint routine
  
  TotalChks# = 0                'Problem in totals using single precision
  NumChks = 0
  
  Colr = 14     'Set Data entry Color

  REDIM OSChk(1) AS OSChkRecType
  
  RecLen = LEN(OSChk(1))
  
  NumRecs = LOF(CRHandle) \ RecLen
  PrnFileNum = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #PrnFileNum
  
  ShowProcessingScrn RptTitle$
  GOSUB PrintChkLstHeader
  
  FOR Cnt = 1 TO NumRecs
    GET CRHandle, Cnt, OSChk(1)
    IF OSChk(1).Cleared = WhatKind THEN
      IF UseDate THEN
        GOSUB GetDateInfo

        IF CheckDate > ThruDate THEN
          GOTO SkipThisChk
        END IF
      END IF

      IF UseBank THEN
        IF BankNum <> OSChk(1).BankCode THEN
          GOTO SkipThisChk
        END IF
      END IF

      IF UseChk THEN
        IF ChkType <> OSChk(1).Src THEN
          GOTO SkipThisChk
        END IF
      END IF

'if we get to here this chk is ok
      NumChks = NumChks + 1
      ChkAmt# = nRound!(OSChk(1).Amt)
      TotalChks# = TotalChks# + ChkAmt#
      SELECT CASE OSChk(1).Src
      CASE 0
        T$ = "AP"
      CASE 1
        T$ = "PR"
      END SELECT
      ToPrint$ = SPACE$(80)
      LSET ToPrint$ = FUsing$(STR$(OSChk(1).ChkNum), "#######")
      'MID$(ToPrint$, 9) = OSChk(1).ChkDate
      MID$(ToPrint$, 10) = OSChkDate$
      MID$(ToPrint$, 22) = T$
      MID$(ToPrint$, 26) = OSChk(1).Desc
      MID$(ToPrint$, 60) = FUsing$(STR$(OSChk(1).Amt), "########,.##")
      MID$(ToPrint$, 78) = STR$(OSChk(1).BankCode)
      PRINT #PrnFileNum, ToPrint$
      LineCnt = LineCnt + 1
    END IF
SkipThisChk:
    ShowPctComp Cnt, NumRecs
    IF LineCnt > MaxLines THEN
      PRINT #PrnFileNum, FF$
      GOSUB PrintChkLstHeader
    END IF
  NEXT
  
  PRINT #PrnFileNum, ""         'Add a blank line after last line
  
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = FUsing$(STR$(NumChks), "######")
  MID$(ToPrint$, 8) = "Checks listed totaling: "
  MID$(ToPrint$, 60) = FUsing$(STR$(TotalChks#), "########,.##")
  
  PRINT #PrnFileNum, ToPrint$ + CrLf$
  CLOSE '#PrnFileNum
  
  PrintRptFile RptTitle$, ReportFile$, 1, RetCode, 0

  IF NOT OpenLockCRFile(CRHandle) THEN
    Oh.Shit! = True
  END IF

  'KILL ReportFile$     'Clean up after ourselves
  'CLOSE
  
ChkListExit:
  
  EXIT SUB

Getout:
  'UNLOCK FileNum
  'CLOSE
  EXIT SUB
  
PrintChkLstHeader:
  Page = Page + 1
  PRINT #PrnFileNum, TAB(29); RptTitle$; TAB(70); "Page:"; Page
  PRINT #PrnFileNum,
  PRINT #PrnFileNum, TownName$; TAB(54); "Report Date: "; DATE$
  PRINT #PrnFileNum, " Chk No   Chk Date   Src              Payee                     Chk Amt     Bank"
  PRINT #PrnFileNum, Dash$
  LineCnt = 5
RETURN

GetDateInfo:
  OSChkDate$ = QPTrim$(OSChk(1).ChkDate)
  ChkYear$ = RIGHT$(OSChkDate$, 2)
  IF ChkYear$ = "19" THEN ChkYear$ = "99"
  ChkYear = VAL(ChkYear$)
  'IF ChkYear = 20 THEN STOP
  SELECT CASE ChkYear
  CASE 20
    ChkYear$ = ChkYear$ + "00"
  CASE 79 TO 99
    ChkYear$ = "19" + ChkYear$
  CASE ELSE
    ChkYear$ = "20" + ChkYear$
  END SELECT
  OSChkDate$ = LEFT$(OSChkDate$, 6) + ChkYear$
  CheckDate = Date2Num(OSChkDate$)
RETURN

END SUB

SUB RemoveCanceledChks
  
  Escape$ = CHR$(27)
  
  Ky$ = INKEY$     'ButNo,DnFlg,NoPresses,col,row
  ButtonPress 1, N, MooseButton, MCol, MRow     ' ----- Check Mouse activity
  IF N AND 2 THEN               'if it was the right moose button and
    DO          'if they are still holding it down then
      GetCursor 0, 0, NewStatus 'wait till they let go of the button
    LOOP WHILE NewStatus
    ButtonPress 0, 0, 0, 0, 0   'this has the effect of clearing all
    Ky$ = Escape$               'moose buttons.
  END IF
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  ActMRow = 16
  BlockClear
  
  DisplayCRScrn "PURGCHK2"
  
  DO
    
    Get.Moose.OR.Key Ky$, MooseButton, MRow, MCol
    
    IF MooseButton THEN
      MRow = (MRow \ 8) + 1     'Convert MCol, MRow to Row and Col.
      MCol = (MCol \ 8) + 1
      IF MRow = ActMRow THEN
        SELECT CASE MCol
        CASE 30 TO 41           'Cancel printing
          PressButton EscKey, ActMRow, 30, 41
        CASE 42 TO 52           'let it rip
          PressButton -68, ActMRow, 42, 52
        END SELECT
      END IF
    END IF
    
    IF LEN(Ky$) THEN
      IF Ky$ = Escape$ THEN
        PurgeEM = False
        EXIT DO
      ELSEIF LEN(Ky$) = 2 THEN
        Ky$ = RIGHT$(Ky$, 1)
        IF Ky$ = "D" THEN
          PurgeEM = True
          EXIT DO
        END IF
      END IF
    END IF
  LOOP
  
  RestScrn TempScrn()
  ERASE TempScrn
  
  IF NOT PurgeEM THEN
    GOTO AbortPurge
  END IF
  
  ShowProcessingScrn "Purging Canceled Checks"
  
  REDIM OSChk(1) AS OSChkRecType
  ChkRecLen = LEN(OSChk(1))
  
  NumChkRecs = LOF(CRHandle) \ ChkRecLen
  
  IF NumChkRecs = 0 THEN
    DisplayCRScrn "MENUBAK"
    OK = MsgBox%("CHECKREC.QSL", "NOCTRANS")
    GOTO AbortPurge
  END IF
  
  NChkFile = FREEFILE
  OPEN "crchk1.dat" FOR OUTPUT AS #NChkFile
  CLOSE NChkFile
  OPEN "crchk1.dat" FOR RANDOM SHARED AS NChkFile LEN = ChkRecLen
  
  FOR CCnt = 1 TO NumChkRecs
    GET CRHandle, CCnt, OSChk(1)
    IF OSChk(1).Cleared = 0 THEN
      PUT NChkFile, , OSChk(1)
    END IF
    ShowPctComp CCnt, NumChkRecs
  NEXT
  CLOSE CRHandle
  CLOSE NChkFile

  KILL "crchk.dat"
  NAME "crchk1.dat" AS "crchk.dat"
  
  'must close rename and then reopen the checks file! should
  'force windows to commit all to disk
  IF NOT OpenLockCRFile(CRHandle) THEN
    BlockClear
    DisplayCRScrn "LOCKERR"
    WaitForAction
    CLS
    HideCursor
    END
  END IF

  DisplayCRScrn "MENUBAK"
  DisplayCRScrn "UPDATEOK"
  WaitForAction
  
AbortPurge:
  
END SUB

SUB SelectChks2Recon
  
  MaxLoad& = (FRE(-1) - 20000) / 66
  IF MaxLoad& > 4999 THEN
    MaxLoad& = 4999
  END IF
  
  ShowProcessingScrn "Reading Check Information File"
  
  REDIM OSChk(1) AS OSChkRecType
  ChkRecLen = LEN(OSChk(1))
  
  NumChkRecs = LOF(CRHandle) \ ChkRecLen
  
  IF NumChkRecs = 0 THEN
    DisplayCRScrn "MENUBAK"
    OK = MsgBox%("CHECKREC.QSL", "NOCTRANS")
    GOTO AbortExit
  END IF
  
  FOR CCnt = 1 TO NumChkRecs
    GET CRHandle, CCnt, OSChk(1)
    DCnt = DCnt + 1
    Show$ = "Checks Found:" + FUsing(STR$(DCnt), "######")
    QPrintRC Show$, 11, 26, 112
    
    REDIM PRESERVE MTChoice(1 TO DCnt) AS FLen2
    REDIM PRESERVE Picked(1 TO DCnt)

    SELECT CASE OSChk(1).Src
    CASE 0
      T$ = "A"
    CASE 1
      T$ = "P"
    END SELECT

    OSChkDate$ = QPTrim$(OSChk(1).ChkDate)
    ChkYear$ = RIGHT$(OSChkDate$, 2)
    IF ChkYear$ = "19" THEN ChkYear$ = "99"
    ChkYear = VAL(ChkYear$)
    SELECT CASE ChkYear
    CASE 20
      ChkYear$ = ChkYear$ + "00"
    CASE 79 TO 99
      ChkYear$ = "19" + ChkYear$
    CASE ELSE
      ChkYear$ = "20" + ChkYear$
    END SELECT
    OSChkDate$ = LEFT$(OSChkDate$, 6) + ChkYear$

    LSET MTChoice(DCnt).V = FUsing(STR$(OSChk(1).ChkNum), "########")
    MID$(MTChoice(DCnt).V, 10) = OSChkDate$
    MID$(MTChoice(DCnt).V, 21) = LEFT$(OSChk(1).Desc, 28)
    MID$(MTChoice(DCnt).V, 51) = T$
    MID$(MTChoice(DCnt).V, 52) = FUsing(STR$(OSChk(1).Amt), "#######.##")
    
    IF OSChk(1).Cleared > 0 THEN
      PCnt = PCnt + 1
      Picked(PCnt) = DCnt
    END IF
    
    ShowPctComp CCnt, NumChkRecs
    
    IF DCnt > MaxLoad& THEN EXIT FOR
    
  NEXT
  
  BlockClear
  
HistTop:
  Help$ = SPACE$(80)
  
  MaxLen = 59   'Set menu width to zero
  Action = 0    '0 means stay in the menu until they select something
  
  IF Choice < 1 THEN
    Choice = 1  'Pre-load choice to highlight
  END IF
  
  Title$ = SPACE$(68)
  LSET Title$ = "   Chk No.  Chk Date      Description              Src   Amount TAG"
  LSET Help$ = "Budgetary Accounting"
  
  Col = 10: Row = 3
  
  DisplayCRScrn "PickOpti"
  QPrintRC Title$, Row - 1, Col - 4, 112

  PickList MTChoice(), Picked(), DCnt, Row, UpDateFlag, Cnf
  
  IF UpDateFlag THEN
    ShowProcessingScrn "Processing Selections"
    FOR CCnt = 1 TO NumChkRecs
      GET CRHandle, CCnt, OSChk(1)
      OSChk(1).Cleared = 0
      PUT CRHandle, CCnt, OSChk(1)
      ShowPctComp CCnt, NumChkRecs
    NEXT
    
    FOR PCnt = 1 TO DCnt
      GET CRHandle, Picked(PCnt), OSChk(1)
      OSChk(1).Cleared = 1
      PUT CRHandle, Picked(PCnt), OSChk(1)
      ShowPctComp PCnt, DCnt
    NEXT
    
    'close and reopen the checks file! should
    'force windows to commit all to disk
    CLOSE CRHandle
    IF NOT OpenLockCRFile(CRHandle) THEN
      BlockClear
      DisplayCRScrn "LOCKERR"
      WaitForAction
      CLS
      HideCursor
      END
    END IF
  END IF
  
AbortExit:
  
END SUB

SUB SortCheckFile
  
  REDIM OSChk(1) AS OSChkRecType
  ChkRecLen = LEN(OSChk(1))
  ChkFile = FREEFILE
  OPEN "crchk1.dat" FOR OUTPUT AS ChkFile
  CLOSE ChkFile
  ChkFile = FREEFILE
  OPEN "crchk1.dat" FOR RANDOM SHARED AS ChkFile LEN = ChkRecLen
  
  NumChkRecs = LOF(CRHandle) \ ChkRecLen
  
  IF NumChkRecs <= 1 THEN
    'CLOSE
    DisplayCRScrn "MENUBAK"
    OK = MsgBox%("CHECKREC.QSL", "NOCTRANS")
    GOTO AbortSort
  END IF
  
  REDIM Index(1 TO NumChkRecs) AS ChkSortType
  
  ShowProcessingScrn "Reading Check Information File"
  
  FOR CCnt = 1 TO NumChkRecs
    GET CRHandle, CCnt, OSChk(1)
    Index(CCnt).ChkNum = OSChk(1).ChkNum
    Index(CCnt).RecNo = CCnt
    ShowPctComp CCnt, NumChkRecs
  NEXT
  
  QPrintRC "           Sorting            ", 9, 26, -1
  
  SortT Index(1), NumChkRecs, 0, 8, 0, -3
  
  QPrintRC "Writing Check Information File", 9, 26, -1
  
  FOR CCnt = 1 TO NumChkRecs
    GET CRHandle, Index(CCnt).RecNo, OSChk(1)
    PUT ChkFile, CCnt, OSChk(1)
    ShowPctComp CCnt, NumChkRecs
  NEXT
  
  CLOSE CRHandle
  CLOSE ChkFile
  
  KILL "crchk.dat"
  NAME "crchk1.dat" AS "crchk.dat"
  
'must close rename and then reopen the checks file! should
'force windows to commit all to disk
  IF NOT OpenLockCRFile(CRHandle) THEN
    BlockClear
    DisplayCRScrn "LOCKERR"
    WaitForAction
    CLS
    HideCursor
    END
  END IF

  DisplayCRScrn "MENUBAK"
  DisplayCRScrn "UPDATEOK"
  WaitForAction

  'CLOSE
  
AbortSort:
  
END SUB

