DEFINT A-Z
DECLARE FUNCTION PromptSaveData% ()
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB DisplaySetupScrn (ScrnName$)
DECLARE SUB RelinkWorkOrders ()
DECLARE FUNCTION ubAcctLookUp% (Acct$)
DECLARE SUB ubGetAcctStruct (GLFundLen%, GLAcctLen%, GLDetLen%)
DECLARE SUB UBLog (Text$)
DECLARE SUB OpenAcctIdx (AcctIdxFileNum%, NumIdxRecs%)
DECLARE SUB OpenAcctFile (AcctFileNum%, NumAccts%)
DECLARE SUB ReLinkHistory ()
DECLARE SUB ShowPctComp2 (BYVAL RecNo&, BYVAL NumOfRecs&)
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
DECLARE SUB AddEditRateTbl (EdFlag%)
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB LoadRateTblRec (Form$(), TblNum%)
DECLARE SUB SaveRateTblRec (Form$(), RecNo%)
DECLARE SUB SelectRate2Edit (RecNo%)
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
  
  '$INCLUDE: 'DefCnf.BI'
DECLARE SUB AcctPickList (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS Config)
  
DECLARE FUNCTION ASCII (Character$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Number$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION QPStrL$ (Number&)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)
DECLARE FUNCTION GetNumRateRecs ()
DECLARE SUB ClearScrn ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE FUNCTION ubFmtAcct$ (AcctNum$, FundLen%, AcctLen%, DetLen%)
DECLARE FUNCTION ubStripAcct$ (AcctNum$)
DECLARE FUNCTION ubFindAcct% (AcctNum$)
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'ubtrans.bi'
  '$INCLUDE: 'PageInfo.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.BI'
  
  '$INCLUDE: 'GL.BI'
  ''$INCLUDE: 'AP.BI'
  
  CONST False = 0, True = NOT False
  
  CRLF$ = CHR$(13) + CHR$(10)
  
  '  EditUBSetUpRec
  
  LibName$ = "UBSETUP"
  ScrnName$ = "UBSETUP"
  NumScrns = 5
  
  REDIM ScrnArray(0)
  
  '--define the multi-choice fields
  'SHARED Choice$()
  
  REDIM Choice$(0 TO 13, 1)
  Choice$(0, 0) = "19"
  Choice$(1, 0) = "HUSKY"
  Choice$(2, 0) = "PC3000"
  Choice$(3, 0) = "MPC2000"
  Choice$(4, 0) = "Sensus"
  Choice$(5, 0) = "ESensus"
  Choice$(6, 0) = "C-SysCon"
  Choice$(7, 0) = "Data 'Genie'"
  Choice$(8, 0) = "Telxon"
  Choice$(9, 0) = "Logicon"
  Choice$(10, 0) = "Itron"
  Choice$(11, 0) = "Badger Radio"
  Choice$(12, 0) = "Z-Schulmberger"
  Choice$(13, 0) = "None"
  
  Choice$(0, 1) = "12"
  Choice$(1, 1) = "1. Account Number"
  Choice$(2, 1) = "2. Search Name"
  Choice$(3, 1) = "3. Meter Number"
  Choice$(4, 1) = "4. Service Address"
  Choice$(5, 1) = "5. Location Number"
  Choice$(6, 1) = "6. 911 Address"
  
  '--Initialize the form name array
  REDIM FormName$(1 TO NumScrns)
  FOR Scr = 1 TO NumScrns
    FormName$(Scr) = ScrnName$ + LTRIM$(STR$(Scr))
  NEXT
  
  '--Get the total number of fields from all pages
  NumFlds = -1
  FOR Scr = 1 TO NumScrns
    NumFlds = NumFlds + LibNumberOfFields(LibName$, FormName$(Scr)) + 1
  NEXT
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  REDIM PageInfo(1 TO NumScrns) AS PageInfoType
  REDIM DistFldNums(1 TO 15) AS INTEGER
  
  '--for each screen, get first and last fields
  StartEl = 0
  FOR Scr = 1 TO NumScrns
    LibGetFldDef LibName$, FormName$(Scr), StartEl, Fld(), Form$(), ErrCode
    PageInfo(Scr).PageNo = Scr
    PageInfo(Scr).FirstFld = StartEl + 1
    '--Increment StartEl to next
    StartEl = StartEl + Fld(StartEl).Fields + 1
    PageInfo(Scr).LastFld = StartEl - 1
  NEXT
  
  REDIM UBSetUp(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUp(1))
  
  PROPFldNo = FldNum%("PRODIST", Fld())
  MethAFld = FldNum%("METHACCT", Fld())
  BilCycFld = FldNum%("BILLCYCL", Fld())
  PreByBok = FldNum%("PREBOOK", Fld())
  
  FOR Cnt = 1 TO 15
    DistFldNums(Cnt) = FldNum%("DISTOR" + QPTrim$(STR$(Cnt)), Fld())
  NEXT
  
  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  
  LoadSetUpRec Form$(), OKFlag                 'load the system setup file
  
  IF OKFlag THEN                               'if it loaded ok then
    CALL UnPackBuffer(0, 0, Form$(), Fld())    'unpack it to the form
  ELSE
    SetUpDefaultsFlag = True
  END IF
  GOSUB FixDistFields
  
  Frm(1).StayOnField = True         'step of the current field only if
                                    'they press enter or up/down arrow
  '--Set screen number to one and display screen
  Scr = 1
  
  DisplaySetupScrn FormName$(Scr)
  
  ShowCursor
  
  Action = 1
  FirstTime = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF Frm(1).Edited THEN
      BeenEdited = True
    END IF
    
    IF Frm(1).FldNo <> Frm(1).PrevFld THEN  'can't prebill by books & cycle!
      SELECT CASE Form$(PreByBok, 0)
      CASE "Y"
        LSET Form$(BilCycFld, 0) = "N"
        Action = 1
      END SELECT
    END IF
    
    IF FirstTime THEN               'if this is the first time in formedit
      FirstTime = False             'then if this is also the first time
      IF SetUpDefaultsFlag THEN     'in the setup system function then
        SetUpDefaultsFlag = False   'clear set default flag
        GOSUB SetUpDefaults         'go setup some basic system options
        Action = 1                  'reset action flag to redisplay form info
      END IF
      GOSUB CheckMethAcctFld    'check there accounting method
      IF ChkGLAcctFlag THEN     'if it's set then go load account info
        GOSUB LoadGLAcctInfo
      END IF
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key    'F10Key
      GOSUB ChkDistFlds         'go check the assigned accounts
      IF SetUpOK THEN           'if there ok then
        SaveSetUpRec Form$()    'save the system setup record
        ExitFlag = True         'set exit flag
      ELSE
      END IF
      
    CASE EscKey
      IF BeenEdited THEN            'ask user if they want to save
        SaveFlag = PromptSaveData   'the setup file
        SELECT CASE SaveFlag
        CASE True                   'user wants to save
          StuffBuf CHR$(0) + CHR$(ABS(F10Key))
          Frm(1).KeyCode = 0
        CASE False                  'user wants to abandon
          ExitFlag = True
        CASE ELSE               'continue editing
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
        CASE 43 TO 54           '--Save Button
          PressButton -68, 22, 43, 54
        CASE 62 TO 72           '--cancel button
          PressButton 27, 22, 57, 70
        END SELECT
      END SELECT                'row
    END IF
    
    'UNREM THIS SECTION WHEN FINISHED WITH THE INTERFACE TO GL
    'if they just stepped of the Accounting Method field
    IF Frm(1).PrevFld = MethAFld AND Frm(1).FldNo <> MethAFld THEN
      GOSUB CheckMethAcctFld    'then do account check
    END IF
    
    'if they on page 3 or higher & chkacctflag then
    IF Frm(1).FldNo > PageInfo(3).FirstFld + 1 AND ChkGLAcctFlag = True THEN
      IF Frm(1).FldNo <> Frm(1).PrevFld AND Fld(Frm(1).PrevFld).Protected = 0 THEN
        IF Frm(1).FldNo < NumFlds THEN
          TempAcct$ = ubStripAcct$(Form$(Frm(1).PrevFld, 0))
          'strip out any dashes in the account number
          IF LEN(TempAcct$) > 0 THEN
            'if there was an account
            TempAcct$ = ubFmtAcct$(TempAcct$, GLFundLen, GLAcctLen, GLDetLen)
            'reformat to make sure it's ok for ubAcctLookUp
            AcctOK = ubAcctLookUp(TempAcct$)    'go look for it
            IF AcctOK > 0 THEN  'if it was a valid account
              LSET Form$(Frm(1).PrevFld, 0) = TempAcct$
              'set it back into the form formated correctly
              Action = 1
            ELSE                'nope not a valid account
              Frm(1).FldNo = Frm(1).PrevFld
              LSET Form$(Frm(1).PrevFld, 0) = ""
              Action = 1
            END IF
          END IF
        END IF
      END IF
    END IF
    
    ''END OF REM SECTION
    
    '--Check screen page
    IF Frm(1).FldNo > PageInfo(Scr).LastFld OR Frm(1).FldNo < PageInfo(Scr).FirstFld THEN
      SELECT CASE Frm(1).FldNo  'if this field is not on the current page
      CASE PageInfo(1).FirstFld TO PageInfo(1).LastFld
        Scr = 1
      CASE PageInfo(2).FirstFld TO PageInfo(2).LastFld
        Scr = 2
      CASE PageInfo(3).FirstFld TO PageInfo(3).LastFld
        Scr = 3
      CASE PageInfo(4).FirstFld TO PageInfo(4).LastFld
        Scr = 4
      CASE PageInfo(5).FirstFld TO PageInfo(5).LastFld
        Scr = 5
      END SELECT
      Action = 1
      DisplaySetupScrn FormName$(Scr)
      IF Scr = 2 AND DidWarning = False THEN
        GOSUB ShowWarning
      END IF
      IF Scr = 3 THEN           'if this is page 3 or higher then
        GOSUB DisplayRevs       'display the revenue sources on
      END IF                    'the left side of the page
      GOSUB FixDistFields
    END IF
    
  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag
  
  ERASE Frm, Form$, Fld, PageInfo, DistFldNums, UBSetUp
  
  IF INSTR(COMMAND$, "TEST") THEN
    HideCursor
    END
  ELSE
    RUN "ubmisc"
  END IF
  
CheckMethAcctFld:
  SELECT CASE Form$(MethAFld, 0)                'if the field is "C"
  CASE "C", "A"
    ChkGLAcctFlag = True
  CASE ELSE
    ChkGLAcctFlag = False
  END SELECT
  
  RETURN
  
LoadGLAcctInfo:
  ubGetAcctStruct GLFundLen, GLAcctLen, GLDetLen
  
RETURN
  
ChkDistFlds:
  REDIM PropDist(1 TO 10) AS INTEGER
  BadDistFlag = False
  REDIM PropDist(1 TO 10) AS INTEGER
  FOR Cnt = 1 TO 10
    PropDist(Cnt) = QPValI(Form$(DistFldNums(Cnt), 0))
  NEXT
  SortT PropDist(1), 10, 0, 2, 0, -1
  '(Element(1), NumElm, Dir, SSize%, MOffSet, MSize%)
  
  FOR Cnt = 1 TO 9
    IF PropDist(Cnt) > 0 AND PropDist(Cnt + 1) > 0 THEN
      IF PropDist(Cnt) = PropDist(Cnt + 1) THEN
        BadDistFlag = True
        EXIT FOR
      END IF
    END IF
  NEXT
  
  FOR Cnt = 1 TO 10
    IF PropDist(Cnt) > 10 THEN
      BadDistFlag = True
      EXIT FOR
    END IF
  NEXT
  
  IF BadDistFlag THEN
    SaveScrn ScrnArray()
    DisplayUBScrn "DUPEDIST"
    WaitForAction
    RestScrn ScrnArray()
  ELSE
    SetUpOK = True
  END IF
RETURN
  
FixDistFields:
  DistMax = 10
  FOR Cnt = 1 TO 10
    Fld(DistFldNums(Cnt)).HiRange = DistMax
  NEXT
RETURN
  
SetUpDefaults:
  
  LSET Form$(5, 0) = "N"
  LSET Form$(6, 0) = "1"
  LSET Form$(7, 0) = "Y"
  FOR Cnt = 8 TO 11
    LSET Form$(Cnt, 0) = "N"
  NEXT
  LSET Form$(12, 0) = "1"
  LSET Form$(13, 0) = " "
  FOR Cnt = 14 TO 16
    LSET Form$(Cnt, 0) = "N"
  NEXT
  LSET Form$(17, 0) = "75"
  LSET Form$(18, 0) = "150"
  LSET Form$(19, 0) = "N"

'This does all of page 2 defaults
  FOR Cnt = 1 TO 15
    SCnt$ = QPTrim$(STR$(Cnt))
    FldOSet = FldNum%("DEP" + SCnt$, Fld())
    LSET Form$(FldOSet, 0) = "N"
    LSET Form$(FldOSet + 1, 0) = "N"
    LSET Form$(FldOSet + 2, 0) = "0"
    LSET Form$(FldOSet + 3, 0) = "N"
    LSET Form$(FldOSet + 4, 0) = SCnt$
    LSET Form$(FldOSet + 5, 0) = "N"
  NEXT

RETURN
  
DisplayRevs:
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBSetUp(1)), VARPTR(UBSetUp(1)), UBSetupLen, 0
  FOR TRevCnt = 1 TO 15
    UBSetUp(1).BillAcct(TRevCnt).REVNAME = UBSetUp(1).Revenues(TRevCnt).REVNAME
    UBSetUp(1).PayAcct(TRevCnt).REVNAME = UBSetUp(1).BillAcct(TRevCnt).REVNAME
    UBSetUp(1).DepAcct(TRevCnt).REVNAME = UBSetUp(1).BillAcct(TRevCnt).REVNAME
  NEXT
  BCopy VARSEG(UBSetUp(1)), VARPTR(UBSetUp(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), UBSetupLen, 0
  CALL UnPackBuffer(0, 0, Form$(), Fld())
RETURN

ShowWarning:
 SaveScrn ScrnArray()
 DidWarning = True
 CursorOff
 DisplaySetupScrn "REVWARN"
 WaitForAction
 RestScrn ScrnArray()
RETURN

SUB DisplaySetupScrn (ScrnName$)
  LibFile2Scrn "UBSETUP", ScrnName$, MonoCode%, Attribute%, ErrorCode%
END SUB

SUB LoadSetUpRec (Form$(), OKFlag)
  
  REDIM UBSetUp(1) AS UBSetupRecType
  
  UBSetupLen = LEN(UBSetUp(1))
  
  Form$(0, 0) = SPACE$(UBSetupLen)
  
  UBFile = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM SHARED AS UBFile LEN = UBSetupLen
  IF LOF(UBFile) = 0 THEN
    OKFlag = False
  ELSE
    OKFlag = True
    GET UBFile, 1, UBSetUp(1)
  END IF
  CLOSE
  IF OKFlag THEN
    BCopy VARSEG(UBSetUp(1)), VARPTR(UBSetUp(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), UBSetupLen, 0
  END IF
  
  ERASE UBSetUp
  
END SUB

SUB SaveSetUpRec (Form$())
  
  REDIM UBSetUp(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUp(1))
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBSetUp(1)), VARPTR(UBSetUp(1)), UBSetupLen, 0
  
  UBFile = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM SHARED AS UBFile LEN = UBSetupLen
  PUT UBFile, 1, UBSetUp(1)
  CLOSE
  
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
  ERASE UBSetUp
  
END SUB

