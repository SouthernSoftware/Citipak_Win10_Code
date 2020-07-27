DEFINT A-Z
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB FGetA (Handle%, SEG Dest AS ANY, NumBytes%)
DECLARE FUNCTION Round# (DblNum#)
'DECLARE SUB SmallPause ()
DECLARE SUB SelectNoteText (NoteText$)
DECLARE FUNCTION CustHasMsg% (RecNo&)
DECLARE SUB PrintTransHist (DetailFlag)
DECLARE SUB PrintConsumpHist ()
DECLARE SUB EditOwnerInfo (RecNo&, OTempName$)
DECLARE SUB ReIndexSystem (PromptFlag%)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Chk4DupeBookSeqNum (Book$, SeqNum$)
DECLARE FUNCTION Compare3% (BYVAL Seg1, BYVAL Adr1, BYVAL Seg2, BYVAL Adr2, NumBytes)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FmtBook$ (Book$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (Number$)

DECLARE SUB AddEditCustomer (RecNo&)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB CustMessageSystem (RecNo&)
DECLARE SUB DeleteCust ()
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
DECLARE SUB PrintCustByRate ()
DECLARE SUB PrintCustList (LocFlag%)
DECLARE SUB PrintLocaList (LocFlag)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveNewCustRec (Form$(), RecNo&)
DECLARE SUB SaveOldCustRec (Form$(), RecNo&)
DECLARE SUB SaveScrn (Array%())
DECLARE SUB SetCustFinal ()
DECLARE SUB ShowCursor ()
DECLARE SUB ShowCustConsHist (CustRec&)
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowWrkOrdHistory (CustRec&)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB UBLog (Text$)
DECLARE SUB WaitForAction ()
DECLARE SUB MPaintBox (TRow%, LCol%, BRow%, RCol%, PColor%)
DECLARE FUNCTION ChkBillFile% ()
DECLARE FUNCTION GetCustMeterType% (UBCustRec() AS ANY, ThisMeter%)
DECLARE SUB MakeSequenceIndex (Text$)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'PageInfo.BI'                  'Form Page info
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'ubTrans.BI'
  '$INCLUDE: 'ubowner.BI'
  '$INCLUDE: 'ubhusky.BI'

  CONST False = 0, True = NOT False
  
  CrLf$ = CHR$(13) + CHR$(10)

SUB AddEditCustomer (RecNo&)

  REDIM UBSetUp(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetUpRecLen
  IF INSTR(UBSetUp(1).utilname, "LUGOFF") > 0 THEN
    LugFlag = True
  END IF
  ERASE UBSetUp

  REDIM UBOwnerRec(1) AS UBOwnerRecType
  UBOwnerRecLen = LEN(UBOwnerRec(1))
  OFirstTime = True

  REDIM NoteFlds(1 TO 3) AS INTEGER

  CrLf$ = CHR$(13) + CHR$(10)

  IF ChkBillFile% THEN
    'STOP
  END IF

  UBLog " IN: Add/Edit Customer"
  IF RecNo& > 0 THEN
    UBLog "EDIT:" + STR$(RecNo&)
  ELSE
    UBLog "ADD NEW"
  END IF

  REDIM ScrnArray(0)
  
  REDIM UBRateRecs(1) AS UBRateTblRecType
  UBRateRecLen = LEN(UBRateRecs(1))
  NumOfRateRecs = FileSize("UBRATE.DAT") \ UBRateRecLen
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  FOR RevCnt = 1 TO MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME)
    IF LEN(TempRev$) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    END IF
  NEXT

  IF NumOfRevs = 0 THEN
    NumOfRevs = 15
  END IF

'******************************************************
  LibName$ = "UB"
  ScrnName$ = "UBCUST"
  NumScrns = 4
  
  '--define the multi-choice fields
  SHARED Choice$()
  REDIM Choice$(0 TO 4, 0)
  
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

'*********Setup the owner control arrays
  OwnNumFlds = LibNumberOfFields("UB", "UBOWNER")
  REDIM OwnFrm(1) AS FormInfo
  REDIM OwnForm$(OwnNumFlds, 2)
  REDIM OwnFld(OwnNumFlds) AS FieldInfo

  OwnStartEL = 0
  LibGetFldDef "UB", "UBOWNER", OwnStartEL, OwnFld(), OwnForm$(), ErrCode
  OwnForm$(0, 0) = SPACE$(UBOwnerRecLen)
'***************

  FOR Cnt = 1 TO 3
    NoteFlds(Cnt) = FldNum%("HHMSG" + LTRIM$(STR$(Cnt)), Fld())
  NEXT

  GOSUB InitMultiChoice
  
  IF NumOfRevs < MaxRevsCnt THEN
    FOR ZZCnt = NumOfRevs + 1 TO MaxRevsCnt
      Fld(FldNum%("RC" + LTRIM$(STR$(ZZCnt)), Fld())).Protected = True
      Fld(FldNum%("MTRTYP" + LTRIM$(STR$(ZZCnt)), Fld())).Protected = True
    NEXT
  END IF
  
  FOR ZZCnt = 1 TO NumOfRevs
    UseAMeter$ = QPTrim$(UBSetUpRec(1).Revenues(ZZCnt).UseMtr)
    IF UseAMeter$ = "" OR UseAMeter$ = "N" THEN
      Fld(FldNum%("MTRTYP" + LTRIM$(STR$(ZZCnt)), Fld())).Protected = True
    END IF
  NEXT
  
  FOR ZZCnt = 1 TO NumOfRevs
    UseARate$ = QPTrim$(UBSetUpRec(1).Revenues(ZZCnt).USERATE)
    IF UseARate$ = "" OR UseARate$ = "N" THEN
      Fld(FldNum%("RC" + LTRIM$(STR$(ZZCnt)), Fld())).Protected = True
    END IF
  NEXT
  
'--Clear all fields
  FOR F = 1 TO NumFlds
    TFormLen = TFormLen + Fld(F).StorLen
    'this calcs the actual len of the forms part of the record,
    'it allows us to have additional fields in our type structure
    'and still put the correct amount in the form itself.
    '    LSET Form$(F, 0) = ""
  NEXT
  
'--Must fill Form$(0, 0) to the correct len, before
  'using the BCopy function.
  Form$(0, 0) = SPACE$(TFormLen)

  IF RecNo& > 0 THEN
    GOSUB LoadCustInfo
    UnPackBuffer 0, 0, Form$(), Fld()
  ELSE
    Frm(1).PageDisable = True
    GOSUB InitCustDefaults
  END IF
  
'--Set screen number to one and display screen
  Scr = 1
  Action = 1
  FirstTime = True
  
  DisplayUBScrn FormName$(Scr)
  
  ShowCursor
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      GOSUB UpDateDisplay
      SaveScrn ScrnArray()
    END IF
    
    IF Frm(1).FldNo <> Frm(1).PrevFld THEN
      GOSUB UpDateDisplay
      IF Frm(1).Edited THEN
        EditedFlag = True
      END IF
    END IF

    IF NOT FinalFlag THEN
      IF Frm(1).FldNo = 3 AND Frm(1).PrevFld < 3 THEN            '>2
        Form$(1, 0) = FmtBook$(Form$(1, 0))
        Form$(2, 0) = FmtSeqN$(Form$(2, 0))
        NBook$ = QPTrim$(Form$(1, 0)) + "-" + QPTrim$(Form$(2, 0))
        IF (OldBook$ <> NBook$) AND (NBook$ <> "00-000000") THEN
          'if they changed or edited the book-seq list
          IF Chk4DupeBookSeqNum(Form$(1, 0), Form$(2, 0)) THEN
            EditedFlag = False
            Frm(1).KeyCode = 0
            IF LEN(OldBook$) > 1 THEN
              LSET Form$(1, 0) = LEFT$(OldBook$, 2)
              LSET Form$(2, 0) = MID$(OldBook$, 4)
              Frm(1).FldNo = 1
            ELSE
              LSET Form$(1, 0) = ""
              LSET Form$(2, 0) = ""
              Frm(1).FldNo = 1
            END IF
          ELSE
            Frm(1).PageDisable = False
          END IF
          Action = 2
        END IF
      END IF
    END IF
'--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key 'F10
      SaveScrn ScrnArray()
      IF RecNo& = 0 THEN
        DisplayUBScrn "UPDATDSK"
        SaveNewCustRec Form$(), RecNo&
        'RecNo& will have a value after the above save rec sub
        GOSUB SaveOwnerInfo
        LogMsg$ = "ADDED NEW:"
      ELSEIF RecNo& > 0 THEN
        SaveOldCustRec Form$(), RecNo&
        LogMsg$ = "UPDATED:"
      END IF
      UBLog LogMsg$ + STR$(RecNo&) + " " + TempName$
      ExitFlag = True

    CASE F2Key
      IF Frm(1).FldNo >= NoteFlds(1) AND Frm(1).FldNo <= NoteFlds(3) THEN
        SelectNoteText NoteText$
        IF LEN(NoteText$) > 0 THEN
          LSET Form$(Frm(1).FldNo, 0) = NoteText$
          Action = 2
        END IF
      END IF
    CASE F3Key
      IF RecNo& > 0 THEN
        ShowWrkOrdHistory RecNo&
        Action = 2
      END IF

    CASE F4KEY
      IF RecNo& > 0 THEN
        ShowCustHistory RecNo&
        Action = 2
      END IF
    
    CASE F6KEY
      IF RecNo& > 0 THEN
        ShowCustConsHist RecNo&
        Action = 2
      END IF

    CASE F7KEY
      IF RecNo& > 0 THEN
        LastFld = Frm(1).FldNo
        SaveScrn ScrnArray()
        CustMessageSystem RecNo&
        GOSUB CheckCustMsg
        RestScrn ScrnArray()
        Frm(1).FldNo = LastFld
        Action = 1
        FirstTime = True
      END IF

    CASE F8KEY
      WhatFld = Frm(1).FldNo
      GOSUB EditOwnerInfo
      Frm(1).FldNo = WhatFld
      Action = 2
      
    CASE EscKey
      IF EditedFlag THEN
        SaveFlag = PromptSaveData
        SELECT CASE SaveFlag
        CASE True               'user wants to save
          StuffBuf CHR$(0) + CHR$(ABS(F10Key))
        CASE False              'user wants to abandon
          IF RecNo& = 0 THEN
            UBLog "ABORTED: ADD NEW"
          ELSE
            UBLog "ABORTED: EDIT:" + STR$(RecNo&) + " " + TempName$
          END IF
          ExitFlag = True
        CASE ELSE               'neither continue editing
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
        CASE 5 TO 12            '--F2 HH Code list
          PressButton F2Key, 22, 5, 12
        CASE 13 TO 21           '--F3 Wrk Hist
          PressButton F3Key, 22, 13, 21
        CASE 22 TO 30           '--F4 Tran Hist
          PressButton F4KEY, 22, 22, 30
        CASE 31 TO 39           '--F6 Cons Hist
          PressButton F6KEY, 22, 31, 39
        CASE 40 TO 47           '--F7 Msgs
          IF RecNo& > 0 THEN
            PressButton F7KEY, 22, 40, 47
          END IF
        CASE 48 TO 55           '--F8 Owner
          PressButton F8KEY, 22, 48, 55
        CASE 56 TO 65           '--F10 Save Button
          PressButton F10Key, 22, 56, 65
        CASE 66 TO 75           '--ESC cancel button
          PressButton 27, 22, 66, 75
        END SELECT
      END SELECT                'row
    END IF
    
'--Check screen page
    IF Frm(1).FldNo > PageInfo(Scr).LastFld OR Frm(1).FldNo < PageInfo(Scr).FirstFld THEN
      SELECT CASE Frm(1).FldNo
      CASE PageInfo(1).FirstFld TO PageInfo(1).LastFld
        Scr = 1
      CASE PageInfo(2).FirstFld TO PageInfo(2).LastFld
        Scr = 2
      CASE PageInfo(3).FirstFld TO PageInfo(3).LastFld
        Scr = 3
      CASE PageInfo(4).FirstFld TO PageInfo(4).LastFld
        Scr = 4
      END SELECT
      DisplayUBScrn FormName$(Scr)
      GOSUB UpDateDisplay
      IF Scr = 3 THEN
        GOSUB DisplayRevSources
      END IF
    END IF
  LOOP UNTIL ExitFlag
  
ExitCustEdit:
  UBLog "OUT: Add/Edit Customer" + CrLf$

EXIT SUB

EditOwnerInfo:

  OwnFrm(1).FldNo = 1
  OExitFlag = False

  REDIM TempScr(0)
  SaveScrn TempScr()

  IF RecNo& > 0 THEN
    GOSUB LoadOwnerInfo
  ELSE
    BCopy SSEG(OwnForm$(0, 0)), SADD(OwnForm$(0, 0)), VARSEG(UBOwnerRec(1)), VARPTR(UBOwnerRec(1)), UBOwnerRecLen, 0
  END IF

  BCopy VARSEG(UBOwnerRec(1)), VARPTR(UBOwnerRec(1)), SSEG(OwnForm$(0, 0)), SADD(OwnForm$(0, 0)), UBOwnerRecLen, 0
  UnPackBuffer 0, 0, OwnForm$(), OwnFld()

  OAction = 1
  
  MPaintBox 2, 5, 22, 75, 8
  DisplayUBScrn "UBOWNER"

  DO

    EditForm OwnForm$(), OwnFld(), OwnFrm(1), Cnf, OAction

    '--Check for Key presses
    SELECT CASE OwnFrm(1).KeyCode
    CASE F10Key 'F10
      IF RecNo& > 0 THEN
        GOSUB SaveOwnerInfo
      END IF
      OExitFlag = True
    CASE EscKey
      BCopy VARSEG(UBOwnerRec(1)), VARPTR(UBOwnerRec(1)), SSEG(OwnForm$(0, 0)), SADD(OwnForm$(0, 0)), UBOwnerRecLen, 0
      RestScrn TempScr()
      OExitFlag = True
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

  LOOP UNTIL OExitFlag

  RestScrn TempScr()
  ERASE TempScr

RETURN

LoadOwnerInfo:
  UBFile = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBFile LEN = UBOwnerRecLen
  GET UBFile, RecNo&, UBOwnerRec(1)
  CLOSE
RETURN

SaveOwnerInfo:
  BCopy SSEG(OwnForm$(0, 0)), SADD(OwnForm$(0, 0)), VARSEG(UBOwnerRec(1)), VARPTR(UBOwnerRec(1)), UBOwnerRecLen, 0
  UBOwnerRec(1).ChkByte = CHR$(1)
  UBFile = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBFile LEN = UBOwnerRecLen
  PUT UBFile, RecNo&, UBOwnerRec(1)
CLOSE

RETURN

DisplayRevSources:
  FOR RevCnt = 1 TO 8
    QPrintRC UBSetUpRec(1).Revenues(RevCnt).REVNAME, RevCnt + 4, 9, -1
  NEXT
  FOR RevCnt = 9 TO 15
    QPrintRC UBSetUpRec(1).Revenues(RevCnt).REVNAME, RevCnt - 4, 44, -1
  NEXT
RETURN
  
InitMultiChoice:
  MinArraySize = 8
  IF NumOfRateRecs > MinArraySize THEN
    MinArraySize = NumOfRateRecs
  END IF
  REDIM Choice$(0 TO MinArraySize, 0 TO 4)
  
  StatusFld = FldNum%("ASTATUS", Fld())
  Choice$(0, 3) = STR$(StatusFld)
  
  FOR Cnt = 1 TO MaxRevsCnt
    Choice$(0, 0) = Choice$(0, 0) + STR$(FldNum%("RC" + LTRIM$(STR$(Cnt)), Fld())) + ","
  NEXT
  FOR Cnt = 1 TO MaxRevsCnt
    Choice$(0, 1) = Choice$(0, 1) + STR$(FldNum%("MTRTYP" + LTRIM$(STR$(Cnt)), Fld())) + ","
  NEXT
  FOR Cnt = 1 TO 7
    Choice$(0, 1) = Choice$(0, 1) + STR$(FldNum%("METER" + LTRIM$(STR$(Cnt)), Fld())) + ","
    Choice$(0, 4) = Choice$(0, 4) + STR$(FldNum%("MTRUNT" + LTRIM$(STR$(Cnt)), Fld())) + ","
  NEXT
  FOR Cnt = 1 TO 4
    Choice$(0, 2) = Choice$(0, 2) + STR$(FldNum%("FRFREQ" + LTRIM$(STR$(Cnt)), Fld())) + ","
  NEXT
  
  FOpenS "UBRATE.DAT", Handle   'open data file
  FOR Cnt = 1 TO NumOfRateRecs
    FGetRTA Handle, UBRateRecs(1), CLNG(Cnt), UBRateRecLen
    Choice$(Cnt, 0) = UBRateRecs(1).RATECODE + "  " + UBRateRecs(1).RATEDESC
  NEXT
  FClose Handle
  
  Choice$(1, 3) = "Active"
  Choice$(2, 3) = "Inactive"
  Choice$(3, 3) = "Balance"
  Choice$(4, 3) = "Pending"
  Choice$(5, 3) = "Delinquent"
  
  Choice$(1, 1) = "Water Only"
  Choice$(2, 1) = "Sewer Only"
  Choice$(3, 1) = "Combined Water & Sewer"
  Choice$(4, 1) = "Electric Meter"
  Choice$(5, 1) = "Demand Electric"
  Choice$(6, 1) = "Gas Meter"
  Choice$(7, 1) = "Touch Read Meter"
  IF LugFlag THEN
    Choice$(8, 1) = "P-Schulmberger"
  END IF
  
  Choice$(1, 4) = "Gallons"
  Choice$(2, 4) = "Cubic Feet"
  Choice$(3, 4) = "Kilowatts"
  
  Choice$(1, 2) = "Recurring"
  Choice$(2, 2) = "NoRecurring"
RETURN
  
InitCustDefaults:
  Form$(3, 0) = "A"
  Form$(4, 0) = DATE$
  Form$(10, 0) = UBSetUpRec(1).DEFCITY
  Form$(11, 0) = UBSetUpRec(1).DEFSTATE
  Form$(12, 0) = UBSetUpRec(1).ZipCode

  Form$(19, 0) = "C"
  Form$(20, 0) = "1"
  
  Form$(25, 0) = "N"
  Form$(26, 0) = "N"
  Form$(27, 0) = "Y"
  Form$(28, 0) = "Y"
  Form$(29, 0) = "N"
  Form$(30, 0) = "N"
  Form$(31, 0) = "N"
  Form$(32, 0) = "0"
  Form$(44, 0) = "100"

  FOR ZZCnt = 100 TO 109
    Form$(ZZCnt, 0) = "0"
  NEXT

  FOR ZZCnt = 110 TO 188 STEP 13
    Form$(ZZCnt + 1, 0) = "1"
    Form$(ZZCnt + 4, 0) = "1"
  NEXT

  FOR ZZCnt = 79 TO 94 STEP 5
    Form$(ZZCnt + 1, 0) = "0"
    Form$(ZZCnt + 3, 0) = "0"
    Form$(ZZCnt + 4, 0) = "1"
  NEXT
RETURN
  
UpDateDisplay:
  AcctNo$ = QPTrim$(STR$(RecNo&))
  IF AcctNo$ = "0" THEN
    AcctNo$ = "??"
  END IF
  TempName$ = QPTrim$(Form$(6, 0))
  IF (LEN(TempName$) = 0) AND RecNo& = 0 THEN
    TempName$ = " NEW "
  ELSEIF (LEN(TempName$) = 0) AND RecNo& > 0 THEN
    TempName$ = " ??? "
  ELSE
    TempName$ = " " + LEFT$(TempName$, 30) + " "
  END IF
  QPrintRC TempName$, 2, 40 - (LEN(TempName$) / 2), 112

  IF GotMsgFlag THEN
    QPrintRC "F7", 22, 41, 254
  ELSE
    QPrintRC "F7", 22, 41, 126
  END IF

  IF Scr = 2 THEN
    IF Frm(1).FldNo >= NoteFlds(1) AND Frm(1).FldNo <= NoteFlds(3) THEN
      QPrintRC "F2", 22, 6, 126
    ELSE
      QPrintRC "F2", 22, 6, 120
    END IF
  END IF
  
  IF Scr = 1 THEN
    QPrintRC AcctNo$, 4, 66, -1
  END IF
RETURN
  
LoadCustInfo:
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  Form$(0, 0) = SPACE$(TFormLen)
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  GET UBFile, RecNo&, UBCustRec(1)
  CLOSE
  
  OldBook$ = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb
  
  BCopy VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), TFormLen, 0
  
  IF UBCustRec(1).Status = "F" THEN
    FinalFlag = True
  END IF
  
  DepAmt# = UBCustRec(1).DepositAmt

  ERASE UBCustRec

  GOSUB CheckCustMsg
RETURN

CheckCustMsg:
  IF CustHasMsg(RecNo&) THEN
    GotMsgFlag = True
  ELSE
    GotMsgFlag = False
  END IF
RETURN
  
END SUB

SUB SaveNewCustRec (Form$(), RecNo&)

'NOTE: The owner info is saved in the main customer editing procedure.

  REDIM UBCustRec(1) AS NewUBCustRecType

  FormLen = LEN(Form$(0, 0))           'Length of Cust Form$(0,0) Element
  
  UBCustRecLen = LEN(UBCustRec(1))     'Length of Cust Record Structure
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1)), FormLen, 0
  
  REDIM UBCustPIN(1) AS UBPINType      'Pin info array
  
  IF FileSize("UBCUSPIN.DAT") > 0 THEN 'If there was a previous pin then
    PHandle = FREEFILE
    OPEN "UBCUSPIN.DAT" FOR RANDOM SHARED AS #PHandle LEN = 4
    GET #PHandle, 1, UBCustPIN(1)      'get last pin used info
    CLOSE #PHandle
  END IF
  
  NextPIN& = UBCustPIN(1).PIN + 1      'Increment last pin used
  
  UBCustPIN(1).PIN = NextPIN&
  PHandle = FREEFILE                   'Save new last pin info
  OPEN "UBCUSPIN.DAT" FOR RANDOM SHARED AS #PHandle LEN = 4
  PUT #PHandle, 1, UBCustPIN(1)
  CLOSE #PHandle
  
  UBCustRec(1).CustPIN = NextPIN&
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  NextRec& = NumOfRecs& + 1
  
  UBCustRec(1).MessageRec = 0          'Set Pointers To Zero
  UBCustRec(1).LastTrans = 0
  UBCustRec(1).CurrBalance = 0
  UBCustRec(1).PrevBalance = 0
  'UBCustRec(1).DepositAmt = 0
  UBCustRec(1).WOLastTrans = 0
  
  PUT UBFile, NextRec&, UBCustRec(1)
  CLOSE UBFile
  
  ReIndexSystem False
  
  RecNo& = NextRec&
  ERASE UBCustRec
  
END SUB

SUB SaveOldCustRec (Form$(), RecNo&)
  
  CursorOff
  REDIM TempScrn(0)

  SaveScrn TempScrn()

  DisplayUBScrn "UPDATDSK"

  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType

  UBCustRecLen = LEN(UBCustRec(1))

  UBFile = FREEFILE
  
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen

  GET UBFile, RecNo&, UBCustRec(1)
'Copy orignal record to temp copy. This holds info not attached directly to
'the customer editing form.
  LSET UBCustRec(2) = UBCustRec(1)

'Copy edited form info to the customers record.
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1)), LEN(Form$(0, 0)), 0

'Restore meter data not in editing form
'060597 Solves date & avguse errors
  FOR MtrCnt = 1 TO 7
    UBCustRec(1).LocMeters(MtrCnt).PastDate = UBCustRec(2).LocMeters(MtrCnt).PastDate
    UBCustRec(1).LocMeters(MtrCnt).ReadFlag = UBCustRec(2).LocMeters(MtrCnt).ReadFlag
    UBCustRec(1).LocMeters(MtrCnt).AvgUse = UBCustRec(2).LocMeters(MtrCnt).AvgUse
    UBCustRec(1).LocMeters(MtrCnt).UseCnt = UBCustRec(2).LocMeters(MtrCnt).UseCnt
  NEXT
  'UBCustRec(1).DelFlag = 0
  PUT UBFile, RecNo&, UBCustRec(1)
  
  CLOSE UBFile
  
'01-20-97 Added skip reindexing if name not changed
  IF UBCustRec(1).SEARCH <> UBCustRec(2).SEARCH OR UBCustRec(1).CustName <> UBCustRec(2).CustName THEN
    ReindexFlag = True
  END IF
  
'01-20-97 Added skip reindexing if book or seq number not changed
  IF (UBCustRec(1).Book <> UBCustRec(2).Book) OR (UBCustRec(1).SeqNumb <> UBCustRec(2).SeqNumb) THEN
    ReindexFlag = True
  END IF
  
  IF ReindexFlag THEN
    KillFile "UBOOKSEQ.DAT"
    ReIndexSystem False
  END IF

  RestScrn TempScrn()

  ERASE UBCustRec, TempScrn

END SUB

SUB SelectNoteText (NoteText$)

  STATIC Choice
  IF Choice = 0 THEN Choice = 1

  REDIM CodeRec(1) AS MtrNotesType
  CodeLen = LEN(CodeRec(1))
  NumCodeRecs = FileSize("UBHHCODE.DAT") \ CodeLen

  IF NumCodeRecs = 0 THEN
    GOTO ExitCodeSelect
  END IF

  REDIM MChoice$(1 TO NumCodeRecs)  'setup array to number of rates
  UBFile = FREEFILE                   'open the rate code file
  OPEN "UBHHCODE.DAT" FOR RANDOM SHARED AS UBFile LEN = CodeLen
  FOR Cnt = 1 TO NumCodeRecs        'step thru all rate recs
    GET UBFile, Cnt, CodeRec(1)
    MChoice$(Cnt) = SPACE$(20)        'setup space in the array element
    LSET MChoice$(Cnt) = CodeRec(1).Note      'load rate info
  NEXT                                        'go till all processed
  CLOSE                                       'close rate file

  'SortStr2 VARPTR(MChoice$(1)), Cnt - 1, 0    'sort the rate codes

  MaxLen = 20   'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  
  '--Center Menu within Screen

  Row = 13                                   'top title box
  Col = 39

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      NoteText$ = ""
      RecNo = 0
      ExitFlag = True
    ELSE
      NoteText$ = MChoice$(Choice)
      ExitFlag = True
    END IF

  LOOP UNTIL ExitFlag

ExitCodeSelect:

END SUB

SUB UBCustomerMenu
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 11)
  
  MChoice$(1) = " Add a New Customer"
  MChoice$(2) = " Edit Existing Customer"
  MChoice$(3) = " Set a Customer to Final"
  MChoice$(4) = " Delete Customer Account"
  MChoice$(5) = " Quick Customer Report Menu"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2) + 2
  Col = ((80 - MaxLen) \ 2)
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    
    TitleBox 2, Col, MaxLen + 3, "Utility Customer Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      'REDIM CRec(1) AS NewUBCustRecType
      'IF FileSize("UBCUST.DAT") \ LEN(CRec(1)) < 3 THEN
      AddEditCustomer 0
      'ELSE
      '  CursorOff
      '  BlockClear
      '  DisplayUBScrn "DEMOONLY"    'show demo version screen
      '  WaitForAction
      'END IF
      
    CASE 2
      DO
        BlockClear
        LookUp RecNo&, "Customer", 2, True, False
        IF RecNo& > 0 THEN
          AddEditCustomer RecNo&
        END IF
      LOOP WHILE RecNo& > 0
    CASE 3
      SetCustFinal
    CASE 4
      DeleteCust
    CASE 5
      RUN "UBQCURPT"
    END SELECT
  LOOP
  
END SUB

