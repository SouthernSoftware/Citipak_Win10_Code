DEFINT A-Z
DECLARE SUB MakeEPPReconTrans (RecNo&, ClearEPPFlag%)
DECLARE SUB ReconcileCustomer ()
DECLARE SUB PrintTransHist ()
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
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()

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
  '$INCLUDE: 'ubEPP.BI'

  CONST False = 0, True = NOT False
  
  CRLF$ = CHR$(13) + CHR$(10)

SUB AddEditCustomer (RecNo&)

  REDIM UBOwnerRec(1) AS UBOwnerRecType
  UBOwnerRecLen = LEN(UBOwnerRec(1))

  REDIM UBEPPRec(1) AS UBCustEPPRecType
  UBEPPRecLen = LEN(UBEPPRec(1))
  
  CRLF$ = CHR$(13) + CHR$(10)

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
  OwnNumFlds = LibNumberOfFields(LibName$, "UBOWNER")
  REDIM OwnFrm(1) AS FormInfo
  REDIM OwnForm$(OwnNumFlds, 2)
  REDIM OwnFld(OwnNumFlds) AS FieldInfo
  OwnStartEL = 0
  LibGetFldDef LibName$, "UBOWNER", OwnStartEL, OwnFld(), OwnForm$(), ErrCode
  OwnForm$(0, 0) = SPACE$(UBOwnerRecLen)
'***************

'*********Setup Equal Payment control arrays
  EPPNumFlds = LibNumberOfFields(LibName$, "UBCUSTEP")
  REDIM EPPFrm(1) AS FormInfo
  REDIM EPPForm$(EPPNumFlds, 2)
  REDIM EPPFld(EPPNumFlds) AS FieldInfo
  EPPStartEL = 0
  LibGetFldDef LibName$, "UBCUSTEP", EPPStartEL, EPPFld(), EPPForm$(), ErrCode
  EPPForm$(0, 0) = SPACE$(UBEPPRecLen)
'***************

  GOSUB InitMultiChoice
  
  IF NumOfRevs < MaxRevsCnt THEN
    FOR ZZCnt = NumOfRevs + 1 TO MaxRevsCnt
      Fld(FldNum%("RC" + LTRIM$(STR$(ZZCnt)), Fld())).Protected = True
      Fld(FldNum%("MTRTYP" + LTRIM$(STR$(ZZCnt)), Fld())).Protected = True
       EPPFld(FldNum%("EQDIS" + LTRIM$(STR$(ZZCnt)), EPPFld())).Protected = True
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
        RestScrn ScrnArray()
        Frm(1).FldNo = LastFld
        Action = 1
      END IF

    CASE F8KEY
      WhatFld = Frm(1).FldNo
      GOSUB EditOwnerInfo
      Frm(1).FldNo = WhatFld
      Action = 2

    CASE F9KEY
      IF RecNo& > 0 THEN
        WhatFld = Frm(1).FldNo
        GOSUB EditEPPInfo
        Frm(1).FldNo = WhatFld
        Action = 2
      END IF
    
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
        CASE 5 TO 13            '--F3 Wrk Hist
          PressButton F3Key, 22, 5, 13
        CASE 14 TO 22           '--F4 Tran Hist
          PressButton F4KEY, 22, 14, 22
        CASE 23 TO 31           '--F6 Cons Hist
          PressButton F6KEY, 22, 23, 31
        CASE 32 TO 39           '--F7 Msgs
          IF RecNo& > 0 THEN
            PressButton F7KEY, 22, 32, 39
          END IF
        CASE 40 TO 47           '--F8
          PressButton F8KEY, 22, 40, 47
        CASE 48 TO 55           '--F9 Equal Payment
          PressButton F9KEY, 22, 48, 55
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
  UBLog "OUT: Add/Edit Customer" + CRLF$

EXIT SUB


EditEPPInfo:

  EFirstTime = True

  EPPFrm(1).FldNo = 1
  EExitFlag = False

  REDIM TempScr(0)
  SaveScrn TempScr()

  IF RecNo& > 0 THEN
    GOSUB LoadEPPInfo
  ELSE
    BCopy SSEG(EPPForm$(0, 0)), SADD(EPPForm$(0, 0)), VARSEG(UBEPPRec(1)), VARPTR(UBEPPRec(1)), UBEPPRecLen, 0
  END IF

  BCopy VARSEG(UBEPPRec(1)), VARPTR(UBEPPRec(1)), SSEG(EPPForm$(0, 0)), SADD(EPPForm$(0, 0)), UBEPPRecLen, 0
  UnPackBuffer 0, 0, EPPForm$(), EPPFld()

  EAction = 1

  MPaintBox 2, 5, 22, 75, 8
  DisplayUBScrn "UBCUSTEP"

  DO
    EditForm EPPForm$(), EPPFld(), EPPFrm(1), Cnf, EAction
    IF EFirstTime THEN
      EFirstTime = False
      FOR RevCnt = 1 TO NumOfRevs
        QPrintRC LEFT$(UBSetUpRec(1).Revenues(RevCnt).REVNAME, 13), RevCnt + 3, 22, -1
      NEXT
      FOR Cnt = 1 TO EPPNumFlds
        EPPFld(Cnt).Protected = 0
      NEXT
    END IF

    '--Check for Key presses
    SELECT CASE EPPFrm(1).KeyCode
    CASE F10Key 'F10
      IF RecNo& > 0 THEN
        GOSUB SaveEPPInfo
      END IF
      EExitFlag = True
    CASE EscKey
      BCopy VARSEG(UBEPPRec(1)), VARPTR(UBEPPRec(1)), SSEG(EPPForm$(0, 0)), SADD(EPPForm$(0, 0)), UBEPPRecLen, 0
      RestScrn TempScr()
      EExitFlag = True
    END SELECT

    '--check for mouse clicks on buttons not attached to the form
    IF EPPFrm(1).Presses THEN
      SELECT CASE EPPFrm(1).MRow
      CASE 21   'Look for the f10 or esc button
        SELECT CASE EPPFrm(1).MCol
        CASE 37 TO 46           '--F10 Save Button
          PressButton F10Key, 21, 37, 46
        CASE 47 TO 57           '--ESC Exit
          PressButton EscKey, 21, 47, 57
        END SELECT
      END SELECT                'row
    END IF

  LOOP UNTIL EExitFlag

  RestScrn TempScr()
  ERASE TempScr

RETURN

LoadEPPInfo:
  UBFile = FREEFILE
  OPEN "UBCUSTEP.DAT" FOR RANDOM SHARED AS UBFile LEN = UBEPPRecLen
  
  GET UBFile, RecNo&, UBEPPRec(1)
  CLOSE
RETURN

SaveEPPInfo:

  BCopy SSEG(EPPForm$(0, 0)), SADD(EPPForm$(0, 0)), VARSEG(UBEPPRec(1)), VARPTR(UBEPPRec(1)), UBEPPRecLen, 0
  UBEPPRec(1).ChkByte = CHR$(1)

'make sure none are blank!
  FOR zz = 1 TO 15
    IF UBEPPRec(1).EPInfo(zz).RevDis < -100000 THEN
      UBEPPRec(1).EPInfo(zz).RevDis = 0
    END IF
  NEXT

  UBFile = FREEFILE
  OPEN "UBCUSTEP.DAT" FOR RANDOM SHARED AS UBFile LEN = UBEPPRecLen
  PUT UBFile, RecNo&, UBEPPRec(1)
  CLOSE
  DisplayUBScrn "UPDATEOK"
  WaitForAction
RETURN


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
        CASE 42 TO 53           '--ESC Button
          PressButton EscKey, 19, 42, 53
        END SELECT
      END SELECT                'row
    END IF

  '--Check screen page
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
  
  Choice$(1, 1) = "Water Only"
  Choice$(2, 1) = "Sewer Only"
  Choice$(3, 1) = "Combined Water & Sewer"
  Choice$(4, 1) = "Electric Meter"
  Choice$(5, 1) = "Demand Electric"
  Choice$(6, 1) = "Gas Meter"
  Choice$(7, 1) = "Touch Read Meter"
  
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
  Form$(12, 0) = UBSetUpRec(1).ZIPCODE

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
  Form$(43, 0) = "100"

  FOR ZZCnt = 98 TO 108
    Form$(ZZCnt, 0) = "0"
  NEXT

  FOR ZZCnt = 109 TO 187 STEP 13
    Form$(ZZCnt + 1, 0) = "1"
    Form$(ZZCnt + 4, 0) = "1"
  NEXT

  FOR ZZCnt = 78 TO 93 STEP 5
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
  
  OldBook$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
  
  BCopy VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), TFormLen, 0
  
  IF UBCustRec(1).Status = "F" THEN
    FinalFlag = True
  END IF
  
  DepAmt# = UBCustRec(1).DepositAmt
  
  ERASE UBCustRec
  
RETURN

END SUB

SUB PrintTransHist

  BlockClear
  LookUp RecNo&, "Customer Transaction", 2, True, False
  
  T$ = SPACE$(10)
  MaxLines = 40

  FF$ = CHR$(12)
  Dash80$ = STRING$(80, "-")

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM TotalConsump(1 TO 7) AS LONG
  REDIM DidCnt(1 TO 7) AS INTEGER

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  FOR RevCnt = 1 TO MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME)
    IF LEN(TempRev$) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    ELSE
      RSET T$ = QPTrim$(LEFT$(TempRev$, 8))
      IF RevCnt <= 8 THEN
        RevText1$ = RevText1$ + T$
      ELSE
        RevText2$ = RevText2$ + T$
      END IF
    END IF
  NEXT

  IF LEN(QPTrim$(RevText2$)) > 0 THEN
    Rev2Flag = True
  END IF
  
  IF RecNo& = 0 THEN
    GOTO ExitTransHist
  END IF
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET #UBCust, RecNo&, UBCustRec(1)
  CLOSE UBCust
  
  UBRpt = FREEFILE
  OPEN "UBTRAHIS.RPT" FOR OUTPUT AS UBRpt

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen

  BlockClear

  DisplayUBScrn "PRORPT"
  QPrintRC "Processing Transaction History.", 11, 26, -1

  GOSUB DOTranHistHeader

  ThisTrans& = UBCustRec(1).LastTrans
  
  FirstTrans = True

  DO WHILE ThisTrans& > 0
    GET #UBTran, ThisTrans&, UBTranRec(1)

      IF FirstTrans THEN
        LastDate$ = Num2Date$(UBTranRec(1).TransDate)
        Year = VAL(RIGHT$(LastDate$, 4))
        PDate = Date2Num(LEFT$(LastDate$, 3) + "01-" + QPTrim$(STR$(Year - 1)))
        'PDate = 0
        FirstTrans = False
      END IF

      IF UBTranRec(1).TransDate < PDate THEN
        EXIT DO
      END IF
      GOSUB DOTransDetail
      PRINT #UBRpt, Dash80$
      LineCnt = LineCnt + 1
      IF LineCnt > MaxLines THEN
        PRINT #UBRpt, FF$
        GOSUB DOTranHistHeader
      END IF
    ThisTrans& = UBTranRec(1).PrevTrans
  LOOP
  GOSUB DOTranHistFooter

  CLOSE

  IF NOT AbortFlag THEN
    PrintRptFile "Customer Transaction Report.", "UBTRAHIS.RPT", 1, RetCode, EntryPoint
  END IF

ExitTransHist:
EXIT SUB

DOTransDetail:
  PRINT #UBRpt, Num2Date(UBTranRec(1).TransDate);
  SELECT CASE UBTranRec(1).TransType
    CASE TranUtilityBill, TranUtilityBill + 100
      PRINT #UBRpt, TAB(19); "Utility Bill"; TAB(34); Num2Date$(UBTranRec(1).ReadDate); TAB(49); Num2Date$(UBTranRec(1).PrevDate);
      PRINT #UBRpt, TAB(67); USING "$$#####.##"; UBTranRec(1).TransAmt;
      GOSUB DoMtrDetail
      GOSUB PrintRevDetail
    CASE TranLateCharge, TranLateCharge + 100
      PRINT #UBRpt, TAB(19); "Late Charge"; TAB(67); USING "$$#####.##"; UBTranRec(1).TransAmt
      GOSUB PrintRevDetail
    CASE TranReconnectFee, TranReconnectFee + 100
      PRINT #UBRpt, TAB(19); "Reconnect Fee";
      GOSUB PrintRevDetail
    CASE TranBillPayment, TranBillPayment + 100
      PRINT #UBRpt, TAB(19); "Bill Payment"; TAB(67); USING "$$#####.##"; UBTranRec(1).TransAmt
      GOSUB PrintRevDetail
    CASE TranAppliedDeposit, TranAppliedDeposit + 100
      PRINT #UBRpt, TAB(19); "Applied Deposit";
      GOSUB PrintRevDetail
    CASE TranPenaltyCharge, TranPenaltyCharge + 100
      PRINT #UBRpt, TAB(19); "Penalty Charge"; TAB(67); USING "$$#####.##"; UBTranRec(1).TransAmt
      GOSUB PrintRevDetail
    CASE TranDepositPayment, TranDepositPayment + 100
      PRINT #UBRpt, TAB(19); "Deposit Payment"; TAB(67); USING "$$#####.##"; UBTranRec(1).TransAmt
      GOSUB PrintRevDetail
    CASE TranDraftPayment, TranDraftPayment + 100
      PRINT #UBRpt, TAB(19); "Draft Payment";
      GOSUB PrintRevDetail
    CASE TranRefundDeposit, TranRefundDeposit + 100
      PRINT #UBRpt, TAB(19); "Refund Deposit";
    CASE TranBeginBalance, TranBeginBalance + 100
      PRINT #UBRpt, TAB(19); "Beginning Balance";
    CASE TranUpwardAdjustment, TranUpwardAdjustment + 100
      PRINT #UBRpt, TAB(19); "UP Adjustment  " + LEFT$(UBTranRec(1).BillMsg, 25); TAB(67); USING "$$#####.##"; UBTranRec(1).TransAmt
      GOSUB PrintRevDetail
    CASE TranDownwardAdjustment, TranDownwardAdjustment + 100
      PRINT #UBRpt, TAB(19); "DN Adjustment  " + LEFT$(UBTranRec(1).BillMsg, 25); TAB(67); USING "$$#####.##"; UBTranRec(1).TransAmt
      GOSUB PrintRevDetail
    CASE TranEPPReconcile
      PRINT #UBRpt, TAB(19); "EPP Reconcile  "; UBTranRec(1).TransDesc; TAB(67); USING "$$#####.##"; UBTranRec(1).TransAmt
      GOSUB PrintRevDetail

    CASE TranMiscPayment, TranMiscPayment + 100
      PRINT #UBRpt, TAB(19); "Misc Payment"
      GOSUB PrintRevDetail
  END SELECT
RETURN

DoMtrDetail:
  DidAMeter = False
  FOR MtrCnt = 1 TO 7
    IF UBTranRec(1).MtrTypes(MtrCnt) > 0 THEN
      DidAMeter = True
      SELECT CASE UBTranRec(1).MtrTypes(MtrCnt)
      CASE MtrWaterOnly
        MeterType$ = "      Water"
      CASE MtrSewerOnly
        MeterType$ = "      Sewer"
      CASE MtrCombined
        MeterType$ = "Water/Sewer"
      CASE MtrElectric
        MeterType$ = "   Electric"
      CASE MtrDemand
        MeterType$ = " D Electric"
      CASE MtrGas
        MeterType$ = "  Gas Meter"
      CASE MtrTouchRead
        MeterType$ = " Touch Read"
      CASE MtrLightsService
        MeterType$ = "  L Service"
      END SELECT
      WhatMtrCNT = UBTranRec(1).MtrTypes(MtrCnt)
      IF WhatMtrCNT = 0 THEN
        WhatMtrCNT = 1
      END IF
      GOSUB PrintMtrDetail
    END IF
  NEXT
  IF NOT DidAMeter THEN
    MeterType$ = "        "
    'WhatMtrCNT = 1
    MtrCnt = 1
    GOSUB PrintMtrDetail
  END IF
RETURN

PrintMtrDetail:
  PRINT #UBRpt, TAB(19); MeterType$;
  PRINT #UBRpt, TAB(34); USING "##########"; UBTranRec(1).CurRead(MtrCnt);
  PRINT #UBRpt, TAB(49); USING "##########"; UBTranRec(1).PrevRead(MtrCnt);
  MeterConsp& = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
  IF MeterConsp& < 0 THEN
    MaxMeterAmt& = 10& ^ (LEN(STR$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
    MeterConsp& = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
  END IF
  PRINT #UBRpt, TAB(67); USING "##########"; MeterConsp&
  IF DidAMeter THEN
    TotalConsump(WhatMtrCNT) = TotalConsump(WhatMtrCNT) + MeterConsp&
    DidCnt(WhatMtrCNT) = DidCnt(WhatMtrCNT) + 1
  END IF

  LineCnt = LineCnt + 1
RETURN

PrintRevDetail:
    'PRINT #UBRpt, TAB(49); USING "$$#####.##"; UBTranRec(1).TransAmt
    PrintedOne = False
    FOR RevCnt = 0 TO 7
      IF UBTranRec(1).RevAmt(RevCnt + 1) <> 0 THEN
        PrintedOne = True
        TabStop = (RevCnt * 10) + 1
        PRINT #UBRpt, TAB(TabStop); USING "#######.##"; UBTranRec(1).RevAmt(RevCnt + 1);
      END IF
    NEXT
    IF PrintedOne THEN
      PRINT #UBRpt,
      LineCnt = LineCnt + 1
    END IF
    RevOffset = 7
    PrintedOne = False
    FOR RevCnt = 0 TO 6
      IF UBTranRec(1).RevAmt(RevCnt + 1 + RevOffset) <> 0 THEN
        PrintedOne = True
        TabStop = (RevCnt * 10) + 1
        PRINT #UBRpt, TAB(TabStop); USING "#######.##"; UBTranRec(1).RevAmt(RevCnt + 1 + RevOffset);
      END IF
    NEXT
    IF PrintedOne THEN
      PRINT #UBRpt,
      LineCnt = LineCnt + 1
    END IF

RETURN

DOTranHistHeader:
  LineCnt = 7
  PRINT #UBRpt, TAB(28); "Transaction History Report. "
  PRINT #UBRpt, "Customer: "; UBCustRec(1).CUSTNAME; TAB(57); "Report Date: "; DATE$
  PRINT #UBRpt,
  PRINT #UBRpt, "Trans Date         Trans Type      Cur.Date       Pre.Date       Trans Total"
  PRINT #UBRpt, "                   Meter Type      Cur.Read       Pre.Read             Usage"
  PRINT #UBRpt, RevText1$
  IF Rev2Flag THEN
    PRINT #UBRpt, RevText2$
    LineCnt = 8
  END IF
  PRINT #UBRpt, Dash80$

RETURN

DOTranHistFooter:
  IF FirstTrans THEN
    PRINT #UBRpt, "NO TRANSACTIONS!!!"
    PRINT #UBRpt, Dash80$
  END IF
  FOR MCnt = 1 TO 7
    IF DidCnt(MCnt) > 0 THEN
      PRINT #UBRpt, USING "Average Consumption: #########"; TotalConsump(MCnt) / DidCnt(MCnt)
    END IF
  NEXT
  PRINT #UBRpt, CHR$(12)
RETURN

END SUB

SUB ReconcileCustomer

  UBLog " IN: Reconcile EPP Customer"
  REDIM Revs(1 TO 15) AS STRING * 13

  REDIM UBEPPRec(1) AS UBCustEPPRecType
  UBEPPRecLen = LEN(UBEPPRec(1))

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  RevBal$ = SPACE$(8)
  
  FOR RevCnt = 1 TO MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME)
    IF LEN(TempRev$) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    ELSE
      Revs(RevCnt) = TempRev$
    END IF
  NEXT
  IF NumOfRevs < 15 THEN
    FOR RevCnt = NumOfRevs + 1 TO 15
      Revs(RevCnt) = "N/A"
    NEXT
  END IF

EPPTop:
  BlockClear
  LookUp RecNo&, "EPP Customer", 2, True, False
  IF RecNo& > 0 THEN
    GOSUB CheckCustEPP
    IF EPPCustOK THEN
      GOTO ProcessEPP
    ELSE
      DisplayUBScrn "ERRSCRN1"
      QPrintRC "Not an Equal Pay Customer!", 10, 27, -1
      QPrintRC "Press any key to continue.", 12, 27, -1
      WaitForAction
      GOTO EPPTop
    END IF
  ELSE
    GOTO ExitReconEPP
  END IF

ProcessEPP:

  DO
    HideCursor
    LibFile2Scrn "UBSETUP", "EPPRECON", MonoCode%, Attribute%, ErrorCode%
    GOSUB ShowCustEPPInfo
    ShowCursor
    Ok = MsgBox%("UBSETUP", "EPPRBTN")
    SELECT CASE Ok
    CASE 1
      OKFlag = False
      ExitFlag = True
    CASE 2
      ShowCustHistory RecNo&
    CASE 3
      ExitFlag = True
      Ok = MsgBox%("UBSETUP", "OK2RECON")
      SELECT CASE Ok
      CASE 1
        OKFlag = False
      CASE 2
        OKFlag = True
      END SELECT
    END SELECT
  LOOP UNTIL ExitFlag

  IF OKFlag THEN
    MakeEPPReconTrans RecNo&, True
    DisplayUBScrn "UPDATEOK"
    WaitForAction
  END IF

ExitReconEPP:

EXIT SUB


ReconLoadEPPInfo:
  UBFile = FREEFILE
  OPEN "UBCUSTEP.DAT" FOR RANDOM SHARED AS UBFile LEN = UBEPPRecLen
  GET UBFile, RecNo&, UBEPPRec(1)
  CLOSE
RETURN

CheckCustEPP:
  EPPCustOK = False
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  GET UBFile, RecNo&, UBCustRec(1)
  CLOSE
  IF UBCustRec(1).EPPFlag = "Y" THEN
    GOSUB ReconLoadEPPInfo
    EPPCustOK = True
    CurBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
    EPPBal# = UBEPPRec(1).EPBalTot
    NewBal# = Round#(CurBal# + EPPBal#)
  END IF
RETURN

ShowCustEPPInfo:
  FOR zz = 1 TO 5
    QPrintRC Revs(zz), 11 + zz, 7, -1
    RSET RevBal$ = FUsing$(STR$(UBEPPRec(1).EPInfo(zz).RevBal), "#####.##")
    QPrintRC RevBal$, 11 + zz, 21, -1
  NEXT
  FOR zz = 6 TO 10
    QPrintRC Revs(zz), 6 + zz, 30, -1
    RSET RevBal$ = FUsing$(STR$(UBEPPRec(1).EPInfo(zz).RevBal), "#####.##")
    QPrintRC RevBal$, 6 + zz, 44, -1
  NEXT
  FOR zz = 11 TO 15
    QPrintRC Revs(zz), zz + 1, 53, -1
    RSET RevBal$ = FUsing$(STR$(UBEPPRec(1).EPInfo(zz).RevBal), "#####.##")
    QPrintRC RevBal$, zz + 1, 67, -1
  NEXT
  QPrintRC UBCustRec(1).Book, 5, 26, -1
  QPrintRC UBCustRec(1).SEQNUMB, 5, 29, -1
  QPrintRC UBCustRec(1).Status, 5, 47, -1
  LSET RevBal$ = FUsing$(STR$(RecNo&), "#####")
  QPrintRC RevBal$, 5, 65, -1
  QPrintRC UBCustRec(1).CUSTNAME, 7, 26, -1
  QPrintRC UBCustRec(1).ADDR1, 8, 26, -1
  QPrintRC UBCustRec(1).CITY, 9, 26, -1
  QPrintRC UBCustRec(1).STATE, 9, 54, -1
  QPrintRC UBCustRec(1).ZIPCODE, 9, 63, -1
  RSET RevBal$ = FUsing$(STR$(CurBal#), "#####.##")
  QPrintRC RevBal$, 18, 33, -1
  RSET RevBal$ = FUsing$(STR$(EPPBal#), "#####.##")
  QPrintRC RevBal$, 18, 58, -1
  RSET RevBal$ = FUsing$(STR$(NewBal#), "#####.##")
  QPrintRC RevBal$, 19, 44, -1
RETURN

END SUB

SUB SaveNewCustRec (Form$(), RecNo&)

'NOTE: The owner info is saved in the main customer editing procedure.

  REDIM UBCustRec(1) AS NewUBCustRecType

  FormLen = LEN(Form$(0, 0))    'Length of Cust Form$(0,0) Element
  
  UBCustRecLen = LEN(UBCustRec(1))              'Length of Cust Record Structure
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1)), FormLen, 0
  
  REDIM UBCustPIN(1) AS UBPINType               'Pin info array
  
  IF FileSize("UBCUSPIN.DAT") > 0 THEN          'If there was a previous pin then
    PHandle = FREEFILE
    OPEN "UBCUSPIN.DAT" FOR RANDOM SHARED AS #PHandle LEN = 4
    GET #PHandle, 1, UBCustPIN(1)               'get last pin used info
    CLOSE #PHandle
  END IF
  
  NextPIN& = UBCustPIN(1).PIN + 1               'Increment last pin used
  
  UBCustPIN(1).PIN = NextPIN&
  PHandle = FREEFILE            'Save new last pin info
  OPEN "UBCUSPIN.DAT" FOR RANDOM SHARED AS #PHandle LEN = 4
  PUT #PHandle, 1, UBCustPIN(1)
  CLOSE #PHandle
  
  UBCustRec(1).CustPIN = NextPIN&
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  NextRec& = NumOfRecs& + 1
  
  UBCustRec(1).MessageRec = 0   ' Set Pointers To Zero
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

  PUT UBFile, RecNo&, UBCustRec(1)
  
  CLOSE UBFile
  
'01-20-97 Added skip reindexing if name not changed
  IF UBCustRec(1).SEARCH <> UBCustRec(2).SEARCH OR UBCustRec(1).CUSTNAME <> UBCustRec(2).CUSTNAME THEN
    ReindexFlag = True
  END IF
  
'01-20-97 Added skip reindexing if book or seq number not changed
  IF (UBCustRec(1).Book <> UBCustRec(2).Book) OR (UBCustRec(1).SEQNUMB <> UBCustRec(2).SEQNUMB) THEN
    ReindexFlag = True
  END IF
  
  IF ReindexFlag THEN
    KillFile "UBOOKSEQ.DAT"
    ReIndexSystem False
  END IF

  RestScrn TempScrn()

  ERASE UBCustRec, TempScrn

END SUB

SUB UBCustomerMenu
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 10)
  
  MChoice$(1) = " Add a New Customer     "
  MChoice$(2) = " Edit Existing Customer "
  MChoice$(3) = " Set a Customer to Final"
  MChoice$(4) = " Delete Customer Account"
  MChoice$(5) = " Reconcile Equal Pay Account"
  MChoice$(6) = " Customer Consumption History"
  MChoice$(7) = " Customer Transaction History"
  MChoice$(8) = " Quick Customer Listing by Name"
  MChoice$(9) = " Quick Customer Listing by Location"
  MChoice$(10) = " Quick Customer Listing by Rate Code"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
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
      BlockClear
      LookUp RecNo&, "Customer", 2, True, False
      IF RecNo& > 0 THEN
        AddEditCustomer RecNo&
      END IF
    CASE 3
      'SetCustFinal
    CASE 4
      'DeleteCust
    CASE 5
      ReconcileCustomer

    CASE 6
      'PrintConsumpHist
    CASE 7
      PrintTransHist
    CASE 8
      'PrintLocaList False
    CASE 9
      'PrintLocaList True
    CASE 10
      'PrintCustByRate
    END SELECT
  LOOP
  
  ' RUN "ubmenu"
  
END SUB

