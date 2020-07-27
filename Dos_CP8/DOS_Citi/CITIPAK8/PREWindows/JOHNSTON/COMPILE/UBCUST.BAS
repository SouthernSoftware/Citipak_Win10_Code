DEFINT A-Z
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

  CONST False = 0, True = NOT False
  
  CrLf$ = CHR$(13) + CHR$(10)

SUB AddEditCustomer (RecNo&)

  REDIM UBOwnerRec(1) AS UBOwnerRecType
  UBOwnerRecLen = LEN(UBOwnerRec(1))
  OFirstTime = True

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
        CASE 6 TO 14            '--F3 Wrk Hist
          PressButton F3Key, 22, 6, 14
        CASE 16 TO 24           '--F4 Tran Hist
          PressButton F4KEY, 22, 16, 24
        CASE 26 TO 34           '--F6 Cons Hist
          PressButton F6KEY, 22, 26, 34
        CASE 36 TO 43           '--F7 Msgs
          IF RecNo& > 0 THEN
            PressButton F7KEY, 22, 36, 43
          END IF
        CASE 45 TO 52           '--F5 Owner
          PressButton F8KEY, 22, 45, 52
        CASE 54 TO 63           '--F10 Save Button
          PressButton F10Key, 22, 54, 63
        CASE 65 TO 74           '--ESC cancel button
          PressButton 27, 22, 65, 74
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
  
  'IF RecNo& = 0 THEN
    Choice$(4, 3) = "Pending"
  'END IF

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
    QPrintRC "F7", 22, 37, 254
  ELSE
    QPrintRC "F7", 22, 37, 126
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
  OldBook$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
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

SUB PrintConsumpHist

  BlockClear
  LookUp RecNo&, "Customer Consumption", 2, True, False

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  IF INSTR(UBSetUpRec(1).UtilName, "TROY") > 0 THEN
    TroyFlag = True
  END IF

  IF RecNo& = 0 THEN
    GOTO ExitConsumpHist
  END IF

  Dash80$ = STRING$(80, "-")

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET #UBCust, RecNo&, UBCustRec(1)
  CLOSE UBCust

'  Cubic = False
'  FOR MTCnt = 1 TO 7
'    IF UBCustRec(1).LocMeters(MTCnt).MtrUnit = "C" THEN
'      Cubic = True
'      EXIT FOR
'    END IF
'  NEXT

  UBRpt = FREEFILE
  OPEN "UBCONSMP.RPT" FOR OUTPUT AS UBRpt

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen

  BlockClear

  DisplayUBScrn "PRORPT"
  QPrintRC "Processing Consumption History.", 11, 26, -1

  GOSUB DoConsRptHeader

  ThisTrans& = UBCustRec(1).LastTrans

  DO WHILE ThisTrans& > 0
    GET #UBTran, ThisTrans&, UBTranRec(1)
    IF UBTranRec(1).TransType = TranUtilityBill THEN
      GOSUB PrintConsDetail
      DidCnt = DidCnt + 1
      IF DidCnt = 12 THEN
        EXIT DO
      END IF
    END IF
    ThisTrans& = UBTranRec(1).PrevTrans
  LOOP
  GOSUB DoConsFooter

  CLOSE

  IF NOT AbortFlag THEN
    PrintRptFile "Customer Consumption Report.", "UBCONSMP.RPT", 1, RetCode, EntryPoint
  END IF


ExitConsumpHist:
EXIT SUB

PrintConsDetail:
  
  DidAMeter = False
  FOR MtrCnt = 1 TO 7
    MtrMulti# = 0
    FOR MCnt = 1 TO 7
      IF UBTranRec(1).MtrTypes(MtrCnt) = GetCustMeterType%(UBCustRec(), MCnt) THEN
        MtrMulti# = UBCustRec(1).LocMeters(MCnt).MtrMulti
        EXIT FOR
      END IF
    NEXT
    IF MtrMulti# = 0 THEN
      IF TroyFlag THEN
        MtrMulti# = 100
      ELSE
        MtrMulti# = 1
      END IF
    END IF

    IF UBTranRec(1).MtrTypes(MtrCnt) > 0 THEN
      DidAMeter = True
      SELECT CASE UBTranRec(1).MtrTypes(MtrCnt)
      CASE MtrWaterOnly
        MeterType$ = "      Water"
        F$ = "W"
      CASE MtrSewerOnly
        MeterType$ = "      Sewer"
        F$ = "S"
      CASE MtrCombined
        MeterType$ = "Water/Sewer"
        F$ = "C"
      CASE MtrElectric
        MeterType$ = "   Electric"
        F$ = "E"
      CASE MtrDemand
        MeterType$ = " D Electric"
        F$ = "D"
      CASE MtrGas
        MeterType$ = "  Gas Meter"
        F$ = "G"
      CASE MtrTouchRead
        MeterType$ = " Touch Read"
        F$ = "T"
      CASE MtrLightsService
        MeterType$ = "  L Service"
      END SELECT
      FOR CCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(CCnt).MtrType = F$ THEN
          IF UBCustRec(1).LocMeters(CCnt).MtrUnit = "C" THEN
            CubMeter = True
          ELSE
            CubMeter = False
          END IF
          EXIT FOR
        END IF
      NEXT
      GOSUB PrintThisMeter
    END IF
  NEXT
  IF NOT DidAMeter THEN
    MeterType$ = "        "
    MtrCnt = 1
    GOSUB PrintThisMeter
  END IF

RETURN

PrintThisMeter:
  PRINT #UBRpt, Num2Date(UBTranRec(1).TransDate); TAB(19); MeterType$;
  PRINT #UBRpt, TAB(34); USING "##########"; UBTranRec(1).CurRead(MtrCnt);
  PRINT #UBRpt, TAB(46); USING "##########"; UBTranRec(1).PrevRead(MtrCnt);
  MeterConsp# = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
  IF MeterConsp# < 0 THEN
    MaxMeterAmt& = 10& ^ (LEN(STR$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
    MeterConsp# = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
  END IF
  MeterConsp# = MeterConsp# * MtrMulti#
  IF CubMeter THEN
    MeterConsp# = MeterConsp# * 7.481
  END IF
  PRINT #UBRpt, TAB(56); USING "##########"; MeterConsp#;
  IF UBTranRec(1).ReadDate <= 0 THEN
    PRINT #UBRpt, "     ??-??-????"
  ELSE
    PRINT #UBRpt, "     "; Num2Date$(UBTranRec(1).ReadDate)
  END IF
  TotalConsp# = TotalConsp# + MeterConsp#

RETURN

DoConsRptHeader:
  PRINT #UBRpt, TAB(28); "Consumption History Report. "
  PRINT #UBRpt,
  PRINT #UBRpt, "Customer: "; UBCustRec(1).CUSTNAME; TAB(57); "Report Date: "; DATE$
  PRINT #UBRpt,
  PRINT #UBRpt, "Transaction                         Current   Previous"
  PRINT #UBRpt, "   Date            Meter Type       Reading    Reading       Usage    ReadDate"
  PRINT #UBRpt, Dash80$
RETURN

DoConsFooter:
  IF DidCnt > 0 THEN
    PRINT #UBRpt, Dash80$
    PRINT #UBRpt, USING "Average Consumption: #########"; TotalConsp# / DidCnt
  ELSE
    PRINT #UBRpt, "NO TRANSACTIONS!!!"
    PRINT #UBRpt, Dash80$
  END IF
RETURN
END SUB

SUB PrintCustByRate

  SHARED Choice$()

  MaxLines = 59
  PageNo = 0
  Dash80$ = STRING$(80, "-")

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it

'*****************
  LibName$ = "UB"
  ScrnName$ = "CUSBYRAT"

  '--Get the total number of fields from all pages
  NumFlds = NumFlds + LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  '--Increment StartEl to next

  REDIM UBRateTblRec(1) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))
  NumOfRateRecs = GetNumRateRecs

  'add show a message
  IF NumOfRateRecs <= 0 THEN
    GOTO ExitCustByRate
  END IF

  REDIM Choice$(0 TO NumOfRateRecs, 0 TO 1)
  Choice$(0, 0) = "1"
  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRateRecs
    GET UBFile, Cnt, UBRateTblRec(1)
    Choice$(Cnt, 0) = SPACE$(30)
    LSET Choice$(Cnt, 0) = UBRateTblRec(1).RATECODE
    MID$(Choice$(Cnt, 0), 6, 24) = UBRateTblRec(1).RATEDESC
  NEXT
  CLOSE
  'SortStr2 VARPTR(Choice$(1)), Cnt - 1, 0

  RateRec = 0
  PrnRate = False

  Action = 1
  FirstTime = True

  BlockClear

  DisplayUBScrn "CUSBYRAT"

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      LSET Form$(2, 0) = "N"
      LSET Form$(3, 0) = "Y"
      Action = 2
    END IF
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key 'F10
      WhatRate$ = QPTrim$(LEFT$(Form$(1, 0), 4))
      IF LEN(WhatRate$) > 0 THEN
        IF QPTrim$(Form$(2, 0)) = "Y" THEN
          PrnRate = True
          FOR Cnt = 1 TO NumOfRateRecs
            ThisRate$ = QPTrim$(LEFT$(Choice$(Cnt, 0), 4))
            IF ThisRate$ = WhatRate$ THEN
              RateRec = Cnt
              EXIT FOR
            END IF
          NEXT
        END IF
        SELECT CASE Form$(3, 0)
        CASE "Y"
          ActOnly = True
        END SELECT
        OkFlag = True

      END IF
    CASE EscKey
      ExitFlag = True
    END SELECT
    
    '--Check screen page
  LOOP UNTIL ExitFlag OR OkFlag

  IF ExitFlag THEN
    GOTO ExitCustByRate
  END IF
  
'****************

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBRpt = FREEFILE
  OPEN "UBCSBYRT.RPT" FOR OUTPUT AS UBRpt

  BlockClear
  ShowProcessingScrn "Customer Listing By Rate Code."

  GOSUB CustByRateHeader

  FOR Cnt = 1 TO IdxNumOfRecs
    GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      IF ActOnly THEN
        IF UBCustRec(1).Status <> "A" THEN
          GOTO SkipCustRate
        END IF
      END IF
      FOR SCnt = 1 TO 15
        ThisRate$ = QPTrim$(UBCustRec(1).Serv(SCnt).RATECODE)
        IF WhatRate$ = ThisRate$ THEN
          PRINT #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; " "; FUsing$(STR$(IdxBuff(Cnt).RecNum), "#####"); "  "; LEFT$(UBCustRec(1).CUSTNAME, 25); "  "; LEFT$(UBCustRec(1).SERVADDR, 30); " "; UBCustRec(1).Status
          CustCnt = CustCnt + 1
          LineCnt = LineCnt + 1
          SELECT CASE UBCustRec(1).Status
          CASE "A"
            Active = Active + 1
          CASE "F"
            Final = Final + 1
          CASE "I"
            InActive = InActive + 1
          CASE "B"
            Balance = Balance + 1
          CASE ELSE
            UnKnown = UnKnown + 1
          END SELECT
          EXIT FOR
        END IF
      NEXT

      IF LineCnt > MaxLines THEN
        PRINT #UBRpt, CHR$(12)
        GOSUB CustByRateHeader
      END IF
    END IF
SkipCustRate:
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp Cnt, IdxNumOfRecs
  NEXT

  GOSUB CustByRateTotals

  ERASE IdxBuff, UBCustRec   'free up memory

  IF PrnRate THEN
    GOSUB PrintRateCode
  END IF

  CLOSE UBCust, UBRpt

  IF NOT AbortFlag THEN
    PrintRptFile "Customer Listing by Rate Code.", "UBCSBYRT.RPT", 1, RetCode, EntryPoint
  END IF

ExitCustByRate:
  EXIT SUB

CustByRateHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Customer Listing By Rate Code      "; "Date: "; DATE$; TAB(70); "Page: "; PageNo
  PRINT #UBRpt, "RATE CODE: "; WhatRate$
  PRINT #UBRpt, "Location   Acct.  Customer Name             Service Address             Status"
  PRINT #UBRpt, Dash80$
  LineCnt = 4
RETURN

CustByRateTotals:
  'PageNo = PageNo + 1
  PRINT #UBRpt,
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Customer Summary"
  PRINT #UBRpt,
  PRINT #UBRpt, USING "  Active: ####"; Active
  PRINT #UBRpt, USING "   Final: ####"; Final
  PRINT #UBRpt, USING "Inactive: ####"; InActive
  PRINT #UBRpt, USING " Balance: ####"; Balance
  PRINT #UBRpt, USING " Unknown: ####"; UnKnown
  PRINT #UBRpt, USING " Deleted: ####"; DeletedCnt
  PRINT #UBRpt,
  PRINT #UBRpt, USING "   TOTAL: ####"; CustCnt
  PRINT #UBRpt, CHR$(12)
RETURN

PrintRateCode:

  REDIM UBRateTblRec(1) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))

  NumRateRecs = FileSize("UBRATE.DAT") \ UBRateTblRecLen

  IF NumRateRecs = 0 THEN
    GOTO PrintRateExit
  END IF

  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  GET UBFile, RateRec, UBRateTblRec(1)
  CLOSE UBFile

  REDIM StepText(1 TO 10) AS STRING * 40

  GOSUB PrintRateHeader
    PRINT #UBRpt, "       Rate Code:  "; UBRateTblRec(1).RATECODE
    PRINT #UBRpt, "     Description:  "; UBRateTblRec(1).RATEDESC
    PRINT #UBRpt, "  Minimum Charge:"; FUsing$(STR$(UBRateTblRec(1).MINAMT), "#######.##")
    PRINT #UBRpt, "   Minimum Units:"; FUsing$(STR$(UBRateTblRec(1).MINUNITS), "##########")
    PRINT #UBRpt, "      Max Amount:"; FUsing$(STR$(UBRateTblRec(1).MaxAmt), "######.##")
    PRINT #UBRpt, "      [ Step ]        [ Beg Unit ]     [ Amount/Unit ]"
    FOR RCnt = 1 TO 10
      LSET StepText$(RCnt) = ""
      IF UBRateTblRec(1).TblBreaks(RCnt).UNITS >= 0 THEN
        MID$(StepText$(RCnt), 8) = FUsing$(STR$(UBRateTblRec(1).TblBreaks(RCnt).UNITS), "########")
      END IF
      IF UBRateTblRec(1).TblBreaks(RCnt).UNITAMT >= 0 THEN
        MID$(StepText$(RCnt), 25) = FUsing$(STR$(UBRateTblRec(1).TblBreaks(RCnt).UNITAMT), "####.######")
      END IF
    NEXT
    PRINT #UBRpt, "     First Break:"; StepText$(1)
    PRINT #UBRpt, "    Second Break:"; StepText$(2)
    PRINT #UBRpt, "     Third Break:"; StepText$(3)
    PRINT #UBRpt, "    Fourth Break:"; StepText$(4)
    PRINT #UBRpt, "     Fifth Break:"; StepText$(5)
    PRINT #UBRpt, "     Sixth Break:"; StepText$(6)
    PRINT #UBRpt, "   Seventh Break:"; StepText$(7)
    PRINT #UBRpt, "    Eighth Break:"; StepText$(8)
    PRINT #UBRpt, "     Ninth Break:"; StepText$(9)
    PRINT #UBRpt, "        All Over:"; StepText$(10)
    PRINT #UBRpt,
    PRINT #UBRpt, Dash80$
    PRINT #UBRpt, CHR$(12)

  ERASE UBRateTblRec, StepText

  GOTO PrintRateExit

PrintRateHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Utility Rate Table Listing."
  PRINT #UBRpt, "RATE CODE: "; WhatRate$; TAB(70); "Page:"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt, Dash80$
  NumPrinted = 0
RETURN

PrintRateExit:

RETURN

END SUB

SUB PrintLocaList (LocFlag)
  
  MaxLines = 59
  PageNo = 0
  Dash80$ = STRING$(80, "-")
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBCustBlank(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  IdxRecLen = 4 'we are using a long integer

  IF LocFlag THEN
    IdxName$ = "UBCUSTBK.IDX"
    Title$ = "Quick Customer Listing by Location."
  ELSE
    IdxName$ = "UBCUSTNM.IDX"
    Title$ = "Quick Customer Listing by Name."
  END IF

  IdxFileSize& = FileSize("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH IdxName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBLOLIST.RPT" FOR OUTPUT AS UBRpt
  
  BlockClear
  ShowProcessingScrn Title$
  GOSUB DoLocaRptHeader
  
  FOR Cnt = 1 TO IdxNumOfRecs
    GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB DoLocaRptHeader
    END IF
    IF UBCustRec(1).DelFlag = 0 THEN

    'IF UBCustRec(1).BillCycl <> 2 THEN
    '  GOTO LocSkip
    'END IF

    'CTy$ = QPTrim$(UBCustRec(1).CustType)
    'IF CTy$ <> "R" THEN
      PRINT #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; " ";
      PRINT #UBRpt, FUsing$(STR$(IdxBuff(Cnt).RecNum), "#####"); "  ";
      PRINT #UBRpt, LEFT$(UBCustRec(1).CUSTNAME, 25); "  ";
      PRINT #UBRpt, LEFT$(UBCustRec(1).SERVADDR, 25); " ";
      'PRINT #UBRpt, UBCustRec(1).Status; "     "; UBCustRec(1).CustType
      PRINT #UBRpt, UBCustRec(1).Status; "     "; UBCustRec(1).BillCycl
      CustCnt = CustCnt + 1
      SELECT CASE UBCustRec(1).Status
      CASE "A"
        Active = Active + 1
      CASE "F"
        Final = Final + 1
      CASE "I"
        InActive = InActive + 1
      CASE "B"
        Balance = Balance + 1
      CASE ELSE
        UnKnown = UnKnown + 1
      END SELECT
      LineCnt = LineCnt + 1
    ELSE
      DeletedCnt = DeletedCnt + 1
    END IF
    'END IF
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
LocSkip:
    ShowPctComp Cnt, IdxNumOfRecs
  NEXT
  
  GOSUB DoLocaRptTotals
  
  CLOSE UBCust, UBLoca, UBRpt
  
  ERASE IdxBuff, UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile Title$, "UBLOLIST.RPT", 1, RetCode, EntryPoint
  END IF
  
  'KillFile "UBLOLIST.RPT"
  
EXIT SUB
  
DoLocaRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Customer Listing Report      "; "Date: "; DATE$; TAB(70); "Page: "; PageNo
  PRINT #UBRpt, "           Acct                                                    Acct    Post"
  PRINT #UBRpt, "Location    No.  Customer Name              Service Address        Status  Route"
  PRINT #UBRpt, Dash80$
  LineCnt = 4
RETURN
  
DoLocaRptTotals:
  PageNo = PageNo + 1
  PRINT #UBRpt,
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Customer Summary"
  PRINT #UBRpt,
  PRINT #UBRpt, USING "  Active: #####"; Active
  PRINT #UBRpt, USING "   Final: #####"; Final
  PRINT #UBRpt, USING "Inactive: #####"; InActive
  PRINT #UBRpt, USING " Balance: #####"; Balance
  PRINT #UBRpt, USING " Unknown: #####"; UnKnown
  PRINT #UBRpt, USING " Deleted: #####"; DeletedCnt
  PRINT #UBRpt,
  PRINT #UBRpt, USING "   TOTAL: #####"; CustCnt
  PRINT #UBRpt, CHR$(12)
RETURN
  
  
END SUB

SUB PrintTransHist (DetailFlag)

'  DetailFlag = True

  BlockClear
  LookUp RecNo&, "Customer Transaction", 2, True, False

  IF RecNo& = 0 THEN
    GOTO ExitTransHist
  END IF
  
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
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET #UBCust, RecNo&, UBCustRec(1)
  CLOSE UBCust

  FOR MTCnt = 1 TO 7
    IF UBCustRec(1).LocMeters(MTCnt).MtrUnit = "C" THEN
      Cubic = True
      EXIT FOR
    END IF
  NEXT

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
   '     EXIT DO
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
      PRINT #UBRpt, TAB(19); "Utility Bill";
      IF DetailFlag THEN
        PRINT #UBRpt, TAB(34); Num2Date$(UBTranRec(1).ReadDate); TAB(49); Num2Date$(UBTranRec(1).PrevDate);
      END IF
      PRINT #UBRpt, TAB(67); USING "$$#####.##"; UBTranRec(1).TransAmt;
      IF DetailFlag THEN
        GOSUB DoMtrDetail
        GOSUB PrintRevDetail
      ELSE
        PRINT #UBRpt,
      END IF
    CASE TranLateCharge, TranLateCharge + 100
      PRINT #UBRpt, TAB(19); "Late Charge"; TAB(67); USING "$$#####.##"; UBTranRec(1).TransAmt
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranReconnectFee, TranReconnectFee + 100
      PRINT #UBRpt, TAB(19); "Reconnect Fee";
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranBillPayment, TranBillPayment + 100
      PRINT #UBRpt, TAB(19); "Bill Payment"; TAB(67); USING "$$#####.##"; UBTranRec(1).TransAmt
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranAppliedDeposit, TranAppliedDeposit + 100
      PRINT #UBRpt, TAB(19); "Applied Deposit";
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranPenaltyCharge, TranPenaltyCharge + 100
      PRINT #UBRpt, TAB(19); "Penalty Charge"; TAB(67); USING "$$#####.##"; UBTranRec(1).TransAmt
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranDepositPayment, TranDepositPayment + 100
      PRINT #UBRpt, TAB(19); "Deposit Payment"; TAB(67); USING "$$#####.##"; UBTranRec(1).TransAmt
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranDraftPayment, TranDraftPayment + 100
      PRINT #UBRpt, TAB(19); "Draft Payment";
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranRefundDeposit, TranRefundDeposit + 100
      PRINT #UBRpt, TAB(19); "Refund Deposit"; TAB(67); USING "$$#####.##"; UBTranRec(1).TransAmt

    CASE TranBeginBalance, TranBeginBalance + 100
      PRINT #UBRpt, TAB(19); "Beginning Balance";
    CASE TranUpwardAdjustment, TranUpwardAdjustment + 100
      PRINT #UBRpt, TAB(19); "UP Adjustment  " + LEFT$(UBTranRec(1).BillMsg, 25); TAB(67); USING "$$#####.##"; UBTranRec(1).TransAmt
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranDownwardAdjustment, TranDownwardAdjustment + 100
      PRINT #UBRpt, TAB(19); "DN Adjustment  " + LEFT$(UBTranRec(1).BillMsg, 25); TAB(67); USING "$$#####.##"; UBTranRec(1).TransAmt
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranMiscPayment, TranMiscPayment + 100
      PRINT #UBRpt, TAB(19); "Misc Payment"
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
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
  IF Cubic THEN
    MeterConsp& = MeterConsp& * 7.481
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

    'PRINT #UBRpt, USING "#######"; ThisTrans&;

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
  IF DetailFlag THEN
    PRINT #UBRpt, "Trans Date         Trans Type      Cur.Date       Pre.Date       Trans Total"
  ELSE
    PRINT #UBRpt, "Trans Date         Trans Type                                    Trans Total"
  END IF
  IF DetailFlag THEN
    PRINT #UBRpt, "                   Meter Type      Cur.Read       Pre.Read             Usage"
  END IF
  IF DetailFlag THEN
    PRINT #UBRpt, RevText1$
    IF Rev2Flag THEN
      PRINT #UBRpt, RevText2$
      LineCnt = 8
    END IF
  ELSE
    LineCnt = 5
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
    'SHELL
    ReIndexSystem False
  END IF

  RestScrn TempScrn()

  ERASE UBCustRec, TempScrn

END SUB

SUB UBCustomerMenu
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 10)
  
  MChoice$(1) = " Add a New Customer"
  MChoice$(2) = " Edit Existing Customer"
  MChoice$(3) = " Set a Customer to Final"
  MChoice$(4) = " Delete Customer Account"
  MChoice$(5) = " Customer Consumption History"
  MChoice$(6) = " Customer Transaction History Summary"
  MChoice$(7) = " Customer Transaction History Detailed"
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
      SetCustFinal
    CASE 4
      DeleteCust
    CASE 5
      PrintConsumpHist
    CASE 6
      PrintTransHist False
    CASE 7
      PrintTransHist True
    CASE 8
      PrintLocaList False
    CASE 9
      PrintLocaList True
    CASE 10
      PrintCustByRate
    END SELECT
  LOOP
  
END SUB

