DEFINT A-Z
DECLARE SUB DeleteCust ()
DECLARE SUB ReUseCust ()
DECLARE SUB SetCustFinal ()
DECLARE SUB PrintLocaList ()
DECLARE SUB AddCustDeposit (DepAmt#)
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE SUB ShowWrkOrdHistory (CustRec&)
DECLARE SUB ShowCustConsHist (CustRec&)
DECLARE SUB AddEditLocation (RecNo&, FromFlag%)
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ClearScrn ()
DECLARE FUNCTION GetNumOfCust% ()
DECLARE SUB SaveOldCustRec (Form$(), RecNo&)
DECLARE SUB CustLookUp (RecNo%)
DECLARE SUB Search4Cust (SEARCH$, RecNo&, CLSFlag%, LocationFlag%)
DECLARE SUB PrintCustList ()
DECLARE SUB SaveNewCustRec (Form$(), RecNo&)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE FUNCTION Compare3% (BYVAL Seg1, BYVAL Adr1, BYVAL Seg2, BYVAL Adr2, NumBytes)
DECLARE SUB InsertT (SEG StartElement AS ANY, ElSize%, NumEls%)
DECLARE SUB EditCustomer ()
DECLARE SUB AddCustomer ()
DECLARE SUB LoadMeterRec (Form$())
DECLARE SUB AddNewMeter (Cust$, EdFlag)
DECLARE SUB EditMeter ()
DECLARE SUB SaveMeterRec (Form$())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB RestScrn (Array%())
DECLARE SUB AddEditCustomer (RecNo&, LocatFlag%, F5Flag%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB SaveCustRec (Form$())
DECLARE SUB AddNewCustomer ()
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB SaveSetUpRec (Form$())
  'DECLARE SUB LoadSetUpRec (Form$(), OKFlag)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB EditCust ()
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB StuffBuf (Ky$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB MScrnSave (UlRow%, UlCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (UlRow%, UlCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB CursorOff ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB ShowCursor ()
DECLARE SUB WaitForAction ()
  
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE FUNCTION FLof& (Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB HideCursor ()
DECLARE FUNCTION QPValL& (Number$)
DECLARE SUB LoadCustRec (Form$(), RecNo&, LocatRec&, TFormLen%, LockedFlag%)
DECLARE SUB BlockClear ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num (WhatDate$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB KillFile (FileName$)
DECLARE FUNCTION FmtBook$ (Book$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION Chk4DupeBookSeqNum (Book$, SeqNum$)
DECLARE SUB CustMessageSystem (RecNo&)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'PageInfo.BI'                  'Form Page info
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'newcust.bi'
  ''$INCLUDE: 'chkcust.bi'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'ubTrans.BI'
  ''$INCLUDE: 'oubtypes.BI'
  
  CONST False = 0, True = NOT False
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE

SUB AddCustDeposit (DepAmt#)
  
  LibName$ = "UB"
  ScrnName$ = "DEPENTRY"
  
  REDIM TScrnArray(1)
  REDIM TScrnArray2(1)
  SaveScrn TScrnArray()
  MPaintBox 3, 5, 22, 75, 8
  SaveScrn TScrnArray2()
  
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
  
  FOR F = 1 TO NumFlds
    TFormLen = TFormLen + Fld(F).StorLen
  NEXT
  
  Form$(0, 0) = SPACE$(TFormLen)
  
  Action = 1
  FirstTime = True
  
  DisplayUBScrn "DEPENTRY"
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = STR$(DepAmt#)
      SaveField 1, Form$(), Fld(), BadField
      Action = 2
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key 'F10
      DepAmt# = Value(Form$(1, 0), ECode)
      RestScrn TScrnArray2()
      ExitFlag = True
    CASE EscKey
      RestScrn TScrnArray2()
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 14   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 33 TO 46           '--F10 Save Button
          PressButton F10Key, 14, 33, 46
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
  LOOP UNTIL ExitFlag
  
  RestScrn TScrnArray()
  ERASE TScrnArray, TScrnArray2, Frm, Form$, Fld
  
END SUB

SUB AddEditCustomer (RecNo&, LocatFlag%, F5Flag%)
  
  REDIM ScrnArray(0)
  'TempName$ = SPACE$(25)
  'add a flag to show if called from location edit
  'stop f5
  
  REDIM UBRateRecs(1) AS UBRateTblRecType
  UBRateRecLen = LEN(UBRateRecs(1))
  NumOfRateRecs = FileSize("UBRATE.DAT") \ UBRateRecLen
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  
  'IF UBSetUpRec(1).FLYDeposit = "Y" THEN
  '  DepositFlag = True
  'END IF
  
  FOR RevCnt = 1 TO MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME)
    IF LEN(TempRev$) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    END IF
  NEXT
  
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
  
  IF DepositFlag AND RecNo& = 0 THEN
    LibFile2Scrn "UB", "F7DEPBTN", -1, -2, ErrorCode
  END IF
  
  ShowCursor
  
  'Frm(1).StayOnField = True
  
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
    
    'Format the Location Number fields
    
    IF NOT FinialFlag THEN
      IF Frm(1).FldNo = 3 AND Frm(1).PrevFld < 3 THEN            '>2
        Form$(1, 0) = FmtBook$(Form$(1, 0))
        Form$(2, 0) = FmtSeqN$(Form$(2, 0))
        NBook$ = QPTrim$(Form$(1, 0)) + "-" + QPTrim$(Form$(2, 0))
        IF OldBook$ <> NBook$ THEN
          'SaveField 0, Form$(), Fld(), BadField
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
        IF DepAmt# > 0 THEN
          GOSUB SaveCustDeposit
        END IF
      ELSEIF RecNo& > 0 THEN
        SaveOldCustRec Form$(), RecNo&
      ELSE
        
      END IF

      IF LocatFlag THEN ExitFlag = True
      DisplayUBScrn "UPDATEOK"
      WaitForAction
      ExitFlag = True
      RestScrn ScrnArray()
      
    CASE F3Key
      ShowWrkOrdHistory RecNo&
      Action = 2

    CASE F4KEY
      IF RecNo& > 0 THEN
        ShowCustHistory RecNo&
        Action = 2
      END IF

    CASE F6KEY
      ShowCustConsHist RecNo&
      Action = 2

    CASE F7KEY

      IF RecNo& > 0 THEN
        LastFld = Frm(1).FldNo
        SaveScrn ScrnArray()
        CustMessageSystem RecNo&
        RestScrn ScrnArray()
        Frm(1).FldNo = LastFld
        Action = 1
      END IF
      'IF RecNo& = 0 AND DepositFlag THEN
      '  LastFld = Frm(1).FldNo
      '  AddCustDeposit DepAmt#
      '  Frm(1).FldNo = LastFld
      '  Action = 1
      'END IF
      
    CASE EscKey
      IF EditedFlag THEN
        SaveFlag = PromptSaveData
        SELECT CASE SaveFlag
        CASE True               'user wants to save
          StuffBuf CHR$(0) + CHR$(ABS(F10Key))
        CASE False              'user wants to abandon
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
        CASE 6 TO 15            '--F3 Wrk Hist
          PressButton F3Key, 22, 6, 15
        CASE 17 TO 26           '--F4 Tran Hist
          PressButton F4KEY, 22, 17, 26
        CASE 28 TO 37           '--F6 Cons Hist
          PressButton F6KEY, 22, 28, 37
        CASE 39 TO 50           '--F7 Msgs
          IF RecNo& > 0 THEN
            PressButton F7KEY, 22, 39, 50
          END IF
        CASE 52 TO 61           '--F10 Save Button
          PressButton F10Key, 22, 52, 61
        CASE 63 TO 74           '--ESC cancel button
          PressButton 27, 22, 63, 74
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
      IF DepositFlag AND RecNo& = 0 THEN
        LibFile2Scrn "UB", "F7DEPBTN", -1, -2, ErrorCode
      END IF
      GOSUB UpDateDisplay
      IF Scr = 3 THEN
        GOSUB DisplayRevSources
      END IF
      
    END IF
    
  LOOP UNTIL ExitFlag
  
ExitCustEdit:
  'HideCursor
EXIT SUB
  
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
  
  FOpenS "UBRATE.DAT", Handle   'open data file
  FOR Cnt = 1 TO NumOfRateRecs
    FGetRTA Handle, UBRateRecs(1), CLNG(Cnt), UBRateRecLen
    Choice$(Cnt, 0) = UBRateRecs(1).RATECODE + "  " + UBRateRecs(1).RATEDESC
  NEXT
  FClose Handle
  
  Choice$(1, 3) = "Active"
  Choice$(2, 3) = "Inactive"
  
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
  
RETURN
  
InitCustDefaults:
  Form$(3, 0) = "A"
  Form$(4, 0) = DATE$
  Form$(10, 0) = UBSetUpRec(1).DEFCITY
  Form$(11, 0) = UBSetUpRec(1).DEFSTATE
  Form$(12, 0) = UBSetUpRec(1).ZIPCODE
  Form$(19, 0) = "1"
  
  Form$(25, 0) = "N"
  Form$(26, 0) = "Y"
  Form$(27, 0) = "Y"
  Form$(28, 0) = "N"
  Form$(29, 0) = "N"
  Form$(30, 0) = "N"
  Form$(31, 0) = "0"
  FOR ZZCnt = 97 TO 106
    Form$(ZZCnt, 0) = "0"
  NEXT
  FOR ZZCnt = 108 TO 186 STEP 13
    Form$(ZZCnt, 0) = "1"
    Form$(ZZCnt + 3, 0) = "1"
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
  
  UBFILE = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFILE LEN = UBCustRecLen
  GET UBFILE, RecNo&, UBCustRec(1)
  CLOSE
  
  OldBook$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
  
  BCopy VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), TFormLen, 0
  
  IF UBCustRec(1).STATUS = "F" THEN
    FinialFlag = True
  END IF
  
  DepAmt# = UBCustRec(1).DepositAmt
  
  ERASE UBCustRec
  
RETURN
  
SaveCustDeposit:
  REDIM UBTransRec(1)  AS UBTransRecType
  REDIM UBCustRec(1)   AS NewUBCustRecType
  
  UBCustRecLen = LEN(UBCustRec(1))
  UBTransRecLen = LEN(UBTransRec(1))
  
  FOpenS "UBCUST.DAT", CHandle
  FOpenS "UBTRANS.DAT", Thandle
  
  FGetRTA CHandle, UBCustRec(1), RecNo&, UBCustRecLen
  UBTransRec(1).TransDate = Date2Num(Form$(4, 0))
  
  UBTransRec(1).TransType = TranDepositPayment
  UBTransRec(1).TransDesc = "DEPOSIT PAYMENT"
  UBTransRec(1).CustLocation = RecNo&
  UBTransRec(1).OperatorNumber = OperNum
  UBTransRec(1).CustAcctNo = RecNo&
  UBTransRec(1).CustStatus = Form$(2, 0)
  UBTransRec(1).TransAmt = DepAmt#
  'UBTransRec(1).RevAmt(1) = UBTransRec(1).TransAmt
  UBCustRec(1).DepositAmt = UBTransRec(1).TransAmt
  UBTransRec(1).PayTypeCode = -1
  UBTransRec(1).PrevTrans = UBCustRec(1).LastTrans
  NextTransRec& = (FLof(Thandle) \ UBTransRecLen) + 1
  
  UBCustRec(1).LastTrans = NextTransRec&
  FPutRTA Thandle, UBTransRec(1), NextTransRec&, UBTransRecLen
  FPutRTA CHandle, UBCustRec(1), RecNo&, UBCustRecLen
  
  FClose CHandle
  FClose Thandle
  
  RETURN
  
END SUB

SUB PrintCustList
  
  MaxLines = 59
  PageNo = 0
  Dash80$ = STRING$(80, "-")
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize("UBCUSTNM.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTNM.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBCULIST.RPT" FOR OUTPUT AS UBRpt
  
  BlockClear
  ShowProcessingScrn "Customer Listing. (Name Order)"
  
  GOSUB DoCustRptHeader
  
  FOR Cnt = 1 TO IdxNumOfRecs
    GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)

    'IF UBCustRec(1).MessageRec = 27 OR UBCustRec(1).MessageRec = 28 THEN
    '  STOP
    'END IF
    IF NOT UBCustRec(1).DelFlag THEN

      IF LineCnt > MaxLines THEN
        PRINT #UBRpt, CHR$(12)
        GOSUB DoCustRptHeader
      END IF
      PRINT #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; " "; FUsing$(STR$(IdxBuff(Cnt).RecNum), "#####"); "  "; LEFT$(UBCustRec(1).CUSTNAME, 25); "  "; LEFT$(UBCustRec(1).SERVADDR, 30); " "; UBCustRec(1).STATUS;
      CustCnt = CustCnt + 1
      SELECT CASE UBCustRec(1).STATUS
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
      IF UBCustRec(1).DelFlag THEN
        PRINT #UBRpt, "D"
      ELSE
        PRINT #UBRpt,
      END IF
      LineCnt = LineCnt + 1
    END IF
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp Cnt, IdxNumOfRecs
  NEXT

  GOSUB DoCustRptTotals
  
  CLOSE UBCust, UBLoca, UBRpt
  ERASE IdxBuff, UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Customer Listing. (Name Order)", "UBCULIST.RPT", 1, RetCode, EntryPoint
  END IF
  
  EXIT SUB
  
  
DoCustRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Customer Listing Report      "; "Date: "; DATE$; TAB(70); "Page: "; PageNo
  PRINT #UBRpt, "           Acct "; TAB(72); "Customer"
  PRINT #UBRpt, "Location    No.   Customer Name             Service Address             Status"
  PRINT #UBRpt, Dash80$
  LineCnt = 4
RETURN
  
DoCustRptTotals:
  PageNo = PageNo + 1
  PRINT #UBRpt,
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Customer Summary"
  PRINT #UBRpt,
  PRINT #UBRpt, USING "  Active: ####"; Active
  PRINT #UBRpt, USING "   Final: ####"; Final
  PRINT #UBRpt, USING "Inactive: ####"; InActive
  PRINT #UBRpt, USING " Balance: ####"; Balance
  PRINT #UBRpt, USING " Unknown: ####"; UnKnown
  PRINT #UBRpt,
  PRINT #UBRpt, USING "   TOTAL: ####"; CustCnt
  PRINT #UBRpt, CHR$(12)
RETURN
  
END SUB

SUB PrintLocaList
  
  MaxLines = 59
  PageNo = 0
  Dash80$ = STRING$(80, "-")
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBCustBlank(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBLOLIST.RPT" FOR OUTPUT AS UBRpt
  
  BlockClear
  ShowProcessingScrn "Customer Listing. (Location Order)"
  GOSUB DoLocaRptHeader
  
  FOR Cnt = 1 TO IdxNumOfRecs
    GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB DoLocaRptHeader
    END IF
    IF NOT UBCustRec(1).DelFlag THEN
      PRINT #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; " "; FUsing$(STR$(IdxBuff(Cnt).RecNum), "#####"); "  "; LEFT$(UBCustRec(1).CUSTNAME, 25); "  "; LEFT$(UBCustRec(1).SERVADDR, 30); " "; UBCustRec(1).STATUS
      CustCnt = CustCnt + 1
      SELECT CASE UBCustRec(1).STATUS
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
    
      IF AskAbandonPrint% THEN
        AbortFlag = True
        EXIT FOR
      END IF
    END IF
    ShowPctComp Cnt, IdxNumOfRecs
  NEXT
  
  GOSUB DoLocaRptTotals
  
  CLOSE UBCust, UBLoca, UBRpt
  
  ERASE IdxBuff, UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Customer Listing. (Location Order)", "UBLOLIST.RPT", 1, RetCode, EntryPoint
  END IF
  
  'KillFile "UBLOLIST.RPT"
  
EXIT SUB
  
DoLocaRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Customer Listing Report      "; "Date: "; DATE$; TAB(70); "Page: "; PageNo
  PRINT #UBRpt, "           Acct "; TAB(72); "Customer"
  PRINT #UBRpt, "Location    No.   Customer Name             Service Address             Status"
  PRINT #UBRpt, Dash80$
  LineCnt = 4
RETURN
  
DoLocaRptTotals:
  PageNo = PageNo + 1
  PRINT #UBRpt,
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Customer Summary"
  PRINT #UBRpt,
  PRINT #UBRpt, USING "  Active: ####"; Active
  PRINT #UBRpt, USING "   Final: ####"; Final
  PRINT #UBRpt, USING "Inactive: ####"; InActive
  PRINT #UBRpt, USING " Balance: ####"; Balance
  PRINT #UBRpt, USING " Unknown: ####"; UnKnown
  PRINT #UBRpt,
  PRINT #UBRpt, USING "   TOTAL: ####"; CustCnt
  PRINT #UBRpt, CHR$(12)
RETURN
  
  
END SUB

SUB SaveNewCustRec (Form$(), RecNo&)
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  
  FormLen = LEN(Form$(0, 0))    'Length of Cust Form$(0,0) Element
  
  UBCustRecLen = LEN(UBCustRec(1))              'Length of Cust Record Structure
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1)), FormLen, 0
  'Copy Cust info from Form$(0,0) to the Cust Rec Structure
  'UBCustRec(1).Status = "I"
  
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
  
  UBFILE = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFILE LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFILE) \ UBCustRecLen
  NextRec& = NumOfRecs& + 1
  
  '***REM THIS in normal operation
  '  UBCustRec(1).CustFrstLocRec = NextRec&
  '  UBCustRec(1).CustLastLocRec = NextRec&
  '***REM THIS
  
  UBCustRec(1).MessageRec = 0   ' Set Pointers To Zero
  UBCustRec(1).LastTrans = 0
  UBCustRec(1).CurrBalance = 0
  UBCustRec(1).PrevBalance = 0
  'UBCustRec(1).DepositAmt = 0
  UBCustRec(1).WOLastTrans = 0
  
  PUT UBFILE, NextRec&, UBCustRec(1)
  CLOSE UBFILE
  
  'Update Cust Name Search index -\-\-\-\-\-\-\-
  
  IdxRecLen = 4 'we are using a integer
  
  IdxFileSize& = FileSize("UBCUSTNM.IDX")
  
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  NextFreeRec = IdxNumOfRecs + 1
  
  REDIM IdxBuff(1 TO IdxNumOfRecs + 1) AS UBCustIndexRecType
  'dim array to num of recs +1
  
  REDIM TempUBCustRec(1) AS NewUBCustRecType
  
  IF IdxNumOfRecs > 0 THEN      'if this is not the first
    FGetAH "UBCUSTNM.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs  'load it
    FOpenS "UBCUST.DAT", C1Handle               'open data file
    LowOffSet = 0               '
    TopOffSet = IdxNumOfRecs + 1
    RecOffSet = IdxNumOfRecs \ 2
    
    IF RecOffSet < 1 THEN RecOffSet = 1
    
    DO
      
      FGetRTA C1Handle, TempUBCustRec(1), CLNG(IdxBuff(RecOffSet).RecNum), UBCustRecLen
      'get the rec from the data file based off the index rec num
      
      SELECT CASE Compare3%(VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1).SEARCH), VARSEG(TempUBCustRec(1)), VARPTR(TempUBCustRec(1).SEARCH), 10)
        'compare the rec we just read to new cust we just added
      CASE -1   'New search name comes before this one
        TopOffSet = RecOffSet
        RecOffSet = RecOffSet \ 2
        IF RecOffSet < LowOffSet THEN
          RecOffSet = LowOffSet + ((TopOffSet - LowOffSet) \ 2)
        END IF
      CASE 1    'New search name comes after this one
        LowOffSet = RecOffSet
        RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
        
      CASE 0    'New search name is the same as this one
        SELECT CASE Compare3%(VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1).CUSTNAME), VARSEG(TempUBCustRec(1)), VARPTR(TempUBCustRec(1).CUSTNAME), 35)
          'so we will examine the whole name
        CASE -1 'Lower
          TopOffSet = RecOffSet
          RecOffSet = RecOffSet \ 2
          IF RecOffSet < LowOffSet THEN
            RecOffSet = LowOffSet + ((TopOffSet - LowOffSet) \ 2)
          END IF
        CASE 1  'Higher
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
    
    IdxBuff(InsertRecAt).RecNum = NextFreeRec
    
    FPutAH "UBCUSTNM.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs + 1
    FClose C1Handle
    
  ELSE
    IdxBuff(1).RecNum = NextFreeRec
    FPutAH "UBCUSTNM.IDX", IdxBuff(1), IdxRecLen, 1
  END IF
  
  ERASE IdxBuff
  
  'Done with name index
  RecNo& = NextRec&
  
  '**********************
  'Update Location index -\-\-\-\-\-\-\-
  
  IdxRecLen = 4 'we are using a integer
  
  IdxFileSize& = FileSize("UBCUSTBK.IDX")
  
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  NextFreeRec = IdxNumOfRecs + 1
  
  REDIM IdxBuff(1 TO IdxNumOfRecs + 1) AS UBCustIndexRecType
  '  'dim array to num of recs +1
  
  REDIM TempUBCustRec(1) AS NewUBCustRecType
  
  IF IdxNumOfRecs > 0 THEN      'if this is not the first
    FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs  'load it
    FOpenS "UBCUST.DAT", C1Handle               'open data file
    
    LowOffSet = 0               '
    TopOffSet = IdxNumOfRecs + 1
    RecOffSet = IdxNumOfRecs \ 2
    
    IF RecOffSet < 1 THEN RecOffSet = 1
    
    DO
      
      FGetRTA C1Handle, TempUBCustRec(1), IdxBuff(RecOffSet).RecNum, UBLocaRecLen
      'get the rec from the data file based off the index rec num
      
      SELECT CASE Compare3%(VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1).Book), VARSEG(TempUBCustRec(1)), VARPTR(TempUBCustRec(1).Book), 8)
        'compare the rec we just read to new cust we just added
      CASE -1   'New Route is before this one
        TopOffSet = RecOffSet
        RecOffSet = RecOffSet \ 2
        IF RecOffSet < LowOffSet THEN
          RecOffSet = LowOffSet + ((TopOffSet - LowOffSet) \ 2)
        END IF
      CASE 1    'New Route is after this one
        LowOffSet = RecOffSet
        RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
        
      CASE 0    'This Should never occure
        'STOP
        LowOffSet = RecOffSet
        RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
      END SELECT
    LOOP WHILE TopOffSet - LowOffSet > 1
    
    InsertRecAt = TopOffSet
    
    AmtToMove = (IdxNumOfRecs + 1) - TopOffSet
    
    IF AmtToMove THEN
      InsertT IdxBuff(InsertRecAt), IdxRecLen, AmtToMove
    END IF
    
    IdxBuff(InsertRecAt).RecNum = NextFreeRec
    
    FPutAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs + 1
    
  ELSE
    IdxBuff(1).RecNum = NextFreeRec
    FPutAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, 1
  END IF
  
  FClose C1Handle
  
  '**********************
  
  'STOP
  REDIM UBBookSeq(1) AS BookSeqRecType
  
  TempBook$ = UBCustRec(1).Book + UBCustRec(1).SEQNUMB
  
  UBBookSeq(1).BookSeq = QPValL(TempBook$)
  
  BookHand = FREEFILE
  OPEN "UBOOKSEQ.DAT" FOR RANDOM SHARED AS BookHand LEN = 4
  NextBookRec = (LOF(BookHand) / 4) + 1
  PUT BookHand, NextBookRec, UBBookSeq(1)
  CLOSE BookHand
  
  ThisBook$ = UBCustRec(1).Book
  REDIM BookList(1)  AS STRING * 2
  
  BookList = FREEFILE
  OPEN "UBOOKLST.DAT" FOR RANDOM SHARED AS BookList LEN = 2
  NumBookRec = LOF(BookHand) / 2
  IF NumBookRec > 0 THEN
    FOR Cnt = 1 TO NumBookRec
      GET BookList, Cnt, BookList(1)
      IF ThisBook$ = BookList(1) THEN
        GotBook = True
        EXIT FOR
      END IF
    NEXT
    IF NOT GotBook THEN
      BookList(1) = ThisBook$
      PUT BookList, NumBookRec + 1, BookList(1)
    END IF
  ELSE
    BookList(1) = ThisBook$
    PUT BookList, NumBookRec + 1, BookList(1)
  END IF
  CLOSE BookList
  
  '******************
  
  ERASE UBCustRec, TempUBCustRec, IdxBuff, BookList
  
END SUB

SUB SaveOldCustRec (Form$(), RecNo&)
  
  CursorOff
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  DisplayUBScrn "UPDATDSK"

  'REDIM UBCustTest(1) AS UBCUSTTESTTYPE
  REDIM UBCustRec(1 TO 2) AS NewUBCustRecType

  'UBCustTestLen = LEN(UBCustTest(1))
  UBCustRecLen = LEN(UBCustRec(1))

  UBFILE = FREEFILE
  
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFILE LEN = UBCustRecLen

  GET UBFILE, RecNo&, UBCustRec(1)
  GET UBFILE, RecNo&, UBCustRec(2)
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1)), LEN(Form$(0, 0)), 0

'060597 This solves date & avguse errors
  FOR MTRCnt = 1 TO 7
    UBCustRec(1).LocMeters(MTRCnt).PastDate = UBCustRec(2).LocMeters(MTRCnt).PastDate
    UBCustRec(1).LocMeters(MTRCnt).ReadFlag = UBCustRec(2).LocMeters(MTRCnt).ReadFlag
    UBCustRec(1).LocMeters(MTRCnt).AvgUse = UBCustRec(2).LocMeters(MTRCnt).AvgUse
    UBCustRec(1).LocMeters(MTRCnt).UseCnt = UBCustRec(2).LocMeters(MTRCnt).UseCnt
  NEXT

  PUT UBFILE, RecNo&, UBCustRec(1)
  
  CLOSE UBFILE
  
  '01-20-97 Added skip reindexing if name not changed
  IF UBCustRec(1).SEARCH <> UBCustRec(2).SEARCH OR UBCustRec(1).CUSTNAME <> UBCustRec(2).CUSTNAME THEN
    RestScrn TempScrn()
    RptTitle$ = "Reading Customer Names."
    TitleRow = 9
    TitleCol = 40 - (LEN(RptTitle$) \ 2) + 1
    DisplayUBScrn "PRORPT"
    'QPrintRC RptTitle$, 10, 29, -1
    QPrintRC RptTitle$, TitleRow, TitleCol, 126
    QPrintRC "Processing:    % Completed.", 13, 28, Cnf.HiLite
    
    UBFILE = FREEFILE
    OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFILE LEN = UBCustRecLen
    NumOfRecs& = LOF(UBFILE) \ UBCustRecLen
    
    REDIM IdxBuff(1 TO NumOfRecs&) AS UBCustReIndexRecType
    
    FOR Cnt = 1 TO NumOfRecs&
      GET UBFILE, Cnt, UBCustRec(1)
      IdxBuff(Cnt).SearchName = UBCustRec(1).SEARCH
      IdxBuff(Cnt).First = QPTrim$(UBCustRec(1).CUSTNAME)
      IdxBuff(Cnt).RecNum = Cnt
      ShowPctComp Cnt, NumOfRecs&
    NEXT
    
    CLOSE UBFILE
    
    '  SortT IdxBuff(1), CINT(NumOfRecs&), 0, 16, 10, 2
    SortT IdxBuff(1), CINT(NumOfRecs&), 0, 16, 0, 10
    'Array(1), NumElem, Dir, StructSize, MemOff, MemSize
    
    KillFile "UBCUSTNM.IDX"
    RptTitle$ = "Writing Customer Index."
    QPrintRC RptTitle$, TitleRow, TitleCol, 126
    QPrintRC "Processing:    % Completed.", 13, 28, Cnf.HiLite
    
    UBFILE = FREEFILE
    OPEN "UBCUSTNM.IDX" FOR RANDOM SHARED AS UBFILE LEN = 4
    
    FOR Cnt = 1 TO NumOfRecs&
      PUT UBFILE, Cnt, IdxBuff(Cnt).RecNum
      ShowPctComp Cnt, NumOfRecs&
    NEXT
    CLOSE UBFILE
    
    ERASE IdxBuff
  END IF
  '-=-=-=-=-=-=-=-=-=-=-
  
  IF (UBCustRec(1).Book <> UBCustRec(2).Book) OR (UBCustRec(1).SEQNUMB <> UBCustRec(2).SEQNUMB) THEN
    
    UBFILE = FREEFILE
    OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFILE LEN = UBCustRecLen
    NumOfRecs& = LOF(UBFILE) \ UBCustRecLen
    
    REDIM IdxBuffL(1 TO NumOfRecs&) AS UBLocaReIndexRecType
    
    FOR Cnt = 1 TO NumOfRecs&
      GET UBFILE, Cnt, UBCustRec(1)
      IdxBuffL(Cnt).Book = UBCustRec(1).Book
      IdxBuffL(Cnt).SEQNUMB = UBCustRec(1).SEQNUMB
      IdxBuffL(Cnt).RecNum = Cnt
    NEXT
    
    CLOSE UBFILE
    
    SortT IdxBuffL(1), CINT(NumOfRecs&), 0, 16, 0, 8
    'Array(1), NumElem, Dir, StructSize, MemOff, MemSize
    
    KillFile "UBCUSTBK.IDX"
    
    UBFILE = FREEFILE
    OPEN "UBCUSTBK.IDX" FOR RANDOM SHARED AS UBFILE LEN = 4
    
    FOR Cnt = 1 TO NumOfRecs&
      PUT UBFILE, Cnt, IdxBuffL(Cnt).RecNum
    NEXT
    CLOSE UBFILE
    
    REDIM UBBookSeq(1) AS BookSeqRecType
    BookSeqLen = LEN(UBBookSeq(1))
    
    OBook& = QPValL(UBCustRec(2).Book + UBCustRec(2).SEQNUMB)
    TBook& = QPValL(UBCustRec(1).Book + UBCustRec(1).SEQNUMB)
    
    BookHand = FREEFILE
    OPEN "UBOOKSEQ.DAT" FOR RANDOM SHARED AS BookHand LEN = 4
    NumBookRec = LOF(BookHand) / BookSeqLen
    
    FOR Cnt = 1 TO NumBookRec
      GET BookHand, Cnt, UBBookSeq(1)
      IF UBBookSeq(1).BookSeq = OBook& THEN
        UBBookSeq(1).BookSeq = TBook&
        PUT BookHand, Cnt, UBBookSeq(1)
        EXIT FOR
      END IF
    NEXT
    CLOSE BookHand
    
    '-=-=-=-=-=-=-=-=-=-=-
  END IF
  'GOTO SkipCustIndex
  '
  'SkipCustIndex:
  RestScrn TempScrn()
END SUB

SUB UBCustomerMenu
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 7)
  
  MChoice$(1) = " Add a New Customer     "
  MChoice$(2) = " Edit Existing Customer "
  MChoice$(3) = " Set a Customer to Final"
  MChoice$(4) = " Delete Customer Account"
  MChoice$(5) = " Quick Customer Listing by Name"
  MChoice$(6) = " Quick Customer Listing by Location"
  
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
  Col = ((80 - MaxLen) \ 2) - 1
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    
    TitleBox 3, Col, MaxLen + 3, "Utility Customer Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      REDIM CRec(1) AS NewUBCustRecType
      'IF FileSize("UBCUST.DAT") \ LEN(CRec(1)) < 3 THEN
      AddEditCustomer 0, False, False
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
        AddEditCustomer RecNo&, False, True
      END IF
    CASE 3
      SetCustFinal
    CASE 4
      DeleteCust
    CASE 5
      PrintCustList
    CASE 6
      PrintLocaList
    END SELECT
  LOOP
  
  ' RUN "ubmenu"
  
END SUB

