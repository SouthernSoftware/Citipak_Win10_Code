DECLARE SUB UtilZapAdjustments ()
  DEFINT A-Z
DECLARE FUNCTION GetCustMeterType% (UBLocatRec() AS ANY, ThisMeter%)
DECLARE SUB UtilAdjustments ()
DECLARE SUB RePrintUtilBills ()
DECLARE SUB PostBillTrans ()
DECLARE FUNCTION GetAdjFactor# ()
DECLARE SUB ClearScrn ()
DECLARE SUB UBBillMenu ()
DECLARE SUB BillPrintMenu ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE SUB PrintUtilBills ()
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE FUNCTION AskAbandonPrint% ()
  
DECLARE FUNCTION aFindRateTbl% (RATECODE$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetRevCharge# (RateTbl AS ANY, TMeterConsp&, MeterMulti&)
DECLARE SUB MakeBillFile (AbortFlag%, FuelAdjAmt#)
DECLARE FUNCTION Round# (N#)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB MakeZipCodeIndex ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB MakePostalIndex (IndexText$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB MakeTempIndex (Indextype%)
DECLARE SUB PreBillReport ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB ManualMeterReading ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB EnterMeterReadingsMenu ()
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION Exist% (FileName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION AskSavePayment
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
  
'$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB SortStr2 (BYVAL Address%, Size%, Direction%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION QPValL& (LongNum$)
DECLARE SUB MPaintBox (TRow%, LCol%, BRow%, RCol%, Kolor%)
DECLARE SUB StuffBuf (Key2Stuff$)
DECLARE SUB ShowCustHistory (CustAcct&)
DECLARE SUB CursorOff ()
DECLARE SUB WaitForAction ()
DECLARE SUB BlockClear ()
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB LookUp (CustAcct&, Text$, Parm1, Parm2, ActiveOnly%)
DECLARE FUNCTION PromptSaveData ()
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  
  '$INCLUDE: 'PageInfo.bi'
  '$INCLUDE: 'newcust.bi'
  ''$INCLUDE: 'ubcust.bi'
  ''$INCLUDE: 'ubloc.bi'
  '$INCLUDE: 'ubrate.BI'
  ''$INCLUDE: 'ubsetbal.BI'
  
  CONST False = 0, True = NOT False
  
  CLS
  UtilZapAdjustments
  CLS
'  RUN "UBMENU"

SUB UtilZapAdjustments
  
  SHARED Choice$()
  
  'AdjDate = Date2Num(DATE$)
  
  '**********************************************************
  REDIM TempScrn(0)
  REDIM UBCustRec(1)   AS NewUBCustRecType
  
  UBCustRecLen = LEN(UBCustRec(1))
  
  NumOfRevs = MaxRevsCnt
  
  REDIM RevText$(1 TO MaxRevsCnt)
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  
  TownName$ = UBSetUpRec(1).UTILNAME
  
  FOR RevCnt = 1 TO MaxRevsCnt
    RevText$(RevCnt) = LEFT$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME), 14)
    IF LEN(RevText$(RevCnt)) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    END IF
  NEXT
  
  IF NumOfRevs < MaxRevsCnt THEN
    REDIM PRESERVE RevText$(1 TO NumOfRevs)
  END IF
  
  LibName$ = "UBZAP"
  ScrnName$ = "UBZAP"
  
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  Frm(1).StayOnField = True
  Frm(1).PageDisable = True
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  CustAcctFld = FldNum%("CUSTACCT", Fld())
  TotAmtFld = FldNum%("TADJAMT", Fld())
  PrevAmtFld = FldNum%("PREVIOUS", Fld())
  DepositFld = FldNum%("DEPOSIT", Fld())

  REDIM CashFlds(1 TO MaxRevsCnt)
  
  FOR Cnt = 1 TO MaxRevsCnt
    CashFlds(Cnt) = FldNum%("AMTADJ" + LTRIM$(STR$(Cnt)), Fld())
  NEXT
  
'  FOR Cnt = NumOfRevs + 1 TO MaxRevsCnt
'    Fld(CashFlds(Cnt)).Protected = True
'  NEXT
  
  '--define the multi-choice fields
  
  'GOSUB ClearForm
  
  Action = 1
  FirstTime = True
  
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF Frm(1).Edited AND Frm(1).PrevFld <> Frm(1).FldNo THEN
      BeenEditedFlag = True     'if the form has been edited
    END IF      'set the edited flag
    
    IF FirstTime THEN
      FirstTime = False         'if this is the first time
      GOSUB ShowRevSources      '
      Action = 1
      CalcFields 0, CashFlds(1), Form$(), Fld()
    END IF
    
    IF Frm(1).FldNo > CustAcctFld AND Frm(1).PrevFld = CustAcctFld THEN
      CustAcct& = QPValL(Form$(CustAcctFld, 0))
      GOSUB GetCustInfo
    ELSEIF Frm(1).FldNo = CustAcctFld AND Frm(1).PrevFld <> CustAcctFld THEN
      MPaintBox 22, 37, 22, 41, 112
      MPaintBox 22, 35, 22, 36, 126
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE EscKey
      IF BeenEditedFlag THEN
        'SaveFlag = PromptSaveData
        'SELECT CASE SaveFlag
        'CASE True               'user wants to save
        '  EscapeFlag = True
        '  StuffBuf CHR$(0) + CHR$(ABS(F10Key))
        'CASE False              'user wants to abandon
          ExitFlag = True
        'CASE ELSE               'continue editing
        'END SELECT
        'Action = 1
      ELSE
        ExitFlag = True
      END IF
      
    CASE F4KEY  'Customer History
      IF CustAcct& > 0 THEN
        ShowCustHistory CustAcct&
        Action = 2
      END IF
      
    CASE F7KEY  'Lookup Customer
      IF Frm(1).FldNo = 1 THEN  'if user is on the Customer field
        SaveScrn TempScrn()     'and F7key then do lookup routine
        MPaintBox 4, 5, 22, 75, 8
        LookUp CustAcct&, "Payment", 2, False, False
        RestScrn TempScrn()
        IF CustAcct& > 0 THEN   'if this is a valid customer
          GOSUB GetCustInfo     'go get customer info
          Frm(1).FldNo = 4
          Action = 1
        ELSE
          GOSUB ClearForm
          Frm(1).FldNo = 1
          Action = 1
        END IF
      END IF
      
    CASE F10Key 'Save
      GOSUB CheckAdjustInfo
      IF AdjustOKFlag THEN
        OFldNo = Frm(1).FldNo
        Ok = MsgBox%("UB.QSL", "OK2ADJ")        'Ask if they are sure
        SELECT CASE Ok
        CASE 1  'Yes do the adjustment
          GOSUB SaveTransaction 'do the save routine
          GOSUB ClearForm
          Frm(1).FldNo = 1
          Action = 1
        CASE ELSE               'Nope continue editing
          Frm(1).FldNo = OFldNo - 1
          Action = 1
        END SELECT
        IF EscapeFlag THEN
          ExitFlag = True
        END IF
      ELSE
        Action = 2
      END IF
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 5 TO 13            'f4 History
          PressButton F4KEY, 22, 5, 13
        CASE 34 TO 42           'f7 Look-Up
          PressButton F7KEY, 22, 34, 42
        CASE 54 TO 63           'f10 Save
          PressButton F10Key, 22, 54, 63
        CASE 65 TO 75           '--cancel button
          PressButton EscKey, 22, 65, 75
        END SELECT
      END SELECT                'row
    END IF
  LOOP UNTIL ExitFlag
  
  ERASE TempScrn, UBCustRec, RevText$, UBSetUpRec
  
  HideCursor
  
ExitAdjustments:
  
  EXIT SUB
  
GetCustInfo:
  NumOfCustRecs& = FileSize("UBCUST.DAT") \ UBCustRecLen
  IF CustAcct& > NumOfCustRecs& OR CustAcct& <= 0 THEN
    CustAcct& = 0
    Ok = MsgBox%("UB.QSL", "BADACCTN")
    Frm(1).FldNo = 1
    GOSUB ClearForm
    GOTO SkipCustInfo
  END IF
  
  GOSUB ClearForm
  
  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = UBCustRecLen
  GET CustFile, CustAcct&, UBCustRec(1)
  CLOSE CustFile
  FOR Cnt = 1 TO 15
    LSET Form$(CashFlds(Cnt), 0) = QPTrim$(STR$(UBCustRec(1).CurrRevAmts(Cnt)))
  NEXT

  LSET Form$(PrevAmtFld, 0) = STR$(UBCustRec(1).PrevBalance)
  LSET Form$(TotAmtFld, 0) = STR$(UBCustRec(1).CurrBalance)
  LSET Form$(CustAcctFld, 0) = STR$(CustAcct&)
  LSET Form$(DepositFld, 0) = STR$(UBCustRec(1).DepositAmt)
  LSET Form$(2, 0) = UBCustRec(1).CUSTNAME
  LSET Form$(3, 0) = UBCustRec(1).ADDR1

  SaveField 0, Form$(), Fld(), BadField
  CustAcct& = QPValL(Form$(CustAcctFld, 0))
  
  MPaintBox 22, 35, 22, 41, 115
  FirstTime = True
  
SkipCustInfo:
  Action = 1
  RETURN
  
ShowRevSources:
  FOR RevCnt = 1 TO NumOfRevs
    QPrintRC RevText$(RevCnt), RevCnt + 3, 42, -1
  NEXT
  RETURN
  
ClearForm:
  TempDate = Date2Num(Form$(AdjDateFld, 0))
  
  IF TempDate > 0 AND TempDate <> AdjDate THEN
    AdjDate = TempDate
  END IF
  
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""       '--Clear all fields
  NEXT
  
  LSET Form$(AdjDateFld, 0) = Num2Date(AdjDate)
  
  BeenEditedFlag = False        'clear the edited flag
  FOR Cnt = 1 TO 7
    QPrintRC "            ", Cnt + 13, 8, -1
    QPrintRC " ", Cnt + 13, 24, -1
  NEXT
  
  RETURN
  
SaveTransaction:
  
  SaveScrn TempScrn()
  DisplayUBScrn "UPDATDSK"

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  GET UBCust, CustAcct&, UBCustRec(1)

  FOR RevCnt = 1 TO MaxRevsCnt
    UBCustRec(1).CurrRevAmts(RevCnt) = Round#(Value#(Form$(CashFlds(RevCnt), 0), ECode))
  NEXT

  UBCustRec(1).CurrBalance = Value#(Form$(TotAmtFld, 0), ECode)
  UBCustRec(1).PrevBalance = Value#(Form$(PrevAmtFld, 0), ECode)
  UBCustRec(1).DepositAmt = Value#(Form$(DepositFld, 0), ECode)

  PUT UBCust, CustAcct&, UBCustRec(1)
  CLOSE UBCust
  
  RestScrn TempScrn()
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  RestScrn TempScrn()
  FirstTime = True
  RETURN
  
CheckAdjustInfo:
  AdjustOKFlag = True
  RETURN

  CurrFld = Frm(1).FldNo
  
  AdjustOKFlag = True
  ConsAdjFlag = False
  BillAdjFlag = False
  
  TestDate = Date2Num(Form$(AdjDateFld, 0))
  IF TestDate < 0 THEN
    Ok = MsgBox%("UB.QSL", "UBBADTDT")
    Action = 2
    AdjustOKFlag = False
    Frm(1).FldNo = 4
    GOTO BadAdjustment
  END IF
  
  IF LEN(QPTrim$(Form$(AdjTypeFld, 0))) = 0 THEN
    Ok = MsgBox%("UB.QSL", "UBBADTYP")
    Action = 2
    AdjustOKFlag = False
    Frm(1).FldNo = 5
    GOTO BadAdjustment
  END IF
  
  FOR Cnt = 1 TO 7
    IF LEN(QPTrim$(Form$(6 + Cnt, 0))) > 0 THEN
      ConsAdjFlag = True
      EXIT FOR
    END IF
  NEXT
  
  FOR Cnt = 1 TO MaxRevsCnt
    IF LEN(QPTrim$(Form$(13 + Cnt, 0))) > 0 THEN
      BillAdjFlag = True
      EXIT FOR
    END IF
  NEXT
  
  IF (NOT ConsAdjFlag AND NOT BillAdjFlag) THEN
    Ok = MsgBox%("UB.QSL", "UBBADADJ")
    Action = 2
    AdjustOKFlag = False
    Frm(1).FldNo = CurrFld - 1
    GOTO BadAdjustment
  END IF
  
  
BadAdjustment:
  RETURN
  
END SUB

