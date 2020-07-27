DECLARE SUB CursorOff ()
  DEFINT A-Z
DECLARE SUB PenaltyProcess ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION FindRateTbl% (RATECODE$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION GetRevCharge# (RateTbl AS ANY, TMeterConsp&)
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
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FLof& (FileHandle%)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB ClearScrn ()
  '$INCLUDE: 'DefCnf.BI'
  
DECLARE FUNCTION WEnvTest% ()
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
DECLARE FUNCTION QPValI% (Number$)
DECLARE SUB WaitForAction ()
DECLARE SUB BlockClear ()
  
DECLARE SUB SaveScrn (TempScrn())
DECLARE SUB RestScrn (TempScrn())
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  
  '$INCLUDE: 'PageInfo.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'ubsetbal.BI'
  
  CONST False = 0, True = NOT False
  
  PenaltyProcess
  
  IF WEnvTest AND INSTR(COMMAND$, "TEST") > 0 THEN
    HideCursor
    '   ClearScrn
    END
  ELSE
    RUN "UBBILLIN"
  END IF
  
SUB PenaltyProcess
  
  REDIM TempScrn(0)
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBTranRec(1 TO 2) AS UBTransRecType
  
  UBCustRecLen = LEN(UBCustRec(1))
  UBTranRecLen = LEN(UBTranRec(1))
  
  LibName$ = "UB"
  ScrnName$ = "UBPENLIL"
  
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
  
  '--Set choices
  
  '--Set screen number to one and display screen
  DisplayUBScrn ScrnName$
  Action = 1
  FirstTime = True
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      Action = 1
      LSET Form$(1, 0) = DATE$
    END IF
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB CheckPenaltyFlds
    CASE EscKey
      ExitFlag = True
    END SELECT
    
  LOOP UNTIL ExitFlag OR OkFlag
  
  IF ExitFlag THEN GOTO ExitPenalty
  
  BlockClear
  ShowProcessingScrn "Applying Penalty Charges"
  
  IF MinBalance# < 0 THEN
    MinBalance# = 0
  END IF
  
  PctAmt# = 1.1111
  
  RevSource = 1
  
  TransDesc$ = QPTrim$(Form$(2, 0))
  
  FOpenS "UBCUST.DAT", CHandle  'open data file
  
  NumCustRecs& = FLof(CHandle) / UBCustRecLen
  
  FOR CCnt& = 1 TO NumCustRecs&
    FGetRTA CHandle, UBCustRec(1), CCnt&, UBCustRecLen
    IF UBCustRec(1).STATUS = "A" THEN           'if they are not inactive
      OkFlag = False
      IF UBCustRec(1).CurrBalance <= 0 THEN
        OkFlag = False
      ELSE
        RateC$ = UBCustRec(1).Serv(1).RATECODE
        IF INSTR(RateC$, "WIN") > 0 OR INSTR(RateC$, "WOUT") > 0 THEN
          OkFlag = True
        ELSE
          OkFlag = False
          'add check for rate code here
        END IF
      END IF
      IF OkFlag THEN
        CustPenalty# = Round#((UBCustRec(1).CurrRevAmts(1) * PctAmt#) - UBCustRec(1).CurrRevAmts(1))
        UBCustRec(1).CurrBalance = Round#(UBCustRec(1).CurrBalance + CustPenalty#)
        UBCustRec(1).CurrRevAmts(1) = Round#(UBCustRec(1).CurrRevAmts(1) + CustPenalty#)
        GOSUB MakeTransaction
      END IF
    END IF
    SkipEM:
    ShowPctComp CCnt&, NumCustRecs&
  NEXT
  FClose CHandle
  
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  ExitPenalty:
  
  ERASE TempScrn, UBCustRec, UBTranRec
  ERASE Frm, Form$, Fld
  
  EXIT SUB
  
  MakeTransaction:
  
  LSET UBTranRec(1) = UBTranRec(2)              'Transrec 2 is blank
  UBTranRec(1).TransAmt = CustPenalty#
  UBTranRec(1).RevAmt(1) = CustPenalty#
  UBTranRec(1).TransDate = Date2Num%(Form$(1, 0))
  UBTranRec(1).TransType = TranPenaltyCharge
  UBTranRec(1).TransDesc = TransDesc$
  UBTranRec(1).CustAcctNo = CCnt&
  UBTranRec(1).CustLocation = LocaRec&
  
  UBTranRec(1).RunBalance = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
  
  UBTranRec(1).PrevTrans = UBCustRec(1).LastTrans
  FOpenS "UBTRANS.DAT", THandle
  NextTransRec& = (FLof(THandle) / UBTranRecLen) + 1
  FPutRTA THandle, UBTranRec(1), NextTransRec&, UBTranRecLen
  FClose THandle
  
  UBCustRec(1).LastTrans = NextTransRec&
  FPutRTA CHandle, UBCustRec(1), CCnt&, UBCustRecLen
  
  RETURN
  
  CheckPenaltyFlds:
  CursorOff
  SaveScrn TempScrn()
  IF Value#(Form$(1, 0), ECode) <= 0 THEN
    DisplayUBScrn "UBNODATE"
    WaitForAction
    Frm(1).FldNo = 1
  ELSE
    OkFlag = True
  END IF
  Action = 1
  RestScrn TempScrn()
  
  RETURN
  
  
END SUB
  
