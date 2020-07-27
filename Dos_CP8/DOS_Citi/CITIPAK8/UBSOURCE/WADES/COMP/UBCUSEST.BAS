DEFINT A-Z
DECLARE SUB CustMessageSystem (RecNo&)
DECLARE SUB CustWorkOrders (CustRec&)
DECLARE SUB ShowCustConsHist (CustRec&)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB PaymentSumReport ()
DECLARE SUB MailingLabel ()
DECLARE SUB AgedBalanceListing ()
DECLARE SUB CustomerCutOffListing ()
DECLARE SUB CustomerInquiry ()
DECLARE SUB DetailedTransJournal ()
DECLARE SUB MasterBalanceListing ()
DECLARE SUB MasterCustomerListing ()
DECLARE SUB MasterDepositListing ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION FmtBook$ (BOOK$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION FindRateTbl% (RateCode$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE SUB AddEditLocation (RecNo&, FromFlag%)
  'DECLARE SUB Search4LNum (LNum$, RecNo&, CLSFlag%)
  'DECLARE SUB Search4Meter (MeterNum$, RecNo&, CLSFlag%)
  'DECLARE SUB ShowSearchWheel (BYVAL Row%, BYVAL Col%)
  'DECLARE SUB Search4SAddr (SAddr$, RecNo&, CLSFlag%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB Search4Cust (SEARCH$, RecNo&, CLSFlag%, LocationFlag%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION ConvDateStr$ (D$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FudgeFactor& ()
DECLARE FUNCTION NovellThere% ()
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE SUB BlockClear ()
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, x%, Y%)
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayScrn (BTmp%(), Element%, MonoCode%, WipeType%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB ExplodBox (UlRow%, UlCol%, BrRow%, BrCol%, Scr(), El%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB GetCursor (x, Y, Button)
DECLARE SUB HideCursor ()
DECLARE SUB BlockClear ()
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB MScrnRest (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnSave (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB Pause3 (MillaSecs%, ProcessorSpeed&)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB ShowCursor ()
DECLARE SUB SaveScrn (array())
DECLARE SUB RestScrn (array())
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)
DECLARE FUNCTION Chk4DupeBookSeqNum (BOOK$, SeqNum$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, LocationFlag%)
DECLARE SUB LoadUBSetUpFile (UBSetupRec() AS ANY, UBSetUpLen%)
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE SUB MakePostalIndex (IdxTypeText$)
DECLARE FUNCTION QPStrL$ (LongNumber&)
DECLARE FUNCTION WEnvTest% ()

  CONST False = 0, True = NOT False

  TYPE Struct
    TrDate AS INTEGER
    RecNum AS LONG
  END TYPE

  DIM SHARED PctC(1) AS STRING * 3

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'PageInfo.BI'                  'Form Page info
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'newcust.bi'

SUB CustEditMeterEST (CustRec&)

  REDIM TempScrn(0)
  REDIM TempScrn2(0)
  SaveScrn TempScrn()

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET #UBCust, CustRec&, UBCustRec(1)
  CLOSE

  
  LibName$ = "UB"
  ScrnName$ = "UBCUSEST"

  MPaintBox 3, 5, 22, 75, 8

  DisplayUBScrn ScrnName$

  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True

  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  FirstTime = True
  Action = 1

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      GOSUB FillEstForm
      Action = 1
    END IF

    SELECT CASE Frm(1).KeyCode

    CASE F10Key
      GOSUB UpDateCustomer
      Done = True
    CASE ESC
      Done = True
    END SELECT

    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 18   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 45 TO 54           'f10 Save
          PressButton F10Key, 18, 45, 54
        CASE 56 TO 67           '--cancel button
          PressButton EscKey, 18, 56, 67
        END SELECT
      END SELECT                'row
    END IF

  LOOP UNTIL Done

  RestScrn TempScrn()
  ERASE UBCustRec

EXIT SUB

FillEstForm:

  HasAMeterFlag = False

  FOR MCnt = 1 TO 7
    LSET Form$((5 * (MCnt - 1)) + 1, 0) = UBCustRec(1).LocMeters(MCnt).MTRNUM
    IF LEN(QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)) > 0 THEN
      HasAMeterFlag = True
      SELECT CASE UBCustRec(1).LocMeters(MCnt).MTRType
      CASE "W"
        LSET Form$((5 * (MCnt - 1)) + 2, 0) = "Water Only"
      CASE "S"
        LSET Form$((5 * (MCnt - 1)) + 2, 0) = "Sewer Only"
      CASE "C"
        LSET Form$((5 * (MCnt - 1)) + 2, 0) = "Combined Water & Sewer"
      CASE "E"
        LSET Form$((5 * (MCnt - 1)) + 2, 0) = "Electric Meter"
      CASE "D"
        LSET Form$((5 * (MCnt - 1)) + 2, 0) = "Demand Electric"
      CASE "G"
        LSET Form$((5 * (MCnt - 1)) + 2, 0) = "Gas Meter"
      CASE "T"
        LSET Form$((5 * (MCnt - 1)) + 2, 0) = "Touch Read Meter"
      END SELECT

      SELECT CASE UBCustRec(1).LocMeters(MCnt).MTRUnit
      CASE "G"
        LSET Form$((5 * (MCnt - 1)) + 3, 0) = "Gallons"
      CASE "C"
        LSET Form$((5 * (MCnt - 1)) + 3, 0) = "Cubic Feet"
      CASE "K"
        LSET Form$((5 * (MCnt - 1)) + 3, 0) = "Kilowatts"
      END SELECT

      IF UBCustRec(1).LocMeters(MCnt).UseCnt > 0 THEN
        LSET Form$((5 * (MCnt - 1)) + 4, 0) = STR$(UBCustRec(1).LocMeters(MCnt).UseCnt)
      ELSE
        LSET Form$((5 * (MCnt - 1)) + 4, 0) = STR$(0)
      END IF
      IF UBCustRec(1).LocMeters(MCnt).AvgUse > 0 THEN
        LSET Form$((5 * (MCnt - 1)) + 5, 0) = STR$(UBCustRec(1).LocMeters(MCnt).AvgUse)
      ELSE
        LSET Form$((5 * (MCnt - 1)) + 5, 0) = STR$(0)
      END IF

    END IF
  NEXT

  IF NOT HasAMeterFlag THEN
    LSET Form$(6, 0) = "No Metered"
    LSET Form$(11, 0) = "Services"
  END IF

  'QPrintRC InfoLine$, 5, 65, -1

RETURN

UpDateCustomer:
  SaveScrn TempScrn2()
  DisplayUBScrn "UPDATDSK"
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET #UBCust, CustRec&, UBCustRec(1)
  FOR MCnt = 1 TO 7
    LSET Form$((5 * (MCnt - 1)) + 1, 0) = UBCustRec(1).LocMeters(MCnt).MTRNUM
    IF LEN(QPTrim$(UBCustRec(1).LocMeters(MCnt).MTRType)) > 0 THEN
      AvgUse& = Value(Form$((5 * (MCnt - 1)) + 5, 0), ECode)
      IF UBCustRec(1).LocMeters(MCnt).AvgUse <> AvgUse& THEN
        UBCustRec(1).LocMeters(MCnt).AvgUse = AvgUse&
        UBCustRec(1).LocMeters(MCnt).UseCnt = 1
      END IF
    END IF
  NEXT
  PUT #UBCust, CustRec&, UBCustRec(1)
  CLOSE

  RestScrn TempScrn2()
  DisplayUBScrn "UPDATEOK"
  WaitForAction
RETURN

END SUB

