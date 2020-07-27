DECLARE SUB TrashReport ()
DEFINT A-Z
DECLARE SUB WestCaroConsumpList ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE SUB SortServAddress (IndexText$)
DECLARE SUB CustEditMeterEST (CustRec&)
DECLARE SUB CustByRate ()
DECLARE SUB CustMessageSystem (RecNo&)
DECLARE SUB ShowWrkOrdHistory (RecNo&)
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
DECLARE FUNCTION FmtBook$ (Book$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION FindRateTbl% (RateCode$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE SUB AddEditLocation (RecNo&, FromFlag%)
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
DECLARE FUNCTION Chk4DupeBookSeqNum (Book$, SeqNum$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, LocationFlag%)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
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

  TrashReport
  ClearScrn

SUB TrashReport

  REDIM TempScrn(0)
  REDIM RevenueName$(15)
  REDIM Revenues(1 TO 15) AS DOUBLE

  REDIM UBTrans(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTrans(1))

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBSetUp(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUp(), UBSetUpRecLen

  ff$ = CHR$(12)
  MaxLines = 55

  FOR RCnt = 1 TO 15
    IF QPTrim$(UBSetUp(1).Revenues(RCnt).REVNAME) = "TRASH" THEN
      WhatRev = RCnt
      EXIT FOR
    END IF
  NEXT

  TownName$ = UBSetUp(1).UTILNAME
  TownLen = LEN(RTRIM$(TownName$))
  TabStop = 40 - (TownLen / 2)
  IF TabStop < 1 THEN TabStop = 1

  '*******************************

SelectFromList3:
  LibName$ = "UB"
  ScrnName$ = "UBTSHSUM"

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

  ' Set Defaults as Anticpated Response
  Form$(1, 0) = DATE$
  Form$(2, 0) = DATE$
  Action = 1

  BlockClear
  DisplayUBScrn ScrnName$

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      BegDate$ = Form$(1, 0)
      EndDate$ = Form$(2, 0)
      BegDate = Date2Num(Form$(1, 0))
      EndDate = Date2Num(Form$(2, 0))
      IF EndDate < BegDate THEN
        Done = False
        BEEP
        SaveScrn TempScrn()
        CursorOff
        DisplayUBScrn "UBBADDAT"
        WaitForAction
        RestScrn TempScrn()
        Frm(1).FldNo = 1
        Action = 1
      ELSE
        Done = True
      END IF

    CASE ESC
      GOTO ExitDetailedListing3
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done

  '***************
  ' Set Up Specifications from Input Screen
  Operator$ = Form$(3, 0)
  Operator = VAL(Operator$)

  IF Operator = 0 THEN
    BegOperator = 0
    EndOperator = 99
  ELSE
    BegOperator = Operator
    EndOperator = Operator
  END IF

  '***************
  MaxLines = 55
  PageNo = 0
  Dash80$ = STRING$(80, "-")

  UBRpt = FREEFILE
  OPEN "UBTRASHR.RPT" FOR OUTPUT AS UBRpt

  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTrans LEN = UBTransRecLen

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  NumOfRecs& = LOF(UBTrans) \ UBTransRecLen

  BlockClear
  ShowProcessingScrn "Trash Payment Report."

  GOSUB DoDetailedRptHeader3

  FOR Cnt& = 1 TO NumOfRecs&
    GET UBTrans, Cnt&, UBTrans(1)
    '*************************************
    '   Main Body of Printing goes here
    IF UBTrans(1).TransType = TranBillPayment THEN
      IF (UBTrans(1).TransDate >= BegDate AND UBTrans(1).TransDate <= EndDate) THEN
        IF (UBTrans(1).OperatorNumber >= BegOperator AND UBTrans(1).OperatorNumber <= EndOperator) THEN
          IF UBTrans(1).RevAmt(WhatRev) <> 0 THEN
            GOSUB PrintThisOne
          END IF
        END IF
      END IF
    END IF

    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, ff$
      GOSUB DoDetailedRptHeader3
    END IF

    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctCompL Cnt&, NumOfRecs&
  NEXT

  GOSUB DoDetailedRptFooter3

  CLOSE

  IF NOT AbortFlag THEN
    PrintRptFile "Trash Payment Report.", "UBTRASHR.RPT", 2, RetCode, EntryPoint
  END IF

  KillFile "UBPAYSUM.RPT"

ExitDetailedListing3:

  EXIT SUB

DoDetailedRptHeader3:
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(30); "Trash Payment Report"               '; TAB(70); "Page #1"
  PRINT #UBRpt, "Beginning Date: "; BegDate$;
  IF VAL(Operator$) = 0 THEN
    PRINT #UBRpt, TAB(65); " Operator #: ALL"
  ELSE
    PRINT #UBRpt, TAB(65); " Operator #: "; Operator$
  END IF
  PRINT #UBRpt, "   Ending Date: "; EndDate$
  PRINT #UBRpt, "Location   Acct  Customer Name                                 Amt Paid"
  PRINT #UBRpt, Dash80$
  LineCnt = 6
RETURN

DoDetailedRptFooter3:
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Total Customers: "; USING "###########"; TransCnt&
  PRINT #UBRpt, "    Trash Total: "; USING "$$#######,#.##"; TotalTrans#
  PRINT #UBRpt, ff$
RETURN

PrintThisOne:
  TransCnt& = TransCnt& + 1
  TotalTrans# = Round#(TotalTrans# + UBTrans(1).RevAmt(WhatRev))
  GET UBCust, UBTrans(1).CustAcctNo, UBCustRec(1)
  PRINT #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; USING "######"; UBTrans(1).CustAcctNo;
  PRINT #UBRpt, TAB(18); UBCustRec(1).CustName; TAB(64); USING "#####.##"; UBTrans(1).RevAmt(WhatRev)
  LineCnt = LineCnt + 1
RETURN

END SUB

