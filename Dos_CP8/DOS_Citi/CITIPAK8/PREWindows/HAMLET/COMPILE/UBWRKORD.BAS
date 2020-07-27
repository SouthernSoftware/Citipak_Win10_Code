DEFINT A-Z
DECLARE SUB RelinkWorkOrders ()
DECLARE FUNCTION Monitor
DECLARE SUB ShowWrkOrdHistory (CustRec&)
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE SUB CustMessageSystem (CustRec&)
DECLARE SUB PrintOpenWorkOrderRpt ()
DECLARE SUB PrintWorkOrders (RecNo&)
DECLARE SUB PrintOpenWorkOdrList ()
DECLARE SUB EnterEditWorkOrder ()
DECLARE SUB WorkOrdersMenu ()
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
DECLARE SUB ShowCustConsHist (CustRec&)
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
DECLARE FUNCTION QPValI% (Number$)
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
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB WaitForAction ()
  
  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  
  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  ''$INCLUDE: 'PageInfo.BI'                  'Form Page info
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ubtrans.BI'
  ''$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'newcust.bi'
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 5)

  MChoice$(1) = " Enter-Edit Customer Work Orders "
  MChoice$(2) = " Print Work Orders by Book"
  MChoice$(3) = " Print Open Work Order Report"
'  MChoice$(4) = "Relink Customer Work Orders"
  MChoice$(4) = " Exit to OS"

  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 1

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear

    TitleBox 2, Col, MaxLen + 3, "Work Order Processing Menu", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    CASE 1
      EnterEditWorkOrder
    CASE 2
      PrintWorkOrders 0
    CASE 3
      PrintOpenWorkOrderRpt
    CASE 4
      HideCursor
      CLS
      END
    END SELECT
  LOOP
  
  IF INSTR(COMMAND$, "TEST") = 0 THEN
    RUN "ubmenu"
  ELSE
    HideCursor
    ClearScrn
    END
  END IF

SUB EnterEditWorkOrder

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM WorkOrderRec(1) AS WorkOrderRecType
  WorkOrderRecLen = LEN(WorkOrderRec(1))
  
  BlockClear
  LookUp RecNo&, "Customer Work Orders", 2, True, False
  IF RecNo& = 0 THEN
    GOTO ExitEnterEditWorkOrd
  END IF
  
  LibName$ = "UBWRKORD"
  ScrnName$ = "WORKORDR"
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
  
  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  
  '--Set screen number to one and display screen
  Scr = 1
  LibFile2Scrn LibName$, ScrnName$, MonoCode, Attribute%, ErrCode
  
  ShowCursor
  
  FirstTime = True
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      GOSUB ChkGetLastWorkOrder
      GOSUB UpDateDisplay
      Action = 1
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F3Key
      ShowWrkOrdHistory RecNo&
      Action = 1
    CASE F4Key
      ShowCustHistory RecNo&
      Action = 1
    CASE F5Key        'Print
      REDIM TempScrn(0)
      SaveScrn TempScrn()
      GOSUB SaveWorkOrderRec
      PrintWorkOrders RecNo&
      RestScrn TempScrn()
      ExitFlag = True
      Action = 1
    CASE F6Key
      ShowCustConsHist RecNo&
      Action = 1
    CASE F7Key
      REDIM TempScrn(0)
      SaveScrn TempScrn()
      CustMessageSystem RecNo&
      RestScrn TempScrn()
      Action = 1
    CASE F10Key
      REDIM TempScrn(0)
      SaveScrn TempScrn()
      DisplayUBScrn "UPDATDSK"
      GOSUB SaveWorkOrderRec
      RestScrn TempScrn()
      DisplayUBScrn "UPDATEOK"
      WaitForAction
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 22   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 5 TO 14
          PressButton F3Key, 22, 5, 14
        CASE 15 TO 24
          PressButton F4Key, 22, 15, 24
        CASE 25 TO 34
          PressButton F5Key, 22, 25, 34
        CASE 35 TO 44
          PressButton F6Key, 22, 35, 44
        CASE 45 TO 53
          PressButton F7Key, 22, 45, 53
        CASE 54 TO 63           '--Save Button
          PressButton F10Key, 22, 54, 63
        CASE 64 TO 75           '--cancel button
          PressButton 27, 22, 64, 75
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
    
  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag
  
  GOTO ExitEnterEditWorkOrd
  
ExitEnterEditWorkOrd:
  
  EXIT SUB
  
ChkGetLastWorkOrder:
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, RecNo&, UBCustRec(1)
  LWTrans = UBCustRec(1).WOLastTrans
  IF LWTrans > 0 THEN
    EditFlag = True
    UBWrkOrd = FREEFILE
    OPEN "UBWRKORD.DAT" FOR RANDOM SHARED AS UBWrkOrd LEN = WorkOrderRecLen
    GET UBWrkOrd, LWTrans, WorkOrderRec(1)
    IF WorkOrderRec(1).CompletedDate > 0 THEN
      EditFlag = False
      AddingFlag = True
    ELSE
      BCopy VARSEG(WorkOrderRec(1)), VARPTR(WorkOrderRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), WorkOrderRecLen, 0
      UnPackBuffer 0, 0, Form$(), Fld()
      EditFlag = True
      AddingFlag = False
    END IF
    Action = 1
  ELSE
    EditFlag = False
    AddingFlag = True
  END IF
  
  CLOSE

  IF AddingFlag THEN
    LSET Form$(2, 0) = DATE$
    Frm(1).FldNo = 3
  END IF
  
RETURN
  
SaveWorkOrderRec:
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, RecNo&, UBCustRec(1)
  
  UBWrkOrd = FREEFILE
  OPEN "UBWRKORD.DAT" FOR RANDOM SHARED AS UBWrkOrd LEN = WorkOrderRecLen
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(WorkOrderRec(1)), VARPTR(WorkOrderRec(1)), WorkOrderRecLen, 0
  
  SELECT CASE AddingFlag
  CASE True
    IF UBCustRec(1).WOLastTrans > 0 THEN
      WorkOrderRec(1).PrevTransRec = UBCustRec(1).WOLastTrans
      WhatTrans = (LOF(UBWrkOrd) \ WorkOrderRecLen) + 1
      UBCustRec(1).WOLastTrans = WhatTrans
    ELSE
      WhatTrans = (LOF(UBWrkOrd) \ WorkOrderRecLen) + 1
      UBCustRec(1).WOLastTrans = WhatTrans
    END IF
  CASE False
    WhatTrans = UBCustRec(1).WOLastTrans
  END SELECT
  WorkOrderRec(1).CustRec = RecNo&
  PUT UBWrkOrd, WhatTrans, WorkOrderRec(1)
  UBCustRec(1).WOLastTrans = WhatTrans
  PUT UBCust, RecNo&, UBCustRec(1)
  
  CLOSE
  
  RETURN
  
UpDateDisplay:
  QPrintRC STR$(RecNo&), 3, 16, -1
  QPrintRC LEFT$(UBCustRec(1).CustName, 25), 4, 17, -1
  QPrintRC UBCustRec(1).Status, 3, 67, -1
  RETURN
  
END SUB

SUB PrintOpenWorkOrderRpt
  
  REDIM TempScrn(0)
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM WorkOrderRec(1) AS WorkOrderRecType
  WorkOrderRecLen = LEN(WorkOrderRec(1))
  
  Dash$ = STRING$(79, "-")
  
  MaxLines = 50
  FF$ = CHR$(12)
  
  'Open Report File
  ReportFile$ = "UBOPNWRK.RPT"
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  CustName$ = SPACE$(30)
  
  ' Location Order ********************************************************
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize&("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  UBCustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
  
  UBWOFile = FREEFILE
  OPEN "UBWRKORD.DAT" FOR RANDOM SHARED AS UBWOFile LEN = WorkOrderRecLen
  
  Cnt& = 1
  GOSUB PrintReadHeading
  ShowProcessingScrn "Processing Open Work Orders"
  FOR LCnt& = 1 TO IdxNumOfRecs
    GET #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)
    IF UBCustRec(1).WOLastTrans > 0 THEN
      GET #UBWOFile, UBCustRec(1).WOLastTrans, WorkOrderRec(1)
      IF WorkOrderRec(1).CompletedDate <= 0 THEN
        PRINT #RptHandle, USING "######"; IdxBuff(LCnt&).RecNum;
        PRINT #RptHandle, TAB(12); UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb; "  "; TAB(28); UBCustRec(1).CustName
        PRINT #RptHandle, UBCustRec(1).SERVADDR; TAB(50); UBCustRec(1).WOLastTrans; TAB(65); Num2Date$(WorkOrderRec(1).CompleteByDate)
        PRINT #RptHandle, Dash$
        LineCnt = LineCnt + 3
      END IF
    END IF
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintReadHeading
    END IF
    ShowPctComp LCnt&, IdxNumOfRecs
  NEXT
  
  PRINT #RptHandle, FF$
  CLOSE
  
  Header$ = "Open Work Orders Report"
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  EXIT SUB
  
PrintReadHeading:
  page = page + 1
  PRINT #RptHandle, TAB(30); "Open Work Order Report"
  PRINT #RptHandle, "Date: "; DATE$; TAB(70); "Page #"; page
  PRINT #RptHandle, "Acct No.   Location        Customer Name"
  PRINT #RptHandle, "Service Address                               Work Order #      Order Date"
  PRINT #RptHandle, STRING$(80, "=")
  LineCnt = 5
  
  RETURN
  
  
END SUB

SUB PrintWorkOrders (RecNo&)
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM WorkOrderRec(1) AS WorkOrderRecType
  WorkOrderRecLen = LEN(WorkOrderRec(1))
  
  Dash$ = STRING$(79, "_")
  
  FF$ = CHR$(12)
  
  IF RecNo& > 0 THEN
    PrintSingleFlag = True
    GOTO skipthis
  END IF
  
  LibName$ = "UBWRKORD"
  ScrnName$ = "UBPRNTWO"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  Action = 1
  FirstTime = True
  BlockClear
  
  LibFile2Scrn LibName$, ScrnName$, MonoCode, Attribute%, ErrCode
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      LSET Form$(1, 0) = "1"
      LSET Form$(2, 0) = "99"
      LSET Form$(3, 0) = "1"
      FirstTime = False
      Action = 1
    END IF
    
    '--Check for Key presses
    
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      ' Check for Proper Order of Route Questions
      IF QPValI(Form$(1, 0)) = 0 THEN
        BegRoute = 1
        EndRoute = 99
        inputting = True
      ELSE
        BegRoute = QPValI(Form$(1, 0))
        EndRoute = QPValI(Form$(2, 0))
      END IF
      
      IF BegRoute > EndRoute THEN
        COLOR 12
        LOCATE 13, 25
        PRINT "Invalid Route Order : Please Redo"
        BEEP
        WaitForAction
        LOCATE 13, 25
        PRINT STRING$(35, " ")
        Form$(1, 0) = ""
        Form$(2, 0) = ""
        Frm(1).FldNo = 1
        Action = 1
      ELSE
        inputting = True
      END IF
      Copies = QPValI(Form$(3, 0))
    CASE EscKey
      EXIT SUB
      
    END SELECT
    SELECT CASE Frm(1).MRow
    CASE 15   'Look for the f10 or esc button
      SELECT CASE Frm(1).MCol
      CASE 37 TO 46
        PressButton F10Key, 15, 37, 46
      CASE 49 TO 60
        PressButton 27, 15, 49, 60
      END SELECT
    END SELECT
  
  LOOP WHILE inputting = False
  
  ERASE Form$, Fld, Frm
  
skipthis:
  
  IF Copies < 1 THEN
    Copies = 1
  END IF
  
  'Open Report File
  ReportFile$ = "WORKORDR.RPT"
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  ' Location Order ********************************************************
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize&("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  UBCustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCustFile LEN = UBCustRecLen
  
  UBWOFile = FREEFILE
  OPEN "UBWRKORD.DAT" FOR RANDOM SHARED AS UBWOFile LEN = WorkOrderRecLen
  
  IF NOT PrintSingleFlag THEN
    Cnt& = 1
    ShowProcessingScrn "Processing Work Orders"
    FOR LCnt& = 1 TO IdxNumOfRecs
      GET #UBCustFile, IdxBuff(LCnt&).RecNum, UBCustRec(1)
      Book = QPValI(UBCustRec(1).Book)
      IF Book >= BegRoute AND Book <= EndRoute THEN
        IF UBCustRec(1).WOLastTrans > 0 THEN
          GET #UBWOFile, UBCustRec(1).WOLastTrans, WorkOrderRec(1)
          IF WorkOrderRec(1).CompletedDate <= 0 THEN
            GOSUB PrintThemOne
          END IF
        END IF
      END IF
      ShowPctComp LCnt&, IdxNumOfRecs
    NEXT
  ELSE
    ShowProcessingScrn "Processing Work Order"
    ShowPctComp 1, 1
    GET #UBCustFile, RecNo&, UBCustRec(1)
    GET #UBWOFile, UBCustRec(1).WOLastTrans, WorkOrderRec(1)
    GOSUB PrintThemOne
  END IF
  
  'PRINT #RptHandle, FF$
  
  CLOSE
  
  ERASE UBCustRec, WorkOrderRec, Frm, Form$, Fld, IdxBuff
  
  Header$ = "Customer Work Orders "
  
  EntryPoint = 1
  LPTPort = 1
  
  PrintRptFile Header$, ReportFile$, LPTPort, RetCode, EntryPoint
  
  EXIT SUB
  
PrintThemOne:
  FOR CopyCnt = 1 TO Copies
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, TAB(14); "W O R K   O R D E R   :   U T I L I T Y   D E P T ."
    PRINT #RptHandle,
    PRINT #RptHandle, "    Date Issued: "; DATE$
    PRINT #RptHandle, "  Work Order # : "; USING "######"; UBCustRec(1).WOLastTrans
    PRINT #RptHandle, "    Location # : "; UBCustRec(1).Book; "-"; UBCustRec(1).SeqNumb
    PRINT #RptHandle, "  Customer Name: "; UBCustRec(1).CustName
    PRINT #RptHandle, "Service Address: "; UBCustRec(1).SERVADDR
    PRINT #RptHandle, STRING$(79, "-")
    PRINT #RptHandle, "Instruction or Description of Work Needed"
    PRINT #RptHandle, STRING$(79, "-")
    PRINT #RptHandle, WorkOrderRec(1).OrdersText.Text(1)
    PRINT #RptHandle, WorkOrderRec(1).OrdersText.Text(2)
    PRINT #RptHandle, WorkOrderRec(1).OrdersText.Text(3)
    PRINT #RptHandle, WorkOrderRec(1).OrdersText.Text(4)
    PRINT #RptHandle, WorkOrderRec(1).OrdersText.Text(5)
    PRINT #RptHandle, WorkOrderRec(1).OrdersText.Text(6)
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, "Remarks Noted by Worker"
    PRINT #RptHandle,
    PRINT #RptHandle, Dash$
    PRINT #RptHandle,
    PRINT #RptHandle, Dash$
    PRINT #RptHandle,
    PRINT #RptHandle, Dash$
    PRINT #RptHandle,
    PRINT #RptHandle, Dash$
    PRINT #RptHandle,
    PRINT #RptHandle, Dash$
    PRINT #RptHandle,
    PRINT #RptHandle,
    PRINT #RptHandle, "BY: ______________________________   DATE: ____________________"
    PRINT #RptHandle,
    PRINT #RptHandle, "Meter Numbers:"

    FOR MtrCnt = 1 TO 7
      IF LEN(QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)) > 0 THEN
        PRINT #RptHandle, QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MtrNum)
      END IF
    NEXT
    PRINT #RptHandle, FF$;
  NEXT CopyCnt
  RETURN
  
  
END SUB

