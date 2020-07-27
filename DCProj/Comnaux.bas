Attribute VB_Name = "Module1"
'DefInt A-Z
'DECLARE SUB MakeMowZipCodeIndex (IndexText$)
'DECLARE SUB Search4911Addr (S911$, RecNo&, CLSFlag%, ActiveOnly%)
'DECLARE FUNCTION CustHasMsg% (RecNo&)
'DECLARE SUB MakeZipCodeIndex (IndexText$)
'DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
'DECLARE FUNCTION GetZipEDigit$ (Zip$)
'DECLARE FUNCTION IsDeleted% (AcctNo&)
'DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
'DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
'DECLARE SUB ShowProcessingScrn (RptTitle$)
'DECLARE SUB AddEditLocation (RecNo&, FromFlag%)
'DECLARE SUB ShowSearchWheel (BYVAL Row%, BYVAL Col%)
'DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
'DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
'DECLARE SUB BlockClear ()
'DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, X%, Y%)
'DECLARE SUB ClearScrn ()
'DECLARE SUB CursorOff ()
'DECLARE SUB DisplayScrn (BTmp%(), Element%, MonoCode%, WipeType%)
'DECLARE SUB DisplayUBScrn (ScrnName$)
'DECLARE SUB ExplodBox (UlRow%, UlCol%, BrRow%, BrCol%, Scr(), El%)
'DECLARE SUB FCreate (FileName$)
'DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
'DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
'DECLARE SUB GetCursor (X, Y, Button)
'DECLARE SUB HideCursor ()
'DECLARE SUB BlockClear ()
'DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
'DECLARE SUB MScrnRest (UlRow, UlCol, LRRow, LRCol, SEG Address)
'DECLARE SUB MScrnSave (UlRow, UlCol, LRRow, LRCol, SEG Address)
'DECLARE SUB Pause3 (MillaSecs%, ProcessorSpeed&)
'DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
'DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
'DECLARE SUB ShowCursor ()
'DECLARE SUB SaveScrn (array())
'DECLARE SUB RestScrn (array())
'DECLARE SUB StuffBuf (Ky$)
'DECLARE SUB TextCursor (FG%, BG%)
'DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
'DECLARE SUB FOpenS (FileName$, Handle)
'DECLARE SUB FClose (Handle%)
'DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
'DECLARE SUB FGetA (Handle%, SEG Dest AS ANY, NumBytes%)
'DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
'DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
'DECLARE SUB WaitForAction ()
'DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)
'DECLARE FUNCTION Chk4DupeBookSeqNum (Book$, SeqNum$)
'DECLARE FUNCTION PromptSaveData% ()
'DECLARE FUNCTION Num2Date$ (DateNumber%)
'DECLARE FUNCTION FUsing$ (Number$, Image$)
'DECLARE FUNCTION QPValI% (Number$)
'DECLARE FUNCTION GetNumRateRecs% ()
'DECLARE FUNCTION FLof& (Handle%)
'DECLARE FUNCTION AskAbandonPrint% ()
'DECLARE FUNCTION FmtBook$ (Book$)
'DECLARE FUNCTION FmtSeqN$ (SeqN$)
'DECLARE FUNCTION FindRateTbl% (RATECODE$, NumOfRates%, UBRateTbls() AS ANY)
'DECLARE FUNCTION GetNumOfRevs% ()
'DECLARE FUNCTION FileSize& (FileName$)
'DECLARE SUB UBLog (Text$)
'DECLARE FUNCTION QPValL& (Number$)
'DECLARE FUNCTION MsgBox% (LibName$, FormName$)
'DECLARE FUNCTION GetNumOfCust% ()
'DECLARE FUNCTION ConvDateStr$ (d$)
'DECLARE FUNCTION Exist% (FileName$)
'DECLARE FUNCTION FudgeFactor& ()
'DECLARE FUNCTION NovellThere% ()
'DECLARE FUNCTION QPStrI$ (Num%)
'DECLARE FUNCTION QPStrL$ (Num&)
'DECLARE FUNCTION QPTrim$ (Text$)
'DECLARE FUNCTION Round# (N#)
'DECLARE SUB Search4Cust (SEARCH$, RecNo&, CLSFlag%, ActiveOnly%)
'DECLARE SUB Search4LNum (LNum$, RecNo&, CLSFlag%, ActiveOnly%)
'DECLARE SUB Search4Meter (MeterNum$, RecNo&, CLSFlag%, ActiveOnly%)
'DECLARE SUB Search4SAddr (SAddr$, RecNo&, CLSFlag%, ActiveOnly%)
'DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
'DECLARE SUB KillFile (File2Kill$)
'DECLARE FUNCTION GetCustMeterType% (UBCustRec() AS ANY, ThisMeter%)
'DECLARE FUNCTION ChkBillFile% ()
'DECLARE FUNCTION MakeMonth$ (TDate$)
'DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
'
'  Type FLen2
'    V As String * 64
'  End Type
'
'  '$INCLUDE: 'DefCnf.bi'      'Defines a TYPE for monitor and color info..
'  '$INCLUDE: 'FORMEDIT.bi'
'  '$INCLUDE: 'fieldinf.BI'
'  '$INCLUDE: 'newcust.bi'
'  '$INCLUDE: 'UBTRANS.bi'
'  '$INCLUDE: 'UBSETUP.bi'
'  '$INCLUDE: 'ubrate.BI'
'  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
'  '$INCLUDE: 'setcnf.bi'
'
'  CONST False = 0, True = NOT False
'  'CONST Wheel$ = "|/Ä\"
'
'  Dim PctC(1) As String * 3
'  'DIM SHARED DebugFlag AS INTEGER
'
'Function AskAbandonPrint%()
'
'  Static BeenHere, Escape$
'
'  If Not BeenHere Then
'    BeenHere = True
'    Escape$ = CHR$(27)
'  End If
'
'  Ky$ = INKEY$  'ButNo,DnFlg,NoPresses,col,row
'  ButtonPress 1, N, MooseButton, MCol, MRow     ' ----- Check Mouse activity
'  If N And 2 Then               'if it was the right moose button and
'    Do          'if they are still holding it down then
'      GetCursor 0, 0, NewStatus 'wait till they let go of the button
'    Loop While NewStatus
'    ButtonPress 0, 0, 0, 0, 0   'this has the effect of clearing all
'    Ky$ = Escape$               'moose buttons.
'  End If
'
'  If Ky$ <> Escape$ Then
'    Exit Function
'  End If
'
'  ReDim TempScrn(0)
'  SaveScrn TempScrn()
'  ActMRow = 14
'  BlockClear
'  DisplayUBScrn "ABORTJOB"
'
'  Do
'
'    Get.Moose.OR.Key Ky$, MooseButton, MRow, MCol
'
'    If MooseButton Then
'      MRow = (MRow \ 8) + 1     'Convert MCol, MRow to Row and Col.
'      MCol = (MCol \ 8) + 1
'      If MRow = ActMRow Then
'        Select Case MCol
'        Case 28 To 39           'Cancel printing
'          PressButton EscKey, ActMRow, 28, 39
'        Case 42 To 55           'let it rip
'          PressButton 13, ActMRow, 42, 55
'        End Select
'      End If
'    End If
'
'    If Len(Ky$) Then
'      If Ky$ = Escape$ Then
'        AskAbandonPrint% = True
'      Else
'        AskAbandonPrint% = False
'      End If
'      Exit Do
'    End If
'  Loop
'
'  RestScrn TempScrn()
'  Erase TempScrn
'
'End Function
'
'Function Chk4DupeBookSeqNum(Book$, SeqNum$)
'  Chk4DupeBookSeqNum = False    'assume it's ok
'  TBookSeq& = QPValL(Book$ + SeqNum$)
'  ReDim UBBookSeq(1) As BookSeqRecType
'  BookSeqLen = Len(UBBookSeq(1))
'  If FileSize("UBOOKSEQ.DAT") > 0 Then
'    FOpenS "UBOOKSEQ.DAT", Handle               'open data file
'    NumBookSeq = FLof(Handle) \ BookSeqLen
'    ReDim UBBookSeq(1 To NumBookSeq) As BookSeqRecType
'    FGetRTA Handle, UBBookSeq(1), 1&, NumBookSeq * BookSeqLen
'    FClose Handle
'    For Cnt = 1 To NumBookSeq
'      If UBBookSeq(Cnt).BookSeq = TBookSeq& Then
'        Ok = MsgBox%("UB.QSL", "DUPEBOOK")
'        Chk4DupeBookSeqNum = True
'        Exit For
'      End If
'    Next
'  End If
'  Erase UBBookSeq
'End Function
'
'Function ChkBillFile%()
'
'  OKFlag = True 'assume all is well
'
'  ReDim BillRec(1) As UBTransRecType
'  RecLen = Len(BillRec(1))
'
'  FHand = FREEFILE
'  Open UBBillsFile For Random Shared As FHand Len = RecLen
'  NumOfRec& = LOF(FHand) \ RecLen
'  Close FHand
'
'  If NumOfRec& = 0 Then
'    KILL UBBillsFile
'    OKFlag = False
'  End If
'
'  ChkBillFile% = OKFlag
'
'  Erase BillRec
'
'End Function
'
'Static Sub ClearScrn()
'  WazzWind 1, 1, 25, 80, 7, 0, 0
'End Sub
'
'Static Sub CursorOff()
'  LOCATE , , 0
'End Sub
'
'Function CustHasMsg(RecNo&)
'
'  ReDim MsgRec(1) As UBMessRecType
'  MsgLen = Len(MsgRec(1))
'
'  NumMsgRec& = FileSize&("UBMESAGE.DAT") / MsgLen
'
'  ReDim UBCustRec(1) As NewUBCustRecType
'  UBCustRecLen = Len(UBCustRec(1))
'
'  If RecNo& > 0 Then
'    UBFile = FREEFILE
'    Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
'    Get UBFile, RecNo&, UBCustRec(1)
'    Close UBFile
'    MRec& = UBCustRec(1).MessageRec
'
'    If MRec& > 0 And MRec& <= NumMsgRec& Then
'      MsgFile = FREEFILE
'      Open "UBMESAGE.DAT" For Random Shared As MsgFile Len = MsgLen
'      Get MsgFile, MRec&, MsgRec(1)
'      Close MsgFile
'      For ZZ = 1 To 15
'        m$ = QPTrim$(MsgRec(1).MessLine(ZZ).Line)
'        If Len(m$) > 0 Then
'          GotMsg = True
'          Exit For
'        End If
'      Next
'    Else
'      GotMsg = False
'    End If
'  Else
'    GotMsg = False
'  End If
'
'  If GotMsg Then
'    CustHasMsg = True
'  Else
'    CustHasMsg = False
'  End If
'
'End Function
'
'Sub CustMessageSystem(RecNo&)
'
'  CustRec& = RecNo&
'
'  ReDim ScrnArray(0)
'
'  ReDim UBCustRec(1) As NewUBCustRecType
'  UBCustRecLen = Len(UBCustRec(1))
'
'  ReDim UBMessRec(1) As UBMessRecType
'  UBMessRecLen = Len(UBMessRec(1))
'
'  UBCust = FREEFILE
'  Open "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
'  Get UBCust, CustRec&, UBCustRec(1)
'  Close
'
'  LibName$ = "UB"
'  ScrnName$ = "UBCUSMES"
'
'  ' Define Fields
'  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
'
'  ' Define Quick Screen Form Editing Arrays
'  ReDim frm(1) As FormInfo
'  ReDim Form$(NumFlds, 2)
'  ReDim Fld(NumFlds) As FieldInfo
'  frm(1).StayOnField = True
'
'  StartEl = 0
'  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
'
'  FirstTime = True
'
'  Action = 1
'
'  DisplayUBScrn ScrnName$
'  QPrintRC STR$(CustRec&), 3, 20, -1
'  QPrintRC UBCustRec(1).CustName, 4, 20, -1
'  QPrintRC UBCustRec(1).Status, 3, 67, -1
'
'  Do
'    EditForm Form$(), Fld(), frm(1), Cnf, Action
'
'    If FirstTime Then
'      FirstTime = False
'      GoSub LoadMessageInfo
'      Action = 1
'    End If
'
'    Select Case frm(1).KeyCode
'    Case F3Key
'      GoSub ClearRecord
'      GoSub ClearForm
'      Action = 1
'    Case F5Key
'      GoSub SaveRecord
'      GoSub PrintMessage
'    Case F10Key
'      SaveScrn ScrnArray()
'      DisplayUBScrn "UPDATDSK"
'      GoSub SaveRecord
'      RestScrn ScrnArray()
'      DisplayUBScrn "UPDATEOK"
'      WaitForAction
'      ExitFlag = True
'      RestScrn ScrnArray()
'      Done = True
'    Case ESC
'      Exit Sub
'    Case Else
'      Done = False
'    End Select
'  Loop Until Done
'
'ExitMessageInquiry:
'  Exit Sub
'  '***************
'
'LoadMessageInfo:
'  MessageRecord = UBCustRec(1).MessageRec
'  If MessageRecord > 0 Then
'    UBMess = FREEFILE
'    Open "UBMESAGE.DAT" For Random Shared As UBMess Len = UBMessRecLen
'    Get UBMess, MessageRecord, UBMessRec(1)
'    Close
'    BCopy VARSEG(UBMessRec(1)), VARPTR(UBMessRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), UBMessRecLen, 0
'    Call UnPackBuffer(0, 0, Form$(), Fld())
'  End If
'Return
'
'SaveRecord:
'  UBMess = FREEFILE
'  Open "UBMESAGE.DAT" For Random Shared As UBMess Len = UBMessRecLen
'  If MessageRecord = 0 Then
'    MessageRecord = LOF(UBMess) / Len(UBMessRec(1)) + 1
'  End If
'
'  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBMessRec(1)), VARPTR(UBMessRec(1)), UBMessRecLen, 0
'  Put UBMess, MessageRecord, UBMessRec(1)
'  Close
'
'  UBCust = FREEFILE
'  Open "UBCUST.DAT" For Random Shared As UBCust Len = UBCustRecLen
'  Get UBCust, CustRec&, UBCustRec(1)
'  UBCustRec(1).MessageRec = MessageRecord
'  Put UBCust, CustRec&, UBCustRec(1)
'  Close
'Return
'
'ClearRecord:
'  If MessageRecord > 0 Then
'    ReDim UBMessRec(1) As UBMessRecType
'    UBMess = FREEFILE
'    Open "UBMESAGE.DAT" For Random Shared As UBMess Len = UBMessRecLen
'    Put UBMess, MessageRecord, UBMessRec(1)
'    Close
'  End If
'Return
'
'ClearForm:
'  For F = 1 To NumFlds
'    LSet Form$(F, 0) = ""
'  Next F
'Return
'
'PrintMessage:
'  SaveScrn ScrnArray()
'  Dash$ = STRING$(80, "-")
'  ReDim UBSetUpRec(1) As UBSetupRecType
'  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
'  TownName$ = UBSetUpRec(1).UTILNAME
'  Erase UBSetUpRec
'  UBRpt = FREEFILE
'  Open "UBCUSMSG.RPT" For Output As UBRpt
'
'  Print #UBRpt, "Customer Messages Listing."; Tab(64); "Date: "; Date$
'  Print #UBRpt, "NAME: "; UBCustRec(1).CustName; "Acct:"; STR$(CustRec&)
'  Print #UBRpt, "Message Text"; Tab(70); "Entry Date"
'  Print #UBRpt, Dash$
'  For MsgLine = 1 To 15
'    Print #UBRpt, UBMessRec(1).MessLine(MsgLine).Line; Tab(70); UBMessRec(1).MessLine(MsgLine).LineDate
'  Next
'  Print #UBRpt, Dash$
'  Print #UBRpt, CHR$(12)
'  Close UBRpt
'  PrintRptFile "Customer Message Listing.", "UBCUSMSG.RPT", 1, RetCode, EntryPoint
'  RestScrn ScrnArray()
'  Action = 1
'
'Return
'
'End Sub
'
'Sub DisplayUBScrn(ScrnName$)
'  LibFile2Scrn "UB", ScrnName$, MonoCode, Attribute%, ErrCode
'End Sub
'
'Function FmtBook$(Book$)
'  Book$ = QPTrim$(Book$)
'  BookLen = Len(Book$)
'
'  Select Case BookLen
'  Case 0
'    FmtBook$ = "00"
'  Case 1
'    FmtBook$ = "0" + Book$
'  Case Else
'    FmtBook$ = Book$
'  End Select
'
'End Function
'
'Function FmtSeqN$(SeqN$)
'
'  SeqN$ = QPTrim$(SeqN$)
'  SeqNLen = Len(SeqN$)
'
'  Select Case SeqNLen
'  Case 0
'    FmtSeqN$ = "000000"
'  Case 1 To 5
'    FmtSeqN$ = "000000"
'    Mid$(FmtSeqN$, (6 - SeqNLen) + 1) = SeqN$
'  Case Else
'    FmtSeqN$ = SeqN$
'  End Select
'
'End Function
'
'Function GetCustMeterType(UBCustRec() As NewUBCustRecType, ThisMeter)
'
'  'Meter Types
'  'CONST MtrWaterOnly = 1
'  'CONST MtrSewerOnly = 2
'  'CONST MtrCombined = 3
'  'CONST MtrElectric = 4
'  'CONST MtrDemand = 5
'  'CONST MtrGas = 6
'  'CONST MtrTouchRead = 7
'
'  LMtrType$ = QPTrim$(UBCustRec(1).LocMeters(ThisMeter).MTRType)
'  LMtrTypeLen = Len(LMtrType$)
'  If LMtrTypeLen > 0 Then
'    Select Case LMtrType$
'    Case "W"
'      LThisMeter = MtrWaterOnly
'    Case "S"
'      LThisMeter = MtrSewerOnly
'    Case "C"
'      LThisMeter = MtrCombined
'    Case "E"
'      LThisMeter = MtrElectric
'    Case "D"
'      LThisMeter = MtrDemand
'    Case "G"
'      LThisMeter = MtrGas
'    Case "T"
'      LThisMeter = MtrTouchRead
'    Case Else
'      LThisMeter = True
'    End Select
'    GetCustMeterType = LThisMeter
'  Else
'    GetCustMeterType = 0
'  End If
'
'End Function
'
''This function returns the number of customer records
'Function GetNumOfCust()
'  ReDim TCustRec(1) As NewUBCustRecType
'  RecLen = Len(TCustRec(1))
'  CFileSize& = FileSize("UBCUST.DAT")
'  GetNumOfCust = CFileSize& \ RecLen
'  Erase TCustRec
'End Function
'
''This function return the number of rate codes
'Function GetNumRateRecs()
'  ReDim UBRateTblRec(1) As UBRateTblRecType
'  UBRateTblRecLen = Len(UBRateTblRec(1))
'  GetNumRateRecs = FileSize("UBRATE.DAT") \ UBRateTblRecLen
'  Erase UBRateTblRec
'End Function
'
'Function GetZipEDigit$(Zip$)
'
'  ZipLen = Len(Zip$)
'  ZipVal = 0
'
'  DashPos = InStr(Zip$, "-")
'  Do While DashPos
'    Zip$ = LEFT$(Zip$, DashPos - 1) + MID$(Zip$, DashPos + 1)
'    DashPos = InStr(Zip$, "-")
'  Loop
'
'  For Cnt = 1 To ZipLen
'    ZipVal = ZipVal + VAL(MID$(Zip$, Cnt, 1))
'  Next
'
'  If ZipVal Mod 10 > 0 Then
'    Dif = 10 - (ZipVal Mod 10)
'  Else
'    Dif = 0
'  End If
'
'  GetZipEDigit$ = QPTrim$(STR$(Dif))
'
'End Function
'
''Returns TRUE if this is a deleted account
'Function IsDeleted%(AcctNo&)
'  ReDim UBCustRec(1) As NewUBCustRecType
'  UBCustRecLen = Len(UBCustRec(1))
'  FOpenS "UBCUST.DAT", C1Handle
'  FGetRTA C1Handle, UBCustRec(1), AcctNo&, UBCustRecLen
'  FClose C1Handle
'  If UBCustRec(1).DelFlag <> 0 Then
'    IsDeleted% = True
'  Else
'    IsDeleted% = False
'  End If
'  Erase UBCustRec
'End Function
'
'Sub LoadUBSetUpFile(UBSetUpRec() As UBSetupRecType, UBSetupLen)
'                       'use the length as an error flag
'  UBSetupLen = -1      'assume the file is not there, or 0 bytes.
'  If Exist("UBSETUP.DAT") Then
'    FOpenS "UBSETUP.DAT", Handle                'open data file
'    If FLof&(Handle) > 0 Then
'      UBSetupLen = Len(UBSetUpRec(1))
'      FGetRTA Handle, UBSetUpRec(1), 1&, UBSetupLen
'    End If
'    FClose Handle
'  End If
'
'End Sub
'
'Sub LookUp(RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
'
'  Static SName$, AcctNum&, MeterNum$, SAddr$, LNum$
'
'  ReDim UBCustRec(1) As NewUBCustRecType
'  UBCustRecLen = Len(UBCustRec(1))
'
'  ReDim UBSetUpRec(1) As UBSetupRecType
'  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
'
'  SName$ = ""
'  AcctNum& = 0
'  MeterNum$ = ""
'  SAddr$ = ""
'  LNum$ = ""
'  S911$ = ""
'
'  If InStr(COMMAND$, "DEBUG") Then
'    DebugFlag = True
'  End If
'
'  Select Case QPValI(UBSetUpRec(1).DefLook)
'  Case 1
'    LScrn = 1
'  Case 2
'    LScrn = 2
'  Case 3
'    LScrn = 3
'  Case 4
'    LScrn = 4
'  Case 5
'    LScrn = 5
'  Case 6
'    LScrn = 6
'  Case Else
'    LScrn = DefaultLook
'  End Select
'
'  CursorOff
'
'  ReDim ScrnArray(0)
'  ReDim ScrnArray2(0)
'
'  SaveScrn ScrnArray()
'
'  ReDim LText$(6)
'
'  MScrn = 6
'
'  LText$(1) = " Account Number:"
'  LText$(2) = "    Search Name:"
'  LText$(3) = "   Meter Number:"
'  LText$(4) = "Service Address:"
'  LText$(5) = "Location Number:"
'  LText$(6) = "      911/Other:"
'
'
'  LibName$ = "UB"
'  ScrnName$ = "LUPACCT"
'
'  '--Initialize the form name array
'  '--Get the total number of fields from all pages
'  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
'
'  '--define Quick Screen form editing arrays
'  ReDim frm(1) As FormInfo
'  ReDim Form$(NumFlds, 2)
'  ReDim Fld(NumFlds) As FieldInfo
'
'  '--for each screen, get first and last fields
'  StartEl = 0
'  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
'
'  '--Clear all fields
'  For F = 1 To NumFlds
'    LSet Form$(F, 0) = ""
'  Next
'  Text$ = Text$ + " Look-Up"
'  TextLen = Len(Text$)
'  TCol = ((80 - TextLen) \ 2)
'  DisplayUBScrn ScrnName$
'
'  QPrintRC Text$, 8, TCol, -1
'
'  GoSub DisplayLookupText
'
'  ShowCursor
'
'  Action = 1
'  FirstTime = True
'
'  Do
'
'    EditForm Form$(), Fld(), frm(1), Cnf, Action
'    If FirstTime Then
'      FirstTime = Not FirstTime
'      Select Case LScrn
'      Case 1
'        If AcctNum& > 0 Then
'          Form$(1, 0) = QPStrL$(AcctNum&)
'        End If
'      Case 2
'        Form$(1, 0) = SName$
'      Case 3
'        Form$(1, 0) = MeterNum$
'      Case 4
'        Form$(1, 0) = SAddr$
'      Case 5
'        Form$(1, 0) = LNum$
'      Case 6
'        Form$(1, 0) = S911$
'      End Select
'      Action = 1
'    End If
'
'    '--Check for Key presses
'    Select Case frm(1).KeyCode
'    Case -68, 13                'F10Key    Proceed with look up
'      CursorOff
'      Select Case LScrn
'      Case 1    'account lookup
'        AcctNum& = QPValL(Form$(1, 0))
'        If AcctNum& < 1 Or AcctNum& > GetNumOfCust Then
'          Ok = MsgBox%("UB.QSL", "BADACCTN")
'        Else
'          If IsDeleted(AcctNum&) Then
'            Ok = MsgBox%("UB.QSL", "DELACCTN")
'          ElseIf ActiveOnly Then
'            CHand = FREEFILE
'            Open "UBCUST.DAT" For Random Shared As CHand Len = UBCustRecLen
'            Get #CHand, AcctNum&, UBCustRec(1)
'            Close CHand
'            If UBCustRec(1).Status = "A" Then
'              RecNo& = AcctNum&
'              OKFlag = True
'            Else
'              Ok = MsgBox%("UB.QSL", "INACTACT")
'            End If
'          Else
'            RecNo& = AcctNum&
'            OKFlag = True
'          End If
'        End If
'
'      Case 2    'Name lookup
'        SName$ = LEFT$(QPTrim$(Form$(0, 0)), 10)
'        If Len(SName$) = 0 Then
'          SName$ = SPACE$(10)
'        End If
'        SaveScrn ScrnArray2()
'        RestScrn ScrnArray()
'        Search4Cust SName$, RecNo&, CLSFlag, ActiveOnly
'        If RecNo& > 0 Then
'          OKFlag = True
'        ElseIf RecNo& = 0 Then
'          Ok = MsgBox%("UB.QSL", "NOMATCH")
'        End If
'        RestScrn ScrnArray2()
'        Action = 1
'        'END IF
'      Case 3    'meter number
'        MeterNum$ = QPTrim$(Form$(0, 0))
'        If Len(MeterNum$) = 0 Then
'          Ok = MsgBox%("UB.QSL", "BADSEARH")
'          Action = 1
'          frm(1).FldNo = 1
'        Else
'          SaveScrn ScrnArray2()
'          RestScrn ScrnArray()
'          Search4Meter MeterNum$, RecNo&, CLSFlag, ActiveOnly
'          If RecNo& > 0 Then
'            OKFlag = True
'          ElseIf RecNo& = 0 Then
'            Ok = MsgBox%("UB.QSL", "NOMATCH")
'          End If
'          RestScrn ScrnArray2()
'          Action = 1
'        End If
'      Case 4    'service address
'        SAddr$ = QPTrim$(Form$(0, 0))
'        If Len(SAddr$) = 0 Then
'          Ok = MsgBox%("UB.QSL", "BADSEARH")
'          Action = 1
'          frm(1).FldNo = 1
'        Else
'          SaveScrn ScrnArray2()
'          RestScrn ScrnArray()
'          Search4SAddr SAddr$, RecNo&, CLSFlag, ActiveOnly
'          If RecNo& > 0 Then
'            OKFlag = True
'          ElseIf RecNo& = 0 Then
'            Ok = MsgBox%("UB.QSL", "NOMATCH")
'          End If
'          RestScrn ScrnArray2()
'          Action = 1
'        End If
'      Case 5    'Location lookup
'        OK2Search = False
'        LNum$ = QPTrim$(Form$(0, 0))
'        DashPos = InStr(LNum$, "-")
'        If Len(LNum$) < 2 Then  'OR DashPos <= 0 THEN
'          Ok = MsgBox%("UB.QSL", "BADACCTN")
'          Action = 1
'          frm(1).FldNo = 1
'        ElseIf DashPos > 1 Then
'          Book$ = FmtBook$(LEFT$(LNum$, DashPos - 1))
'          SeqN$ = FmtSeqN$(MID$(LNum$, DashPos + 1))
'          LNum$ = Book$ + "-" + SeqN$
'          LSet Form$(1, 0) = LNum$
'          SaveField 0, Form$(), Fld(), BadField
'          Action = 1
'          OK2Search = True
'        Else
'          Book$ = FmtBook$(LEFT$(LNum$, 2))
'          SeqN$ = FmtSeqN$(MID$(LNum$, 3))
'          LNum$ = Book$ + "-" + SeqN$
'          LSet Form$(1, 0) = LNum$
'          SaveField 0, Form$(), Fld(), BadField
'          Action = 1
'          OK2Search = True
'        End If
'        If OK2Search Then
'          SaveScrn ScrnArray2()
'          RestScrn ScrnArray()
'          Search4LNum LNum$, RecNo&, CLSFlag, ActiveOnly
'          If RecNo& > 0 Then
'            OKFlag = True
'          ElseIf RecNo& = 0 Then
'            Ok = MsgBox%("UB.QSL", "NOMATCH")
'          End If
'          RestScrn ScrnArray2()
'          Action = 1
'        End If
'      Case 6   '911 Address
'        S911$ = QPTrim$(Form$(0, 0))
'        If Len(S911$) = 0 Then
'          Ok = MsgBox%("UB.QSL", "BADSEARH")
'          Action = 1
'          frm(1).FldNo = 1
'        Else
'          SaveScrn ScrnArray2()
'          RestScrn ScrnArray()
'          Search4911Addr S911$, RecNo&, CLSFlag, ActiveOnly
'          If RecNo& > 0 Then
'            OKFlag = True
'          ElseIf RecNo& = 0 Then
'            Ok = MsgBox%("UB.QSL", "NOMATCH")
'          End If
'          RestScrn ScrnArray2()
'          Action = 1
'        End If
'      End Select
'
'    Case -65    'F7Key
'      If LScrn < MScrn Then
'        LScrn = LScrn + 1
'      Else
'        LScrn = 1
'      End If
'      LSet Form$(1, 0) = ""
'      Action = 1
'      FirstTime = True
'      SaveField 0, Form$(), Fld(), BadField
'      GoSub DisplayLookupText
'    Case 27
'      RecNo& = 0
'      ExitFlag = True
'    End Select
'
'    If frm(1).Presses Then
'      Select Case frm(1).MRow
'      Case 16
'        Select Case frm(1).MCol
'        Case 22 To 33           'ESC Cancel button
'          PressButton 27, 16, 22, 33
'        Case 35 To 45           'F7 Toggle Choice
'          PressButton -65, 16, 35, 45
'        Case 47 To 59           'F10 Save Button
'          PressButton -68, 16, 47, 59
'        End Select
'      End Select
'    End If
'
'  Loop Until ExitFlag Or OKFlag
'  RestScrn ScrnArray()
'
'  Erase frm, Form$, Fld, UBSetUpRec
'  Erase ScrnArray, ScrnArray2, UBCustRec
'  Erase LText$
'
'  Exit Sub
'
'DisplayLookupText:
'  QPrintRC LText$(LScrn), 12, 15, -1
'Return
'
'
'End Sub
'
'Function MakeMonth$(TDate$)
'  Month = VAL(LEFT$(TDate$, 2))
'  Select Case Month
'  Case 1
'    MakeMonth$ = "January"
'  Case 2
'    MakeMonth$ = "February"
'  Case 3
'    MakeMonth$ = "March"
'  Case 4
'    MakeMonth$ = "April"
'  Case 5
'    MakeMonth$ = "May"
'  Case 6
'    MakeMonth$ = "June"
'  Case 7
'    MakeMonth$ = "July"
'  Case 8
'    MakeMonth$ = "August"
'  Case 9
'    MakeMonth$ = "September"
'  Case 10
'    MakeMonth$ = "October"
'  Case 11
'    MakeMonth$ = "November"
'  Case 12
'    MakeMonth$ = "December"
'  End Select
'End Function
'
'Sub MakeMowZipCodeIndex(IndexText$)
'
'  ShowProcessingScrn "Creating " + IndexText$ + " Index "
'  QPrintRC "    Reading Customer Records     ", 11, 25, -1
'
'  ReDim UBCustRec(1) As NewUBCustRecType
'  CustRecLen = Len(UBCustRec(1))
'
'  NumOfBillRec = FileSize("UBCUST.DAT") \ CustRecLen
'
'  CHandle = FREEFILE
'  Open "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
'
'  ReDim ZipIndex(1 To NumOfBillRec) As MOWZipIndexType
'  For BCnt = 1 To NumOfBillRec
'    Get CHandle, BCnt, UBCustRec(1)
'    ZipIndex(BCnt).ZIPCODE = UBCustRec(1).ZIPCODE
'    ZipIndex(BCnt).RecNum = BCnt
'    ShowPctComp BCnt, NumOfBillRec              'show user percentage complete
'  Next
'  Close
'
'  QPrintRC "         Sorting Index.        ", 11, 25, -1
'  SortT ZipIndex(1), NumOfBillRec, 0, 16, 0, 10
'  QPrintRC "      Writing Index Records      ", 11, 25, -1
'
'  IHandle = FREEFILE
'  Open TempIndexName For Output As IHandle
'  Close IHandle
'
'  IHandle = FREEFILE
'  Open TempIndexName For Random Shared As IHandle Len = 4
'  For Cnt = 1 To NumOfBillRec
'    Prec& = ZipIndex(Cnt).RecNum
'    Put IHandle, Cnt, Prec&
'    ShowPctComp Cnt, NumOfBillRec               'show user percentage complete
'  Next
'  Close IHandle
'
'  Erase UBCustRec, ZipIndex
'
'End Sub
'
'Sub MakePostalIndex(IndexText$)
'
'  ShowProcessingScrn "Creating " + IndexText$ + " Index"
'  QPrintRC "    Reading Customer Records     ", 11, 25, -1
'
'  ReDim UBCustRec(1) As NewUBCustRecType
'  CustRecLen = Len(UBCustRec(1))
'
'  NumCustRecs = GetNumOfCust%
'
'  ReDim PostalIndex(1 To NumCustRecs) As UBPostalIndexType
'  IndexRecLen = Len(PostalIndex(1))
'
'  CHandle = FREEFILE
'  Open "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
'  For Cnt = 1 To NumCustRecs
'    Get CHandle, Cnt, UBCustRec(1)
'    PostalIndex(Cnt).ZIPCODE = UBCustRec(1).ZIPCODE
'    RSet PostalIndex(Cnt).Route = QPTrim$(UBCustRec(1).POSTRTE)
'    PostalIndex(Cnt).RecNum = Cnt
'    ShowPctComp Cnt, NumCustRecs                'show user percentage complete
'  Next
'
'  Close CHandle
'  QPrintRC "         Sorting Index.        ", 11, 25, -1
'  SortT PostalIndex(1), NumCustRecs, 0, 16, 10, 4
'  QPrintRC "      Writing Index Records      ", 11, 25, -1
'  IHandle = FREEFILE
'
'  FCreate TempIndexName
'
'  Open TempIndexName For Random Shared As IHandle Len = 4
'  For Cnt = 1 To NumCustRecs
'    Prec& = PostalIndex(Cnt).RecNum
'    Put IHandle, Cnt, Prec&
'    ShowPctComp Cnt, NumCustRecs                'show user percentage complete
'  Next
'  Close IHandle
'
'  Erase UBCustRec, PostalIndex
'
'End Sub
'
'Sub MakeSequenceIndex(IndexText$)
'  ShowProcessingScrn "Creating " + IndexText$ + " Index"
'  QPrintRC "    Reading Location Records     ", 11, 25, -1
'
'  ReDim UBCustRec(1) As NewUBCustRecType
'  CustRecLen = Len(UBCustRec(1))
'
'  NumCustRecs& = GetNumOfCust%
'
'  ReDim SequenceIndex(1 To NumCustRecs&) As UBSequenceIndexType
'  IndexRecLen = Len(SequenceIndex(1))
'
'  CHandle = FREEFILE
'  Open "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
'  For Cnt = 1 To NumCustRecs&
'    Get CHandle, Cnt, UBCustRec(1)
'    SequenceIndex(Cnt).SeqNumber = UBCustRec(1).SEQ
'    SequenceIndex(Cnt).RecNum = Cnt
'    ShowPctComp Cnt, NumCustRecs&               'show user percentage complete
'  Next
'
'  Close CHandle
'
'  QPrintRC "         Sorting Index.        ", 11, 25, -1
'
'  SortT SequenceIndex(1), CInt(NumCustRecs&), 0, 16, 0, -2
'
'  QPrintRC "      Writing Index Records      ", 11, 25, -1
'
'  FCreate TempIndexName
'
'  IHandle = FREEFILE
'  Open TempIndexName For Random Shared As IHandle Len = 4
'
'  For Cnt = 1 To NumCustRecs&
'    Prec& = SequenceIndex(Cnt).RecNum
'    Put IHandle, Cnt, Prec&
'    ShowPctComp Cnt, NumCustRecs&               'show user percentage complete
'  Next
'  Close IHandle
'
'  Erase UBCustRec, SequenceIndex
'
'End Sub
'
'Sub MakeZipCodeIndex(IndexText$)
'
'  MakeMowZipCodeIndex IdxTypeText$
'  Exit Sub
'
'  ShowProcessingScrn "Creating " + IndexText$ + " Index "
'  QPrintRC "    Reading Customer Records     ", 11, 25, -1
'
'  'REDIM ZipIndex(1 TO 1)  AS PSAZipIndexType
'  ReDim UBCustRec(1) As NewUBCustRecType
'  CustRecLen = Len(UBCustRec(1))
'
'  NumOfBillRec = FileSize("UBCUST.DAT") \ CustRecLen
'
'  CHandle = FREEFILE
'  Open "UBCUST.DAT" For Random Shared As CHandle Len = CustRecLen
'
'  'REDIM ZipIndex(1 TO NumOfBillRec)   AS PSAZipIndexType
'
'  ReDim ZipIndex(1 To NumOfBillRec) As MOWZipIndexType
'  For BCnt = 1 To NumOfBillRec
'    Get CHandle, BCnt, UBCustRec(1)
'    ZipIndex(BCnt).ZIPCODE = UBCustRec(1).ZIPCODE
'    'ZipIndex(BCnt).SName = UBCustRec(1).SEARCH
'    ZipIndex(BCnt).RecNum = BCnt
'    ShowPctComp BCnt, NumOfBillRec              'show user percentage complete
'  Next
'
'  Close
'
'  QPrintRC "         Sorting Index.        ", 11, 25, -1
'
'  SortT ZipIndex(1), NumOfBillRec, 0, 16, 0, 10
'
'  First = 1
'  Last = 1
'
'  SZip$ = ZipIndex(1).ZIPCODE
'
'  For ZCnt = 2 To NumOfBillRec
'    EZip$ = ZipIndex(ZCnt).ZIPCODE
'    If SZip$ <> EZip$ Then
'      Last = ZCnt - 1
'      GoSub SortThisZip
'      First = ZCnt
'      SZip$ = EZip$
'    End If
'    ShowPctComp ZCnt, NumOfBillRec              'show user percentage complete
'  Next
'  Last = ZCnt - 1
'  GoSub SortThisZip
'
'  QPrintRC "      Writing Index Records      ", 11, 25, -1
'
'  IHandle = FREEFILE
'  Open TempIndexName For Output As IHandle
'  Close IHandle
'
'  IHandle = FREEFILE
'  Open TempIndexName For Random Shared As IHandle Len = 4
'  For Cnt = 1 To NumOfBillRec
'    Prec& = ZipIndex(Cnt).RecNum
'    Put IHandle, Cnt, Prec&
'    ShowPctComp Cnt, NumOfBillRec               'show user percentage complete
'  Next
'  Close IHandle
'
'  Erase UBCustRec, ZipIndex
'
'  Exit Sub
'
'SortThisZip:
'  If First < Last Then
'    'SortT ZipIndex(First), Last - First + 1, 0, 32, 10, 10
'    SortT ZipIndex(First), Last - First + 1, 0, 16, 10, 10
'  End If
'Return
'
'End Sub
'
'Function PromptSaveData%()
'
'  ReDim TempScrn(0)
'  SaveScrn TempScrn()
'
'  LibName$ = "UB"
'  SaveFlag = 2
'
'  FormName$ = "SAVE1ST"
'  NumFlds = LibNumberOfFields(LibName$, FormName$)
'
'  ReDim frm(1) As FormInfo
'  ReDim Form$(NumFlds, 2)       'DIM the form data array
'  ReDim Fld(NumFlds) As FieldInfo               'DIM the field information array
'  StartEl = 0   'Load first form at array start
'  LibGetFldDef LibName$, FormName$, StartEl, Fld(), Form$(), ErrCode
'
'
'  '----- Set the "Action" flag to force the editor to initialize itself and
'  '      display the data on the form.
'  Action = 1
'
'  '----- Setup TYPE for setting and reading form editing information.
'  frm(1).FldNo = 1              'Start editing on field #1
'  frm(1).InsStat = False        'Set insert state (True = Insert on)
'  frm(1).StartEl = 0            'Set form starting element to 0 and
'
'  DisplayUBScrn FormName$
'
'  Do
'    EditForm Form$(), Fld(), frm(1), Cnf, Action
'    Select Case frm(1).KeyCode
'    Case F0Key
'      SaveFlag = True
'    Case EscKey
'      SaveFlag = 1
'    Case 88, 120                'X Key
'      SaveFlag = False
'    End Select
'
'  Loop While SaveFlag = 2       'proper key not set
'
'  PromptSaveData = SaveFlag
'  CursorOff
'
'  RestScrn TempScrn()
'
'  Erase TempScrn, Form$, Fld, frm
'
'End Function
'
'Sub ReIndexSystem(PromptFlag%)
'
'  UBLog " IN: Reindex Utility Files"
'
'  BlockClear
'  If PromptFlag% Then
'    Ok = MsgBox%("UB", "MUSTEXIT")
'    Select Case Ok
'    Case 2
'      GoTo ExitReindex
'    End Select
'  End If
'
'  'BlockClear
'  ReDim UBCustRec(1) As NewUBCustRecType
'  UBCustRecLen = Len(UBCustRec(1))              'Length of Cust Record Structure
'
'  ReDim UBTransRec(1) As UBTransRecType
'  UBTranRecLen = Len(UBTransRec(1))             'Length of Tran Record Structure
'
'  ShowProcessingScrn "Reading Customer Names"
'  UBLog "BEGIN: Customer Name Reindex"
'  UBFile = FREEFILE
'  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
'  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
'
'  ReDim IdxBuff(1 To NumOfRecs&) As nUBCustReIndexRecType
'
'  For Cnt = 1 To NumOfRecs&
'    Get UBFile, Cnt, UBCustRec(1)
'    IdxBuff(Cnt).SearchName = UBCustRec(1).SEARCH
'    If UBCustRec(1).DelFlag Then
'      IdxBuff(Cnt).DelFlag = "Y"
'    Else
'      IdxBuff(Cnt).DelFlag = ""
'    End If
'    IdxBuff(Cnt).Status = UBCustRec(1).Status
'    IdxBuff(Cnt).RecNum = Cnt
'    ShowPctComp Cnt, NumOfRecs&
'  Next
'
'  Close UBFile
'
'  QPrintRC " Sorting Customer Names", 11, 29, -1
'
'  SortT IdxBuff(1), CInt(NumOfRecs&), 0, 16, 0, 10
'
'  GoSub ClearBlock
'  QPrintRC "Writing Customer Index", 9, 30, -1
'  QPrintRC "Processing:    % Complete", 13, 28, -1
'
'  KillFile "UBCUSTNM.IDX"
'  UBFile = FREEFILE
'  Open "UBCUSTNM.IDX" For Random Shared As UBFile Len = 4
'  For Cnt = 1 To NumOfRecs&
'    Put UBFile, Cnt, IdxBuff(Cnt).RecNum
'    ShowPctComp Cnt, NumOfRecs&
'  Next
'  Close UBFile
'
'  GoSub ClearBlock
'  QPrintRC "Writing Customer Search Data", 9, 27, 126
'  QPrintRC "Processing:    % Complete", 13, 28, -1
'
'  KillFile "UBCUSTSN.DAT"
'  UBFile = FREEFILE
'  Open "UBCUSTSN.DAT" For Random Shared As UBFile Len = Len(IdxBuff(1))
'  For Cnt = 1 To NumOfRecs&
'    Put UBFile, Cnt, IdxBuff(Cnt)
'    ShowPctComp Cnt, NumOfRecs&
'  Next
'  Close UBFile
'
'  Erase IdxBuff
'  UBLog "FINISH: Customer Name Reindex"
'  GoSub ClearBlock
'
'  QPrintRC "Reading Location Information", 9, 27, 126
'  QPrintRC "Processing:    % Complete", 13, 28, -1
'  UBLog "BEGIN: Book\Sequence Reindex"
'
'  UBFile = FREEFILE
'  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
'  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
'
'  ReDim LIdxBuff(1 To NumOfRecs&) As UBLocaReIndexRecType
'
'  For Cnt = 1 To NumOfRecs&
'    Get UBFile, Cnt, UBCustRec(1)
'    LIdxBuff(Cnt).Book = UBCustRec(1).Book
'    LIdxBuff(Cnt).SEQNUMB = UBCustRec(1).SEQNUMB
'    LIdxBuff(Cnt).RecNum = Cnt
'    ShowPctComp Cnt, NumOfRecs&
'  Next
'
'  Close UBFile
'
'  QPrintRC " Sorting Locations Names", 11, 29, -1
'
'  SortT LIdxBuff(1), CInt(NumOfRecs&), 0, 16, 0, 8
'  'Array(1), NumElem, Dir, StructSize, MemOff, MemSize
'
'  GoSub ClearBlock
'  QPrintRC "Writing Location Index", 9, 30, -1
'  QPrintRC "Processing:    % Complete", 13, 28, -1
'  'here
'  KillFile "UBCUSTBK.IDX"
'
'  UBFile = FREEFILE
'  Open "UBCUSTBK.IDX" For Random Shared As UBFile Len = 4
'
'  For Cnt = 1 To NumOfRecs&
'    Put UBFile, Cnt, LIdxBuff(Cnt).RecNum
'    ShowPctComp Cnt, NumOfRecs&
'  Next
'  Close UBFile
'
'  UBLog "FINISH: Book\Sequence Reindex"
'  ReDim BookSeq(1) As BookSeqRecType
'
'  KillFile "UBOOKSEQ.DAT"
'  UBLog "BEGIN: Rebuild Book\Sequence List"
'  BookHand = FREEFILE
'  Open "UBOOKSEQ.DAT" For Random Shared As BookHand Len = 4
'  For Cnt = 1 To NumOfRecs&
'    BookSeq(1).BookSeq = QPValL(LIdxBuff(Cnt).Book + LIdxBuff(Cnt).SEQNUMB)
'    Put BookHand, Cnt, BookSeq(1)
'  Next
'  Close BookHand
'  UBLog "FINISH: Rebuild Book\Sequence List"
'
'  Erase LIdxBuff, BookSeq, IdxBuff
'  Erase UBCustRec, UBTransRec
'
'  BlockClear
'  DisplayUBScrn "UPDATEOK"
'  WaitForAction
'
'ExitReindex:
'  UBLog "OUT: Reindex Utility Files" + CrLf$
'  Exit Sub
'
'ClearBlock:
'  HideCursor
'  Blank$ = SPACE$(40)
'  For Cnt = 8 To 15
'    QPrintRC Blank$, Cnt, 21, -1
'  Next
'  ShowCursor
'Return
'
'End Sub
'
''****************************************************************************
''Rounds a double precision value to nearest hundreth
''****************************************************************************
'Function Round#(N#)
'  Round# = Int(N# * 100 + 0.5000001) / 100
'End Function
'
'Sub Search4911Addr(S911$, RecNo&, CLSFlag%, ActiveOnly%)
'
'  Static Choice
'  ReDim ScrnArray(0)
'  SaveScrn ScrnArray()
'
'  WPos = 1
'
'  ShowProcessingScrn "Searching 911 Address."
'
''  DisplayUBScrn "SHOWSCRH"
'
'  ReDim UBCustRec(1) As NewUBCustRecType
'  UBCustRecLen = Len(UBCustRec(1))
'
'  CustFileSize& = FileSize("UBCUST.DAT")
'  NumCustRecs = CustFileSize& \ UBCustRecLen
'
'  'REDIM MChoice(1 TO 1) AS FLen2
'
'  ReDim RecBuff(1 To 1) As Long
'
'  FOpenS "UBCUST.DAT", C1Handle 'open data file
'
'  MatchCnt = 0
'  For Cnt = 1 To NumCustRecs
'    FGetRTA C1Handle, UBCustRec(1), CLng(Cnt), UBCustRecLen
'    If Not UBCustRec(1).DelFlag Then
'      If (ActiveOnly = 0) Or ((ActiveOnly = True) And ((UBCustRec(1).Status = "A"))) Then
'        GoSub ChkLoadEM
'      ElseIf (ActiveOnly = 1) And (UBCustRec(1).Status = "I") Then
'        GoSub ChkLoadEM
'      End If
'    End If
'    ShowPctComp Cnt, NumCustRecs
'  Next
'
'  FClose C1Handle
'
'  If Not DebugFlag Then
'    FreeMem& = FRE(-1)
'    If FreeMem& >= 65536 Then
'      FreeMem& = 65536
'    End If
'    MemNeeded& = MatchCnt * 64&
'    If MemNeeded& > FreeMem& Then
'      QPrintRC "Matched: " + QPStrI(MatchCnt), 24, 1, 15
'      QPrintRC " Needed: " + QPStrL(MemNeeded&) + "  Over: " + QPStrL(MemNeeded& - FreeMem&), 25, 1, 15
'      RecNo& = -1
'      WaitForAction
'      GoTo Exit911Search
'    End If
'  End If
'
'  If MatchCnt = 0 Then
'    GoTo Exit911Search
'    RecNo& = -1
'  ElseIf MatchCnt > 1 Then
'    ReDim MChoice(1 To MatchCnt) As FLen2
'    FOpenS "UBCUST.DAT", C1Handle               'open data file
'    For Cnt = 1 To MatchCnt
'      FGetRTA C1Handle, UBCustRec(1), CLng(RecBuff(Cnt)), UBCustRecLen
'      Book$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
'      LSet MChoice(Cnt).V = UBCustRec(1).Status
'      Mid$(MChoice(Cnt).V, 4) = LEFT$(QPTrim$(UBCustRec(1).CustName), 20)
'      Mid$(MChoice(Cnt).V, 26) = QPTrim$(UBCustRec(1).Addr911)
'      Mid$(MChoice(Cnt).V, 50, 9) = Book$
'      Mid$(MChoice(Cnt).V, 61) = MKL$(RecBuff(Cnt))
'    Next
'    FClose C1Handle
'
'    If DCnt = 0 Then
'      RecNo& = 0
'      GoTo Exit911Search
'    End If
'
'    'FClose L1Handle
'
'    QPrintRC "Sorting. . .   ", 12, 32, -1
'    SortT MChoice(1), MatchCnt, 0, 64, 26, 14
'
'    MaxLen = 59 'Set menu width to zero
'    Action = 0  '0 means stay in the menu until they select something
'    If Choice = 0 Then
'      Choice = 1                'Pre-load choice to highlight
'    ElseIf Choice > MatchCnt Then
'      Choice = 1                'Pre-load choice to highlight
'    End If
'    Title$ = SPACE$(MaxLen + 4)
'    LSet Title$ = " Stat    Customer          911 Address             Location"
'    '--Find max menu width
'    '--Center Menu within Screen
'    Row = 4
'    col = ((80 - 60) \ 2) - 1
'
'    If CLSFlag Then
'      Row = 4
'      BoxBot = 17               'limit the box length
'      BlockClear
'    Else
'      Row = 6
'      BoxBot = 14               'limit the box length to go no lower than line 20
'      RestScrn ScrnArray()
'    End If
'
'    LOCATE Row, col, 0
'
'    Do
'      TitleBox BoxBot + 3, col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
'      QPrintRC "Matched:" + STR$(MatchCnt), BoxBot + 4, col + 2, 15
'      QPrintRC Title$, Row - 1, col, 112
'      MPaintBox Row, col + MaxLen + 4, Row, col + MaxLen + 5, 8
'      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
'      If Ky$ = CHR$(27) Then
'        RecNo& = -1
'        Exit Do 'choice = 0
'      End If
'      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
'    Loop Until RecNo& > 0
'  Else
'    RecNo& = RecBuff(MatchCnt)
'  End If
'
'Exit911Search:
'
'  'cls
'  'Shell
'  RestScrn ScrnArray()
'  Erase ScrnArray, UBCustRec, MChoice
'  Exit Sub
'
'ChkLoadEM:
'  Cnted = Cnted + 1
'  If InStr(UBCustRec(1).Addr911, S911$) > 0 Then
'    DCnt = DCnt + 1
'    MatchCnt = MatchCnt + 1
'    ReDim Preserve RecBuff(1 To MatchCnt) As Long
'    RecBuff(MatchCnt) = Cnt
'  End If
'Return
'
'End Sub
'
'Sub Search4Cust(SEARCH$, RecNo&, CLSFlag%, ActiveOnly%)
'
'  ShowProcessingScrn "Searching Customers Info."
'
'  Static Choice, LastSEARCH$
'
'  If LastSEARCH$ <> SEARCH$ Then
'    LastSEARCH$ = SEARCH$
'    Choice = 1
'  End If
'
'  ReDim ScrnArray(0)
'  SaveScrn ScrnArray()
'
'  'DisplayUBScrn "SHOWSCRH"
'
'  ReDim MChoice(1 To 1) As FLen2
'  ReDim UBCustRec(1) As NewUBCustRecType
'
'  MaxBlockCnt = 1024            'Max Buff size: 16384   (1024*16)
'  ReDim UBCustSN(1 To MaxBlockCnt) As nUBCustReIndexRecType
'
'  UBCustRecLen = Len(UBCustRec(1))
'  UBCustSNLen = Len(UBCustSN(1))
'
'  SearchLen = Len(SEARCH$)
'  Match = False
'
''This search reads 1024 recs at a pop into a search buffer
''Basics file i/o CAN NOT read in this way.
''For a Cust file of "5000" recs
''Basic will do "5000" disk accesses.
''Our search ONLY NEEDS 5!
''5000 Recs, BlockSize 1024.
'' 4reads * 1024 = 4096 Recs
''   5000 - 4096 = 904 Odd Recs
'' 1read  * 904 Recs
'' 4 + 1 = 5 Reads!
''This is many many times faster than Basic can do.
'
'  FOpenS "UBCUSTSN.DAT", C1Handle               'open data file
'  FOpenS "UBCUST.DAT", R1Handle 'open data file
'
'  NumOfCust& = FLof&(C1Handle) / UBCustSNLen
'
''************************************
'  NumChunks& = NumOfCust& \ MaxBlockCnt
''****DO NOT CHANGE THE DIVISION HERE!
'  OddRecs = NumOfCust& Mod MaxBlockCnt
'
'  If NumChunks& = 0 Then        'if the actual cust count is less than
'    MaxBlockCnt = OddRecs       'the work buffer
'    NumChunks& = 1
'    OddRecs = 0
'  End If
'
'  BlockSize = UBCustSNLen * MaxBlockCnt
'
'  'Find matching record
'  For CCnt& = 1 To NumChunks&
'    FGetRTA C1Handle, UBCustSN(1), CCnt&, BlockSize
'    For RecCnt = 1 To MaxBlockCnt
'      UBSearchN$ = LEFT$(UBCustSN(RecCnt).SearchName, SearchLen)
'      If (SEARCH$ = UBSearchN$) Then
'        If Len(QPTrim$(UBCustSN(RecCnt).DelFlag)) Then GoTo DelSkip2
'        If (ActiveOnly = 0) Or ((ActiveOnly = True) And ((UBCustSN(RecCnt).Status = "A"))) Then
'          WhatRec& = UBCustSN(RecCnt).RecNum
'          'IF WhatRec& = 3129 THEN STOP
'          GoSub CustLoadEM2
'        ElseIf (ActiveOnly = 1) And (UBCustSN(RecCnt).Status = "I") Then
'          WhatRec& = UBCustSN(RecCnt).RecNum
'          GoSub CustLoadEM2
'        End If
'      End If
'DelSkip2:
'    Next
'    ShowPctCompL CCnt&, NumChunks&
'    'ShowSearchWheel 12, 44
'  Next
'
'  If OddRecs > 0 Then                  'this is always less than max (1024)
'    BlockSize = UBCustSNLen * OddRecs  'Adj Block size to get last chunk
'    FGetA C1Handle, UBCustSN(1), BlockSize
'    For RecCnt = 1 To OddRecs   'search'em
'      UBSearchN$ = LEFT$(UBCustSN(RecCnt).SearchName, SearchLen)
'      If (SEARCH$ = UBSearchN$) Then
'        If Len(QPTrim$(UBCustSN(RecCnt).DelFlag)) Then GoTo DelSkip3
'        If (ActiveOnly = 0) Or ((ActiveOnly = True) And ((UBCustSN(RecCnt).Status = "A"))) Then
'          WhatRec& = UBCustSN(RecCnt).RecNum
'          GoSub CustLoadEM2
'        ElseIf (ActiveOnly = 1) And (UBCustSN(RecCnt).Status = "I") Then
'          WhatRec& = UBCustSN(RecCnt).RecNum
'          GoSub CustLoadEM2
'        End If
'      End If
'DelSkip3:
'    Next
'  End If
'
'  FClose C1Handle               'close files
'  FClose R1Handle
'
'  If DCnt = 0 Then
'    RecNo& = 0
'    GoTo ExitSearch2
'  Else
'
'    SortT MChoice(1), DCnt, 0, 64, 0, 64
'    MaxLen = 59 'Set menu width to zero
'    Action = 0  '0 means stay in the menu until they select something
'
'    If Choice < 1 Then
'      Choice = 1                'Pre-load choice to highlight
'    End If
'
'    Title$ = SPACE$(MaxLen + 4)
'    LSet Title$ = "  Customer/Owner       Service Address        Location No.  S"
'    '--Find max menu width
'
'    '--Center Menu within Screen
'
'    Row = 4
'    col = ((80 - 60) \ 2) - 1
'
'    If CLSFlag Then
'      Row = 4
'      BoxBot = 17               'limit the box length
'      BlockClear
'    Else
'      Row = 6
'      BoxBot = 14               'limit the box length to go no lower than line 20
'      RestScrn ScrnArray()
'    End If
'
'    LOCATE Row, col, 0
'
'    Do
'      TitleBox BoxBot + 3, col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
'      QPrintRC "Matched:" + STR$(DCnt), BoxBot + 4, col + 2, 15
'      QPrintRC Title$, Row - 1, col, 112
'      MPaintBox Row, col + MaxLen + 4, Row, col + MaxLen + 5, 8
'      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
'      If Ky$ = CHR$(27) Then
'        RecNo& = -1
'        Exit Do 'choice = 0
'      End If
'      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
'    Loop Until RecNo& > 0
'  End If
'
'ExitSearch2:
'  RestScrn ScrnArray()
'
'  Erase ScrnArray, UBCustRec, MChoice, UBCustSN
'
'  Exit Sub
'
'CustLoadEM2:
'
'  FGetRTA R1Handle, UBCustRec(1), WhatRec&, UBCustRecLen
'
'  DCnt = DCnt + 1
'  'IF FRE(-1) < 5000 THEN STOP
'  ReDim Preserve MChoice(1 To DCnt) As FLen2
'  LSet MChoice(DCnt).V = LEFT$(QPTrim$(UBCustRec(1).CustName), 20)
'  Mid$(MChoice(DCnt).V, 22, 22) = LEFT$(QPTrim$(UBCustRec(1).SERVADDR), 25)
'  Mid$(MChoice(DCnt).V, 48, 9) = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
'  Mid$(MChoice(DCnt).V, 59, 1) = UBCustRec(1).Status
'  Mid$(MChoice(DCnt).V, 61) = MKL$(WhatRec&)
'
'Return
'
'End Sub
'
'Sub Search4LNum(LocNum$, RecNo&, CLSFlag%, ActiveOnly%)
'
'  ReDim ScrnArray(0)
'  SaveScrn ScrnArray()
'
'  ShowProcessingScrn "Searching Location Info."
'
''  DisplayUBScrn "SHOWSCRH"
'
'  ReDim UBCustRec(1) As NewUBCustRecType
'  UBCustRecLen = Len(UBCustRec(1))
'
'  IdxRecLen = 4 'we are using a integer
'  IdxFileSize& = FileSize("UBCUSTBK.IDX")
'  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
'
'  ReDim IdxBuff(1 To IdxNumOfRecs) As UBCustIndexRecType
'
'  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
'
'  SearchLen = Len(LocNum$)
'
'  Match = False
'  FirstRec = 1
'  LastRec = IdxNumOfRecs
'
'  BotOffSet = 0
'  TopOffSet = IdxNumOfRecs
'
'  FOpenS "UBCUST.DAT", C1Handle 'open data file
'  'Find matching record
'  MidRec = (LastRec + FirstRec) \ 2
'
'  Do
'    If LastSRec = MidRec Then Exit Do
'    LastSRec = MidRec
'    FGetRTA C1Handle, UBCustRec(1), CLng(IdxBuff(MidRec).RecNum), UBCustRecLen
'    UBSearchN$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
'
''    ShowSearchWheel 12, 44
'
'    If (LocNum$ = UBSearchN$) And (UBCustRec(1).DelFlag = 0) Then
'      If MidRec - BotOffSet > 1 Then
'        MidRec = MidRec - 1
'      Else
'        FirstMatchRec = MidRec
'      End If
'    ElseIf LocNum$ < UBSearchN$ Then            'lower
'      TopOffSet = MidRec
'      MidRec = TopOffSet - ((TopOffSet - BotOffSet) \ 2)
'    Else        'higher
'      BotOffSet = MidRec
'      MidRec = BotOffSet + ((TopOffSet - BotOffSet) \ 2) + 1
'      If MidRec = IdxNumOfRecs + 1 Then
'        Exit Do
'      End If
'    End If
'    If TopOffSet = BotOffSet Then Exit Do
'    'Look into this
'    'IF TopOffSet = BotOffSet THEN EXIT DO
'  Loop Until FirstMatchRec
'  ShowPctComp 1, 1
'  FClose C1Handle
'
'  If FirstMatchRec = 0 Then
'    RecNo& = 0
'  Else
'    RecNo& = IdxBuff(FirstMatchRec).RecNum
'  End If
'
'  If ActiveOnly And UBCustRec(1).Status <> "A" Then
'    RecNo& = 0
'  ElseIf (ActiveOnly = 1) And (UBCustRec(1).Status <> "I") Then
'    RecNo& = 0
'  End If
'ExitLSearch:
'
'  'cls
'  'Shell
'  RestScrn ScrnArray()
'  Erase ScrnArray, UBCustRec, IdxBuff
'End Sub
'
'Sub Search4Meter(MeterNum$, RecNo&, CLSFlag%, ActiveOnly%)
'
'  Static Choice, OMeterNum$
'
'  If OMeterNum$ <> MeterNum$ Then
'    Choice = 1
'    OMeterNum$ = MeterNum$
'  End If
'
'  ReDim ScrnArray(0)
'  SaveScrn ScrnArray()
'
'  WPos = 1
'
'  ShowProcessingScrn "Searching Meter Numbers."
'
'  ReDim UBCustRec(1) As NewUBCustRecType
'  UBCustRecLen = Len(UBCustRec(1))
'
'  CustFileSize& = FileSize("UBCUST.DAT")
'  NumCustRecs = CustFileSize& \ UBCustRecLen
'
'  'REDIM MChoice(1 TO 1) AS FLen2
'
'  ReDim RecBuff(1 To 1) As Long
'
'  FOpenS "UBCUST.DAT", C1Handle 'open data file
'  'Find matching record
'
'  MatchCnt = 0
'  For Cnt = 1 To NumCustRecs
'    'ShowSearchWheel 12, 44
'    FGetRTA C1Handle, UBCustRec(1), CLng(Cnt), UBCustRecLen
'    If Not UBCustRec(1).DelFlag Then
'      'IF NOT ActiveOnly OR (ActiveOnly AND (UBCustRec(1).Status = "A")) THEN
'      If (ActiveOnly = 0) Or ((ActiveOnly = True) And ((UBCustRec(1).Status = "A"))) Then
'        GoSub CheckEM2
'      ElseIf (ActiveOnly = 1) And (UBCustRec(1).Status = "I") Then
'        GoSub CheckEM2
'      End If
'    End If
'    ShowPctComp Cnt, NumCustRecs
'  Next
'
'  If MatchCnt = 0 Then
'    RecNo& = 0
'    FClose C1Handle
'    GoTo ExitMeterSearch
'  End If
'  If Not DebugFlag Then
'    FreeMem& = FRE(-1)
'    If FreeMem& >= 65536 Then
'      FreeMem& = 65536
'    End If
'    MemNeeded& = MatchCnt * 64&
'    If MemNeeded& > FreeMem& Then
'      FClose C1Handle
'      QPrintRC "Matched:>" + QPStrI(MatchCnt), 24, 1, 15
'      QPrintRC " Needed: " + QPStrL(MemNeeded&) + "  Free: " + QPStrL(FreeMem&), 25, 1, 15
'      RecNo& = -1
'      WaitForAction
'      GoTo ExitMeterSearch
'    End If
'  End If
'  ReDim MChoice(1 To MatchCnt) As FLen2
'
'  For Cnt = 1 To MatchCnt
'    'ShowSearchWheel 12, 44
'    FGetRTA C1Handle, UBCustRec(1), RecBuff(Cnt), UBCustRecLen
'    TCustName$ = LEFT$(QPTrim$(UBCustRec(1).CustName), 30)
'    Cnt = Cnt - 1
'    For MeterCnt = 1 To 7
'      If InStr(UBCustRec(1).LocMeters(MeterCnt).MtrNum, MeterNum$) > 0 Then
'        Cnt = Cnt + 1
'        LSet MChoice(Cnt).V = TCustName$
'        Book$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
'        Mid$(MChoice(Cnt).V, 32, 12) = UBCustRec(1).LocMeters(MeterCnt).MtrNum
'        Mid$(MChoice(Cnt).V, 50, 9) = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
'        Mid$(MChoice(Cnt).V, 61) = MKL$(RecBuff(Cnt))
'      End If
'    Next
'  Next
'
'  FClose C1Handle
'
'  If MatchCnt = 0 Then
'    RecNo& = 0
'  ElseIf MatchCnt > 1 Then
'    QPrintRC "Sorting. . .  ", 11, 34, -1
'
'    MaxLen = 59 'Set menu width to zero
'    Action = 0  '0 means stay in the menu until they select something
'    If Choice = 0 Then
'      Choice = 1                'Pre-load choice to highlight
'    ElseIf Choice > MatchCnt Then
'      Choice = 1                'Pre-load choice to highlight
'    End If
'    Title$ = SPACE$(MaxLen + 4)
'    LSet Title$ = "  Customer/Owner               Meter No.         Location No."
'
'    '--Find max menu width
'    '--Center Menu within Screen
'
'    Row = 4
'    col = ((80 - 60) \ 2) - 1
'
'    If CLSFlag Then
'      Row = 4
'      BoxBot = 17               'limit the box length
'      BlockClear
'    Else
'      Row = 6
'      BoxBot = 14               'limit the box length to go no lower than line 20
'      RestScrn ScrnArray()
'    End If
'
'    LOCATE Row, col, 0
'
'    Do
'      TitleBox BoxBot + 3, col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
'      QPrintRC "Matched:" + STR$(MatchCnt), BoxBot + 4, col + 2, 15
'      QPrintRC Title$, Row - 1, col, 112
'      MPaintBox Row, col + MaxLen + 4, Row, col + MaxLen + 5, 8
'      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
'      If Ky$ = CHR$(27) Then
'        RecNo& = -1
'        Exit Do 'choice = 0
'      End If
'      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
'    Loop Until RecNo& > 0
'  Else
'    RecNo& = CVL(MID$(MChoice(1).V, 61, 4))
'  End If
'
'ExitMeterSearch:
'
'  RestScrn ScrnArray()
'  Erase ScrnArray, UBCustRec, RecBuff, MChoice
'  Exit Sub
'
'CheckEM2:
'  For MeterCnt = 1 To 7
'    If InStr(UBCustRec(1).LocMeters(MeterCnt).MtrNum, MeterNum$) > 0 Then
'      MatchCnt = MatchCnt + 1
'      ReDim Preserve RecBuff(1 To MatchCnt) As Long
'      RecBuff(MatchCnt) = Cnt
'    End If
'  Next
'Return
'
'End Sub
'
'Sub Search4SAddr(SAddr$, RecNo&, CLSFlag%, ActiveOnly%)
'
'  Static Choice
'  ReDim ScrnArray(0)
'  SaveScrn ScrnArray()
'
'  WPos = 1
'
'  ShowProcessingScrn "Searching Service Addrs."
'
''  DisplayUBScrn "SHOWSCRH"
'
'  ReDim UBCustRec(1) As NewUBCustRecType
'  UBCustRecLen = Len(UBCustRec(1))
'
'  CustFileSize& = FileSize("UBCUST.DAT")
'  NumCustRecs = CustFileSize& \ UBCustRecLen
'
'  'REDIM MChoice(1 TO 1) AS FLen2
'
'  ReDim RecBuff(1 To 1) As Long
'
'  FOpenS "UBCUST.DAT", C1Handle 'open data file
'
'  MatchCnt = 0
'  For Cnt = 1 To NumCustRecs
''    ShowSearchWheel 12, 44
'    FGetRTA C1Handle, UBCustRec(1), CLng(Cnt), UBCustRecLen
'    If Not UBCustRec(1).DelFlag Then
'      If (ActiveOnly = 0) Or ((ActiveOnly = True) And ((UBCustRec(1).Status = "A"))) Then
'        GoSub CheckLoadEM2
'      ElseIf (ActiveOnly = 1) And (UBCustRec(1).Status = "I") Then
'        GoSub CheckLoadEM2
'      End If
'    End If
'    ShowPctComp Cnt, NumCustRecs
'  Next
'
'  FClose C1Handle
'
'  If Not DebugFlag Then
'    FreeMem& = FRE(-1)
'    If FreeMem& >= 65536 Then
'      FreeMem& = 65536
'    End If
'    MemNeeded& = MatchCnt * 64&
'    If MemNeeded& > FreeMem& Then
'      QPrintRC "Matched: " + QPStrI(MatchCnt), 24, 1, 15
'      QPrintRC " Needed: " + QPStrL(MemNeeded&) + "  Over: " + QPStrL(MemNeeded& - FreeMem&), 25, 1, 15
'      RecNo& = -1
'      WaitForAction
'      GoTo ExitSAddrSearch
'    End If
'  End If
'
'  If MatchCnt = 0 Then
'    GoTo ExitSAddrSearch
'    RecNo& = -1
'  ElseIf MatchCnt > 1 Then
'    ReDim MChoice(1 To MatchCnt) As FLen2
'    FOpenS "UBCUST.DAT", C1Handle               'open data file
'    For Cnt = 1 To MatchCnt
'      FGetRTA C1Handle, UBCustRec(1), CLng(RecBuff(Cnt)), UBCustRecLen
'      Addr$ = LEFT$(QPTrim$(UBCustRec(1).SERVADDR), 25)
'      Book$ = UBCustRec(1).Book + "-" + UBCustRec(1).SEQNUMB
'      LSet MChoice(Cnt).V = LEFT$(QPTrim$(UBCustRec(1).CustName), 20)
'      Mid$(MChoice(Cnt).V, 22, 25) = Addr$
'      Mid$(MChoice(Cnt).V, 50, 9) = Book$
'      Mid$(MChoice(Cnt).V, 61) = MKL$(RecBuff(Cnt))
'    '  ShowSearchWheel 12, 44
'    Next
'    FClose C1Handle
'
'    If DCnt = 0 Then
'      RecNo& = 0
'      GoTo ExitSAddrSearch
'    End If
'
'    'FClose L1Handle
'
'    QPrintRC "Sorting. . .   ", 12, 32, -1
'    SortT MChoice(1), MatchCnt, 0, 64, 21, 25
'
'    MaxLen = 59 'Set menu width to zero
'    Action = 0  '0 means stay in the menu until they select something
'    If Choice = 0 Then
'      Choice = 1                'Pre-load choice to highlight
'    ElseIf Choice > MatchCnt Then
'      Choice = 1                'Pre-load choice to highlight
'    End If
'    Title$ = SPACE$(MaxLen + 4)
'    LSet Title$ = "  Customer                      Address"
'    '--Find max menu width
'
'    '--Center Menu within Screen
'
'    Row = 4
'    col = ((80 - 60) \ 2) - 1
'
'    If CLSFlag Then
'      Row = 4
'      BoxBot = 17               'limit the box length
'      BlockClear
'    Else
'      Row = 6
'      BoxBot = 14               'limit the box length to go no lower than line 20
'      RestScrn ScrnArray()
'    End If
'
'    LOCATE Row, col, 0
'
'    Do
'      TitleBox BoxBot + 3, col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
'      QPrintRC "Matched:" + STR$(MatchCnt), BoxBot + 4, col + 2, 15
'      QPrintRC Title$, Row - 1, col, 112
'      MPaintBox Row, col + MaxLen + 4, Row, col + MaxLen + 5, 8
'      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
'      If Ky$ = CHR$(27) Then
'        RecNo& = -1
'        Exit Do 'choice = 0
'      End If
'      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
'    Loop Until RecNo& > 0
'  Else
'    RecNo& = RecBuff(MatchCnt)
'  End If
'
'ExitSAddrSearch:
'
'  'cls
'  'Shell
'  RestScrn ScrnArray()
'  Erase ScrnArray, UBCustRec, MChoice
'  Exit Sub
'
'CheckLoadEM2:
'  Cnted = Cnted + 1
'  If InStr(UBCustRec(1).SERVADDR, SAddr$) > 0 Then
'    DCnt = DCnt + 1
'    MatchCnt = MatchCnt + 1
'    ReDim Preserve RecBuff(1 To MatchCnt) As Long
'    RecBuff(MatchCnt) = Cnt
'  End If
'Return
'
'End Sub
'
'Sub ShowCustConsHist(CustRec&)
'
'  ReDim TempScrn(0)
'  SaveScrn TempScrn()
'
'  ReDim Metered(1 To 15)
'  ReDim UBSetUpRec(1) As UBSetupRecType
'  UBSetupLen = Len(UBSetUpRec(1))
'  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
'  If InStr(UBSetUpRec(1).UTILNAME, "TROY") > 0 Then
'    TroyFlag = True
'  End If
'  If InStr(UBSetUpRec(1).UTILNAME, "HAMLET") > 0 Then
'    HamFlag = True
'  End If
'
'  NumOfRevs = MaxRevsCnt
'  For RevCnt = 1 To 15
'    RLen = Len(QPTrim$(LEFT$(UBSetUpRec(1).Revenues(RevCnt).REVNAME, 14)))
'    If RLen >= 0 Then
'      NumOfRevs = RevCnt - 1
'      Exit For
'    End If
'    If UBSetUpRec(1).Revenues(RevCnt).UseMtr = "Y" Then
'      Metered(RevCnt) = True
'    End If
'  Next
'
'  ReDim MChoice(1 To 1) As FLen2
'  ReDim UBTranRec(1) As UBTransRecType
'  ReDim UBCustRec(1) As NewUBCustRecType
'
'  UBCustRecLen = Len(UBCustRec(1))
'  UBTranRecLen = Len(UBTranRec(1))
'
'  UBFile = FREEFILE
'  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
'  Get UBFile, CustRec&, UBCustRec(1)
'  Close UBFile
'
'  CurBal# = UBCustRec(1).CurrBalance
'  PreBal# = UBCustRec(1).PrevBalance
'
'  UBTran = FREEFILE
'  Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen
'
'  PrevTranRec& = UBCustRec(1).LastTrans
'
'  If PrevTranRec& > 0 Then
'    Do While PrevTranRec& > 0
'      Get UBTran, PrevTranRec&, UBTranRec(1)
'      If UBTranRec(1).TransType = TranUtilityBill Or UBTranRec(1).TransType = TranUtilityBill + 100 Then
'        For MtrCnt = 1 To 7
'          If UBTranRec(1).MtrTypes(MtrCnt) <> 0 Then
'            DCnt = DCnt + 1
'            ReDim Preserve MChoice(1 To DCnt) As FLen2
'            If HamFlag Then
'              LSet MChoice(DCnt).V = Num2Date(UBTranRec(1).ReadDate)
'            Else
'              LSet MChoice(DCnt).V = Num2Date(UBTranRec(1).TransDate)
'            End If
'            Select Case UBTranRec(1).MtrTypes(MtrCnt)
'            Case MtrWaterOnly
'              MeterType$ = "Water"
'            Case MtrSewerOnly
'              MeterType$ = "Sewer"
'            Case MtrCombined
'              MeterType$ = "Combined"
'            Case MtrElectric
'              MeterType$ = "Electric"
'            Case MtrDemand
'              MeterType$ = "D Electric"
'            Case MtrGas
'              MeterType$ = "Gas Meter"
'            Case MtrTouchRead
'              MeterType$ = "Touch Read"
'            Case MtrLightsService
'              MeterType$ = "L Service"
'            Case -1
'              MeterType$ = "L Service"
'            End Select
'
'            Mid$(MChoice(DCnt).V, 13) = MeterType$
'            Mid$(MChoice(DCnt).V, 26) = FUsing$(STR$(UBTranRec(1).CurRead(MtrCnt)), "##########")
'            Mid$(MChoice(DCnt).V, 38) = FUsing$(STR$(UBTranRec(1).PrevRead(MtrCnt)), "##########")
'            MeterConsp& = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
'            If MeterConsp& < 0 Then
'              MaxMeterAmt& = 10& ^ (Len(STR$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
'              MeterConsp& = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
'            End If
''working here
'            MTRMulti# = 0
'            For MCnt = 1 To 7
'              If UBTranRec(1).MtrTypes(MtrCnt) = GetCustMeterType%(UBCustRec(), MCnt) Then
'                MTRMulti# = UBCustRec(1).LocMeters(MCnt).MTRMulti
'                If UBCustRec(1).LocMeters(MCnt).MTRUnit = "C" Then
'                  MeterConsp& = MeterConsp& * 7.481
'                  Exit For
'                End If
'              End If
'            Next
'            If MTRMulti# = 0 Then
'              If TroyFlag Then
'                MTRMulti# = 100
'              Else
'                MTRMulti# = 1
'              End If
'            End If
'            Mid$(MChoice(DCnt).V, 52) = FUsing$(STR$(MTRMulti# * MeterConsp&), "##########")
'          End If
'        Next
'      End If
'      PrevTranRec& = UBTranRec(1).PrevTrans
'    Loop
'
'    Close UBTran
'    RestScrn TempScrn()
'    MPaintBox 3, 5, 22, 75, 8
'
'    MaxLen = 62 'Set menu width to zero
'    Action = 0  '0 means stay in the menu until they select something
'
'    If Choice < 1 Then
'      Choice = 1                'Pre-load choice to highlight
'    End If
'
'    Title$ = SPACE$(MaxLen + 4)
'    Balance$ = Title$
'    LSet Title$ = " Trans Date   Meter Type      Current    Previous    Consumption"
'
'    '--Find max menu width
'    '--Center Menu within Screen
'
'    Row = 4
'    col = 8
'    Row = 6
'    BoxBot = 17 'limit the box length to go no lower than line 20
'
'    TitleBox BoxBot + 3, col, MaxLen + 3, "Press <ESC> to continue.", Cnf
'
'    QPrintRC Title$, Row - 1, col, 112
'    MPaintBox Row, col + MaxLen + 4, Row, col + MaxLen + 5, 8
'
'    Do
'      LOCATE Row, col, 0
'      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
'      If Ky$ = CHR$(27) Then
'        RestScrn TempScrn()
'        Exit Do
'      End If
'    Loop
'  Else
'    Close UBTran
'    Ok = MsgBox%("UB.QSL", "NOCTRANS")
'    RestScrn TempScrn()
'  End If
'
'  RestScrn TempScrn()
'  Erase Metered, UBSetUpRec, MChoice
'  Erase TempScrn, UBTranRec, UBCustRec
'
'Exit Sub
'
'
'End Sub
'
'Sub ShowCustHistory(CustRec&)
'
'  u$ = CHR$(24)
'  d$ = CHR$(25)
'
'  ReDim TempScrn(0)
'  SaveScrn TempScrn()
'
'  DisplayUBScrn "UBCUHIST"
'
'  ReDim RevText$(1 To MaxRevsCnt)
'  ReDim Metered(1 To 15)
'  ReDim UBSetUpRec(1) As UBSetupRecType
'  UBSetupLen = Len(UBSetUpRec(1))
'  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
'  NumOfRevs = MaxRevsCnt
'  For RevCnt = 1 To 15
'    RevText$(RevCnt) = LEFT$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME), 14)
'    If Len(RevText$(RevCnt)) = 0 Then
'      NumOfRevs = RevCnt - 1
'      Exit For
'    End If
'    If UBSetUpRec(1).Revenues(RevCnt).UseMtr = "Y" Then
'      Metered(RevCnt) = True
'    End If
'  Next
'
'  ReDim MChoice(1 To 1) As FLen2
'
'  ReDim UBTranRec(1) As UBTransRecType
'  ReDim UBCustRec(1) As NewUBCustRecType
'
'  UBCustRecLen = Len(UBCustRec(1))
'  UBTranRecLen = Len(UBTranRec(1))
'
'  UBFile = FREEFILE
'  Open "UBCUST.DAT" For Random Shared As UBFile Len = UBCustRecLen
'  Get UBFile, CustRec&, UBCustRec(1)
'  Close UBFile
'
'  CurBal# = UBCustRec(1).CurrBalance
'  PreBal# = UBCustRec(1).PrevBalance
'
'Top:
'
'  UBTran = FREEFILE
'  Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen
'
'  PrevTranRec& = UBCustRec(1).LastTrans
'
'  If PrevTranRec& > 0 Then
'    Do While PrevTranRec& > 0
'      DCnt = DCnt + 1
'      ReDim Preserve MChoice(1 To DCnt) As FLen2
'      Get UBTran, PrevTranRec&, UBTranRec(1)
'      LSet MChoice(DCnt).V = Num2Date(UBTranRec(1).TransDate)
'      'MID$(MChoice(DCnt).V, 15) = UBTranRec(1).TransDesc
'      GoSub GetTransType
'      Mid$(MChoice(DCnt).V, 13) = TType$
'      Mid$(MChoice(DCnt).V, 41) = FUsing(STR$(UBTranRec(1).TransAmt), "#####.##")
'      'this will show th actual trans number in the list
'      'MID$(MChoice(DCnt).V, 50) = FUsing(STR$(PrevTranRec&), "######")
'      Mid$(MChoice(DCnt).V, 52) = FUsing(STR$(UBTranRec(1).RunBalance), "#####.##")
'      Mid$(MChoice(DCnt).V, 61) = MKL$(PrevTranRec&)
'      PrevTranRec& = UBTranRec(1).PrevTrans
'    Loop
'
'    Close UBTran
'
'    RestScrn TempScrn()
'    MPaintBox 3, 5, 22, 75, 8
'    ReDim TempScrn2(0)
'    SaveScrn TempScrn2()
'
'HistTop:
'
'    MaxLen = 59 'Set menu width to zero
'    Action = 0  '0 means stay in the menu until they select something
'
'    If Choice < 1 Then
'      Choice = 1                'Pre-load choice to highlight
'    End If
'
'    Title$ = SPACE$(MaxLen + 4)
'    Balance$ = Title$
'    LSet Title$ = "  Trans Date       Description           Trans Amt    Balance  "
'    LSet Balance$ = " Balance:" + FUsing(STR$(CurBal# + PreBal#), ",#####.##") + "   Cur:" + FUsing(STR$(CurBal#), ",#####.##") + "  Prev:" + FUsing(STR$(PreBal#), ",#####.##")
'
'    '--Find max menu width
'    '--Center Menu within Screen
'
'    Row = 4
'    col = ((80 - 60) \ 2) - 1
'
'    Row = 6
'    BoxBot = 17 'limit the box length to go no lower than line 20
'
'    'TitleBox BoxBot + 3, Col, MaxLen + 3, "       Press <ESC> to continue.", Cnf
'
'    WazzWind BoxBot + 2, col, BoxBot + 5, MaxLen + 3 + col, 10, 4, True
'    QPrintRC "  Use:  " + u$ + "-" + d$ + " to select.", BoxBot + 3, col + 3, 15
'    QPrintRC u$, BoxBot + 3, col + 11, 14
'    QPrintRC d$, BoxBot + 3, col + 13, 14
'
'    QPrintRC "Total: " + STR$(DCnt), BoxBot + 4, col + 3, 15
'    QPrintRC "Press:   [ESC] to continue.", BoxBot + 3, col + 33, 15
'    QPrintRC "        [ENTER] for detail.", BoxBot + 4, col + 33, 15
'    QPrintRC "ESC", BoxBot + 3, col + 43, 14
'    QPrintRC "ENTER", BoxBot + 4, col + 42, 14
'
'    QPrintRC Balance$, Row - 2, col, 112
'    QPrintRC Title$, Row - 1, col, 112
'    MPaintBox Row, col + MaxLen + 4, Row, col + MaxLen + 5, 8
'    'FirstTime = True
'
'    'SLEEP
'
'    Do
'      LOCATE Row, col, 0
'      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
'      If Ky$ = CHR$(27) Then
'        RestScrn TempScrn()
'        Exit Do 'choice = 0
'      ElseIf Ky$ = CHR$(13) Then
'        RestScrn TempScrn2()
'        GoTo ShowTransDetail
'      End If
'    Loop        'UNTIL EditLocRec& > 0
'  Else
'    Close UBTran
'    Ok = MsgBox%("UB.QSL", "NOCTRANS")
'    RestScrn TempScrn()
'  End If
'
'  RestScrn TempScrn()
'  Erase RevText$, Metered, UBSetUpRec, MChoice
'  Erase TempScrn, UBTranRec, UBCustRec
'
'  Exit Sub
'
'ShowTransDetail:
'  CursorOff
'  TransRecNum& = CVL(RIGHT$(MChoice(Choice).V, 4))
'  UBTran = FREEFILE
'  Open "UBTRANS.DAT" For Random Shared As UBTran Len = UBTranRecLen
'  Get UBTran, TransRecNum&, UBTranRec(1)
'  Close UBTran
'
'  DisplayUBScrn "TRDETAIL"
'
'  QPrintRC Num2Date(UBTranRec(1).TransDate), 3, 23, 15
'
'  'CONST TranUtilityBill = 1          '   1=Utility bill
'  'CONST TranLateCharge = 2           '   2=late charge
'  'CONST TranReconnectFee = 3         '   3=reconnect fee
'  'CONST TranBillPayment = 4          '   4=Bill Payment
'  'CONST TranAppliedDeposit = 5       '   5=Applied Deposit
'  'CONST TranPenaltyCharge = 6        '   6=Penalty Charge
'  'CONST TranDepositPayment = 7       '   7=Deposit Payment
'  'CONST TranDraftPayment = 8         '   8=Draft Payment
'  'CONST TranRefundDeposit = 9       '    9=Refund Deposit
'  'CONST TranBeginBalance = 10        '  10=Beginning Balance
'  'CONST TranUpwardAdjustment = 11    '  11=Bill Adjustments
'  'CONST TranDownwardAdjustment = 12  '  12=Bill Adjustments
'
'  GoSub GetTransType
'
'  QPrintRC FUsing$(STR$(UBTranRec(1).TransAmt), "#####.##"), 4, 25, 15
'
'  QPrintRC TType$, 4, 50, 15
'  QPrintRC UBTranRec(1).TransDesc, 3, 50, 15
'
'  For RevCnt = 1 To NumOfRevs
'    QPrintRC RevText$(RevCnt), RevCnt + 6, 8, 15
'    QPrintRC FUsing$(STR$(UBTranRec(1).RevAmt(RevCnt)), "#####.##"), RevCnt + 6, 25, 15
'    QPrintRC FUsing$(STR$(UBTranRec(1).TaxAmt(RevCnt)), "###.##"), RevCnt + 6, 36, 15
'    '(Number$, Image$)
'  Next
'
'  For Cnt = 1 To 7
'    If Metered(Cnt) Then
'      QPrintRC FUsing$(STR$(UBTranRec(1).CurRead(Cnt)), "#########"), Cnt + 6, 42, 15
'      QPrintRC FUsing$(STR$(UBTranRec(1).PrevRead(Cnt)), "#########"), Cnt + 6, 53, 15
'      If QPTrim$(UBTranRec(1).ESTREAD(Cnt)) = "" Then
'        QPrintRC "N", Cnt + 6, 70, 15
'      Else
'        QPrintRC "Y", Cnt + 6, 70, 15
'      End If
'    End If
'  Next
'
'  WaitForAction
'  RestScrn TempScrn2()
'  GoTo HistTop
'
'GetTransType:
'
'  Select Case UBTranRec(1).TransType
'  Case TranUtilityBill, TranUtilityBill + 100
'    TType$ = "Utility Bill "
'  Case TranLateCharge, TranReconnectFee, TranLateCharge + 100, TranReconnectFee + 100
'    TType$ = "Penalty, Reconnect Fee"
'  Case TranBillPayment, TranBillPayment + 100
'    TDesc$ = QPTrim$(UBTranRec(1).TransDesc)
'    If InStr(UBTranRec(1).TransDesc, "PAYMENT") = 0 And Len(TDesc$) > 0 Then
'      TType$ = "Utility Payment " + LEFT$(QPTrim$(UBTranRec(1).TransDesc), 10)
'    Else
'      TType$ = "Utility Payment"
'    End If
'  Case TranPenaltyPayment
'    TType$ = "Penalty Payment"
'  Case TranPenaltyCharge
'    TType$ = "Penalty/Late Fee"
'  Case TranAppliedDeposit
'    TType$ = "Applied Deposit"
'  Case TranDepositPayment, TranDepositPayment + 100
'    TDesc$ = QPTrim$(UBTranRec(1).TransDesc)
'    If InStr(UBTranRec(1).TransDesc, "DEPOSIT") = 0 And Len(TDesc$) > 0 Then
'      TType$ = "Deposit Payment " + LEFT$(QPTrim$(UBTranRec(1).TransDesc), 10)
'    Else
'      TType$ = "Deposit Payment"
'    End If
'  Case TranDraftPayment
'    TType$ = "Draft Payment"
'  Case TranBeginBalance, TranBeginBalance + 100
'    TType$ = "Beginning Balance"
'  Case 9
'    TType$ = "Deposit Refund"
'  Case TranUpwardAdjustment
'    TType$ = "Upward Adjustment"
'  Case TranDownwardAdjustment
'    TType$ = "Downward Adjustment"
'  Case Else
'    TType$ = STR$(UBTranRec(1).TransType) + " ???"
'  End Select
'
'Return
'
'End Sub
'
'Static Sub ShowPctComp(ByVal RecNo%, ByVal NumOfRecs%)
'  RSet PctC(1) = QPStrI$(Int((RecNo / NumOfRecs) * 100))
'  'HideCursor
'  QPrintRC PctC(1), 13, 40, Cnf.HiLite
'  'ShowCursor
'  '  QPrintRC STR$(FRE("")), 25, 1, Cnf.HiLite
'End Sub
'
'Static Sub ShowPctCompL(ByVal RecNo&, ByVal NumOfRecs&)
'  RSet PctC(1) = QPStrL$(Int((RecNo& / NumOfRecs&) * 100))
'  'HideCursor
'  QPrintRC PctC(1), 13, 40, Cnf.HiLite
'  'ShowCursor
'  '  QPrintRC STR$(FRE("")), 25, 1, Cnf.HiLite
'End Sub
'
'Sub ShowProcessingScrn(RptTitle$)
'  TitleRow = 9
'  TitleCol = 40 - (Len(RptTitle$) \ 2) + 1
'  CursorOff
'  BlockClear
'  DisplayUBScrn "PRORPT"
'  HideCursor
'  QPrintRC RptTitle$, TitleRow, TitleCol, 126
'  QPrintRC "Processing:    % Completed.", 13, 28, Cnf.HiLite
'  ShowCursor
'End Sub
'
