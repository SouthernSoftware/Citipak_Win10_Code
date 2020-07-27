DEFINT A-Z
DECLARE SUB MakeZipCodeIndex (IndexText$)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE FUNCTION GetZipEDigit$ (Zip$)
DECLARE FUNCTION IsDeleted% (AcctNo&)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB AddEditLocation (RecNo&, FromFlag%)
DECLARE SUB ShowSearchWheel (BYVAL Row%, BYVAL Col%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB BlockClear ()
DECLARE SUB ButtonPress (ButNo%, Down%, Presses%, X%, Y%)
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayScrn (BTmp%(), Element%, MonoCode%, WipeType%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB ExplodBox (UlRow%, UlCol%, BrRow%, BrCol%, Scr(), El%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB GetCursor (X, Y, Button)
DECLARE SUB HideCursor ()
DECLARE SUB BlockClear ()
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB MScrnRest (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB MScrnSave (UlRow, UlCol, LRRow, LRCol, SEG Address)
DECLARE SUB Pause3 (MillaSecs%, ProcessorSpeed&)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
'DECLARE SUB MQPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB ShowCursor ()
DECLARE SUB SaveScrn (Array())
DECLARE SUB RestScrn (Array())
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)
  
DECLARE FUNCTION Chk4DupeBookSeqNum (BOOK$, SeqNum$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION QPValI% (Number$)
  
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION FmtBook$ (BOOK$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION FindRateTbl% (RATECODE$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION FileSize& (FileName$)
  
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION ConvDateStr$ (D$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FudgeFactor& ()
DECLARE FUNCTION NovellThere% ()
DECLARE FUNCTION QPStrI$ (Num%)
DECLARE FUNCTION QPStrL$ (Num&)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (N#)
DECLARE SUB Search4Cust (SEARCH$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB Search4LNum (LNum$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB Search4Meter (MeterNum$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB Search4SAddr (SAddr$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)

  TYPE FLen2
    V AS STRING * 64
  END TYPE
  
  '$INCLUDE: 'DefCnf.bi'      'Defines a TYPE for monitor and color info..
  '$INCLUDE: 'FORMEDIT.bi'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'UBTRANS.bi'
  '$INCLUDE: 'UBSETUP.bi'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'setcnf.bi'
  
  CONST False = 0, True = NOT False
  CONST Wheel$ = "|/Ä\"
  
  DIM SHARED PctC(1) AS STRING * 3
  DIM SHARED DebugFlag AS INTEGER

FUNCTION AskAbandonPrint%
  
  STATIC BeenHere, Escape$
  
  IF NOT BeenHere THEN
    BeenHere = True
    Escape$ = CHR$(27)
  END IF
  
  Ky$ = INKEY$  'ButNo,DnFlg,NoPresses,col,row
  ButtonPress 1, N, MooseButton, MCol, MRow     ' ----- Check Mouse activity
  IF N AND 2 THEN               'if it was the right moose button and
    DO          'if they are still holding it down then
      GetCursor 0, 0, NewStatus 'wait till they let go of the button
    LOOP WHILE NewStatus
    ButtonPress 0, 0, 0, 0, 0   'this has the effect of clearing all
    Ky$ = Escape$               'moose buttons.
  END IF
  
  IF Ky$ <> Escape$ THEN
    EXIT FUNCTION
  END IF
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  ActMRow = 14
  BlockClear
  DisplayUBScrn "ABORTJOB"
  
  DO
    
    Get.Moose.OR.Key Ky$, MooseButton, MRow, MCol
    
    IF MooseButton THEN
      MRow = (MRow \ 8) + 1     'Convert MCol, MRow to Row and Col.
      MCol = (MCol \ 8) + 1
      IF MRow = ActMRow THEN
        SELECT CASE MCol
        CASE 28 TO 39           'Cancel printing
          PressButton EscKey, ActMRow, 28, 39
        CASE 42 TO 55           'let it rip
          PressButton 13, ActMRow, 42, 55
        END SELECT
      END IF
    END IF
    
    IF LEN(Ky$) THEN
      IF Ky$ = Escape$ THEN
        AskAbandonPrint% = True
      ELSE
        AskAbandonPrint% = False
      END IF
      EXIT DO
    END IF
  LOOP
  
  RestScrn TempScrn()
  ERASE TempScrn
  
END FUNCTION

FUNCTION Chk4DupeBookSeqNum (BOOK$, SeqNum$)
  
  Chk4DupeBookSeqNum = False    'assume it's ok
  
  TBookSeq& = QPValL(BOOK$ + SeqNum$)
  
  REDIM UBBookSeq(1) AS BookSeqRecType
  BookSeqLen = LEN(UBBookSeq(1))
  
  IF FileSize("UBOOKSEQ.DAT") > 0 THEN
    FOpenS "UBOOKSEQ.DAT", Handle               'open data file
    NumBookSeq = FLof(Handle) \ BookSeqLen
    REDIM UBBookSeq(1 TO NumBookSeq) AS BookSeqRecType
    FGetRTA Handle, UBBookSeq(1), 1&, NumBookSeq * BookSeqLen
    FClose Handle
    
    FOR Cnt = 1 TO NumBookSeq
      IF UBBookSeq(Cnt).BookSeq = TBookSeq& THEN
        Ok = MsgBox%("UB.QSL", "DUPEBOOK")
        Chk4DupeBookSeqNum = True
        EXIT FOR
      END IF
    NEXT
  END IF
  
END FUNCTION

SUB CustMessageSystem (RecNo&)

  CustRec& = RecNo&

  REDIM ScrnArray(0)

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBMessRec(1) AS UBMessRecType
  UBMessRecLen = LEN(UBMessRec(1))

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, CustRec&, UBCustRec(1)
  CLOSE
  
  LibName$ = "UB"
  ScrnName$ = "UBCUSMES"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  Frm(1).StayOnField = True

  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

  FirstTime = True

  Action = 1

  DisplayUBScrn ScrnName$
  QPrintRC STR$(CustRec&), 3, 20, -1
  QPrintRC UBCustRec(1).CUSTNAME, 4, 20, -1
  QPrintRC UBCustRec(1).STATUS, 3, 67, -1

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      GOSUB LoadMessageInfo
      Action = 1
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F3Key
      GOSUB ClearRecord
      GOSUB ClearForm
      Action = 1
    CASE F10Key
      SaveScrn ScrnArray()
      DisplayUBScrn "UPDATDSK"
      GOSUB SaveRecord
      RestScrn ScrnArray()
      DisplayUBScrn "UPDATEOK"
      WaitForAction
      ExitFlag = True
      RestScrn ScrnArray()
      Done = True
    CASE ESC
      EXIT SUB
    CASE ELSE
      Done = False
    END SELECT
  LOOP UNTIL Done

ExitMessageInquiry:
  EXIT SUB
'***************

LoadMessageInfo:
  MessageRecord = UBCustRec(1).MessageRec
  IF MessageRecord > 0 THEN
    UBMess = FREEFILE
    OPEN "UBMESAGE.DAT" FOR RANDOM SHARED AS UBMess LEN = UBMessRecLen
    GET UBMess, MessageRecord, UBMessRec(1)
    CLOSE
    BCopy VARSEG(UBMessRec(1)), VARPTR(UBMessRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), UBMessRecLen, 0
    CALL UnPackBuffer(0, 0, Form$(), Fld())
  END IF
RETURN

SaveRecord:
  UBMess = FREEFILE
  OPEN "UBMESAGE.DAT" FOR RANDOM SHARED AS UBMess LEN = UBMessRecLen
  IF MessageRecord = 0 THEN
    MessageRecord = LOF(UBMess) / LEN(UBMessRec(1)) + 1
  END IF

  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBMessRec(1)), VARPTR(UBMessRec(1)), UBMessRecLen, 0
  PUT UBMess, MessageRecord, UBMessRec(1)
  CLOSE

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET UBCust, CustRec&, UBCustRec(1)
  UBCustRec(1).MessageRec = MessageRecord
  PUT UBCust, CustRec&, UBCustRec(1)
  CLOSE

RETURN

ClearRecord:
  IF MessageRecord > 0 THEN
    REDIM UBMessRec(1) AS UBMessRecType
    UBMess = FREEFILE
    OPEN "UBMESAGE.DAT" FOR RANDOM SHARED AS UBMess LEN = UBMessRecLen
    PUT UBMess, MessageRecord, UBMessRec(1)
    CLOSE
  END IF
RETURN

ClearForm:
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
RETURN


END SUB

SUB DisplayUBScrn (ScrnName$)
  LibFile2Scrn "UB", ScrnName$, MonoCode, Attribute%, ErrCode
END SUB

FUNCTION FmtBook$ (BOOK$)
  BOOK$ = QPTrim$(BOOK$)
  BookLen = LEN(BOOK$)
  
  SELECT CASE BookLen
  CASE 0
    FmtBook$ = "00"
  CASE 1
    FmtBook$ = "0" + BOOK$
  CASE ELSE
    FmtBook$ = BOOK$
  END SELECT
  
END FUNCTION

FUNCTION FmtSeqN$ (SeqN$)
  
  SeqN$ = QPTrim$(SeqN$)
  SeqNLen = LEN(SeqN$)
  
  SELECT CASE SeqNLen
  CASE 0
    FmtSeqN$ = "000000"
  CASE 1 TO 5
    FmtSeqN$ = "000000"
    MID$(FmtSeqN$, (6 - SeqNLen) + 1) = SeqN$
  CASE ELSE
    FmtSeqN$ = SeqN$
  END SELECT
  
END FUNCTION

  'This function returns the number of customer records
FUNCTION GetNumOfCust
  REDIM TCustRec(1) AS NewUBCustRecType
  RecLen = LEN(TCustRec(1))
  CFileSize& = FileSize("UBCUST.DAT")
  GetNumOfCust = CFileSize& \ RecLen
  ERASE TCustRec
END FUNCTION

FUNCTION GetNumRateRecs
  REDIM UBRateTblRec(1) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))
  GetNumRateRecs = FileSize("UBRATE.DAT") \ UBRateTblRecLen
  ERASE UBRateTblRec
END FUNCTION

FUNCTION GetZipEDigit$ (Zip$)

  ZipLen = LEN(Zip$)
  ZipVal = 0

  DashPos = INSTR(Zip$, "-")
  DO WHILE DashPos
    Zip$ = LEFT$(Zip$, DashPos - 1) + MID$(Zip$, DashPos + 1)
  LOOP

  FOR Cnt = 1 TO ZipLen
    ZipVal = ZipVal + VAL(MID$(Zip$, Cnt, 1))
  NEXT

  Dif = 10 - (ZipVal MOD 10)

  GetZipEDigit$ = QPTrim$(STR$(Dif))

END FUNCTION

FUNCTION IsDeleted% (AcctNo&)
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  FOpenS "UBCUST.DAT", C1Handle 'open data file
  FGetRTA C1Handle, UBCustRec(1), AcctNo&, UBCustRecLen
  FClose C1Handle
  IF UBCustRec(1).DelFlag <> 0 THEN
    IsDeleted% = True
  ELSE
    IsDeleted% = False
  END IF
  ERASE UBCustRec
END FUNCTION

SUB LoadUBSetUpFile (UBSetUpRec() AS UBSetupRecType, UBSetupLen)
  
  UBSetupLen = -1               'use the length as an error flag
  'assume the file is not there, or 0 bytes.
  IF Exist("UBSETUP.DAT") THEN
    FOpenS "UBSETUP.DAT", Handle                'open data file
    IF FLof&(Handle) > 0 THEN
      UBSetupLen = LEN(UBSetUpRec(1))
      FGetRTA Handle, UBSetUpRec(1), 1&, UBSetupLen
    END IF
    FClose Handle
  END IF
  
END SUB

SUB MakePostalIndex (IndexText$)
  ShowProcessingScrn "Creating " + IndexText$ + " Index"
  QPrintRC "    Reading Customer Records     ", 11, 25, -1
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  CustRecLen = LEN(UBCustRec(1))
  
  NumCustRecs = GetNumOfCust%
  
  REDIM PostalIndex(1 TO NumCustRecs)  AS UBPostalIndexType
  IndexRecLen = LEN(PostalIndex(1))
  
  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = CustRecLen
  FOR Cnt = 1 TO NumCustRecs
    GET CHandle, Cnt, UBCustRec(1)
    PostalIndex(Cnt).ZipCode = UBCustRec(1).ZipCode
    PostalIndex(Cnt).Route = UBCustRec(1).POSTRTE
    PostalIndex(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumCustRecs                'show user percentage complete
  NEXT
  
  CLOSE CHandle
  
  QPrintRC "         Sorting Index.        ", 11, 25, -1
  
  SortT PostalIndex(1), NumCustRecs, 0, 16, 10, 4
  
  SortT PostalIndex(1), NumCustRecs, 0, 16, 0, 14
  
  'SortT (Array(1), NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  
  QPrintRC "      Writing Index Records      ", 11, 25, -1
  IHandle = FREEFILE
  
  FCreate TempIndexName
  
  OPEN TempIndexName FOR RANDOM SHARED AS IHandle LEN = 4
  FOR Cnt = 1 TO NumCustRecs
    Prec& = PostalIndex(Cnt).RecNum
    PUT IHandle, Cnt, Prec&
    ShowPctComp Cnt, NumCustRecs                'show user percentage complete
  NEXT
  CLOSE IHandle
  
  ERASE UBCustRec, PostalIndex
  
END SUB

SUB MakeSequenceIndex (IndexText$)
  ShowProcessingScrn "Creating " + IndexText$ + " Index"
  QPrintRC "    Reading Location Records     ", 11, 25, -1
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  CustRecLen = LEN(UBCustRec(1))
  
  NumCustRecs& = GetNumOfCust%
  
  REDIM SequenceIndex(1 TO NumCustRecs&)  AS UBSequenceIndexType
  IndexRecLen = LEN(SequenceIndex(1))
  
  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = CustRecLen
  FOR Cnt = 1 TO NumCustRecs&
    GET CHandle, Cnt, UBCustRec(1)
    SequenceIndex(Cnt).SeqNumber = UBCustRec(1).SEQ
    SequenceIndex(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumCustRecs&               'show user percentage complete
  NEXT
  
  CLOSE CHandle
  
  QPrintRC "         Sorting Index.        ", 11, 25, -1
  
  SortT SequenceIndex(1), CINT(NumCustRecs&), 0, 16, 0, -2
  
  QPrintRC "      Writing Index Records      ", 11, 25, -1
  
  FCreate TempIndexName
  
  IHandle = FREEFILE
  OPEN TempIndexName FOR RANDOM SHARED AS IHandle LEN = 4
  
  FOR Cnt = 1 TO NumCustRecs&
    Prec& = SequenceIndex(Cnt).RecNum
    PUT IHandle, Cnt, Prec&
    ShowPctComp Cnt, NumCustRecs&              'show user percentage complete
  NEXT
  CLOSE IHandle
  
END SUB

SUB MakeZipCodeIndex (IndexText$)

  ShowProcessingScrn "Creating " + IndexText$ + " Index "
  QPrintRC "    Reading Customer Records     ", 11, 25, -1
                                 
  'REDIM ZipIndex(1 TO 1)  AS PSAZipIndexType
  REDIM UBCustRec(1)      AS NewUBCustRecType
  CustRecLen = LEN(UBCustRec(1))

  NumOfBillRec = FileSize("UBCUST.DAT") \ CustRecLen

  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = CustRecLen
  'shell
  REDIM ZipIndex(1 TO NumOfBillRec)   AS PSAZipIndexType

  FOR BCnt = 1 TO NumOfBillRec
    GET CHandle, BCnt, UBCustRec(1)
    ZipIndex(BCnt).ZipCode = UBCustRec(1).ZipCode
    ZipIndex(BCnt).SName = UBCustRec(1).SEARCH
    ZipIndex(BCnt).RecNum = BCnt
    ShowPctComp BCnt, NumOfBillRec        'show user percentage complete
  NEXT

  CLOSE

  QPrintRC "         Sorting Index.        ", 11, 25, -1
  
  SortT ZipIndex(1), NumOfBillRec, 0, 32, 0, 10
  
  First = 1
  Last = 1

  SZip$ = ZipIndex(1).ZipCode

  FOR ZCnt = 2 TO NumOfBillRec
    EZip$ = ZipIndex(ZCnt).ZipCode
    IF SZip$ <> EZip$ THEN
      Last = ZCnt - 1
      GOSUB SortThisZip
      First = ZCnt
      SZip$ = EZip$
    END IF
    ShowPctComp ZCnt, NumOfBillRec    'show user percentage complete
  NEXT
  Last = ZCnt - 1
  GOSUB SortThisZip
  
  QPrintRC "      Writing Index Records      ", 11, 25, -1

  IHandle = FREEFILE
  OPEN TempIndexName FOR OUTPUT AS IHandle
  CLOSE IHandle

  IHandle = FREEFILE
  OPEN TempIndexName FOR RANDOM SHARED AS IHandle LEN = 4
  FOR Cnt = 1 TO NumOfBillRec
    Prec& = ZipIndex(Cnt).RecNum
    PUT IHandle, Cnt, Prec&
    ShowPctComp Cnt, NumOfBillRec       'show user percentage complete
  NEXT
  CLOSE IHandle
  
  ERASE UBCustRec, ZipIndex

EXIT SUB

SortThisZip:
  IF First < Last THEN
    SortT ZipIndex(First), Last - First + 1, 0, 32, 10, 10
  END IF
RETURN

END SUB

SUB PrintHelp (Help$)
  t$ = SPACE$(40)
  LSET t$ = Help$
  QPrintRC t$, 25, 2, -1
  t$ = ""
END SUB

SUB PrintTitle (Title$)
  t$ = SPACE$(40)
  LSET t$ = Title$
  QPrintRC t$, 1, 2, -1
  QPrintRC DATE$, 1, 70, -1
  t$ = ""
END SUB

SUB Search4LNum (LocNum$, RecNo&, CLSFlag%, ActiveOnly%)
  'done
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  DisplayUBScrn "SHOWSCRH"
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  IdxRecLen = 4 'we are using a integer
  IdxFileSize& = FileSize("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  SearchLen = LEN(LocNum$)
  
  Match = False
  FirstRec = 1
  LastRec = IdxNumOfRecs
  
  BotOffSet = 0
  TopOffSet = IdxNumOfRecs
  
  FOpenS "UBCUST.DAT", C1Handle 'open data file
  'Find matching record
  MidRec = (LastRec + FirstRec) \ 2
  
  DO
    IF LastSRec = MidRec THEN EXIT DO
    LastSRec = MidRec
    FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(MidRec).RecNum), UBCustRecLen
    UBSearchN$ = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
    ShowSearchWheel 12, 44
    
    IF (LocNum$ = UBSearchN$) AND (UBCustRec(1).DelFlag = 0) THEN
      IF MidRec - BotOffSet > 1 THEN
        MidRec = MidRec - 1
      ELSE
        FirstMatchRec = MidRec
      END IF
    ELSEIF LocNum$ < UBSearchN$ THEN            'lower
      TopOffSet = MidRec
      MidRec = TopOffSet - ((TopOffSet - BotOffSet) \ 2)
    ELSE        'higher
      BotOffSet = MidRec
      MidRec = BotOffSet + ((TopOffSet - BotOffSet) \ 2) + 1
      IF MidRec = IdxNumOfRecs + 1 THEN
        EXIT DO
      END IF
    END IF
    IF TopOffSet = BotOffSet THEN EXIT DO
    'Look into this
    'IF TopOffSet = BotOffSet THEN EXIT DO
  LOOP UNTIL FirstMatchRec
  
  FClose C1Handle
  
  IF FirstMatchRec = 0 THEN
    RecNo& = 0
  ELSE
    RecNo& = IdxBuff(FirstMatchRec).RecNum
  END IF
  
  IF ActiveOnly AND UBCustRec(1).STATUS <> "A" THEN
    RecNo& = 0
  ELSEIF (ActiveOnly = 1) AND (UBCustRec(1).STATUS <> "I") THEN
    RecNo& = 0
  END IF
ExitLSearch:
  
  'cls
  'Shell
  RestScrn ScrnArray()
  ERASE ScrnArray, UBCustRec, IdxBuff
END SUB

SUB Search4Meter (MeterNum$, RecNo&, CLSFlag%, ActiveOnly%)
  
  STATIC Choice, OMeterNum$
  
  IF OMeterNum$ <> MeterNum$ THEN
    Choice = 1
    OMeterNum$ = MeterNum$
  END IF
  
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  WPos = 1
  
  DisplayUBScrn "SHOWSCRH"
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  CustFileSize& = FileSize("UBCUST.DAT")
  NumCustRecs = CustFileSize& \ UBCustRecLen

  'REDIM MChoice(1 TO 1) AS FLen2

  REDIM RecBuff(1 TO 1) AS LONG
  
  FOpenS "UBCUST.DAT", C1Handle 'open data file
  'Find matching record
  
  MatchCnt = 0
  FOR Cnt = 1 TO NumCustRecs
    ShowSearchWheel 12, 44
    FGetRTA C1Handle, UBCustRec(1), CLNG(Cnt), UBCustRecLen
    IF NOT UBCustRec(1).DelFlag THEN
      'IF NOT ActiveOnly OR (ActiveOnly AND (UBCustRec(1).Status = "A")) THEN
      IF (ActiveOnly = 0) OR ((ActiveOnly = True) AND ((UBCustRec(1).STATUS = "A"))) THEN
        GOSUB CheckEM2
      ELSEIF (ActiveOnly = 1) AND (UBCustRec(1).STATUS = "I") THEN
        GOSUB CheckEM2
      END IF
    END IF
  NEXT
  
  IF MatchCnt = 0 THEN
    RecNo& = 0
    FClose C1Handle
    GOTO ExitMeterSearch
  END IF
  IF NOT DebugFlag THEN
    FreeMem& = FRE(-1)
    IF FreeMem& >= 65536 THEN
      FreeMem& = 65536
    END IF
    MemNeeded& = MatchCnt * 64&
    IF MemNeeded& > FreeMem& THEN
      FClose C1Handle
      QPrintRC "Matched:>" + QPStrI(MatchCnt), 24, 1, 15
      QPrintRC " Needed: " + QPStrL(MemNeeded&) + "  Free: " + QPStrL(FreeMem&), 25, 1, 15
      RecNo& = -1
      WaitForAction
      GOTO ExitMeterSearch
    END IF
  END IF
  REDIM MChoice(1 TO MatchCnt) AS FLen2
  
  FOR Cnt = 1 TO MatchCnt
    ShowSearchWheel 12, 44
    FGetRTA C1Handle, UBCustRec(1), RecBuff(Cnt), UBCustRecLen
    TCustName$ = LEFT$(QPTrim$(UBCustRec(1).CUSTNAME), 30)
    Cnt = Cnt - 1
    FOR MeterCnt = 1 TO 7
      IF INSTR(UBCustRec(1).LocMeters(MeterCnt).MTRNUM, MeterNum$) > 0 THEN
        Cnt = Cnt + 1
        LSET MChoice(Cnt).V = TCustName$
        BOOK$ = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
        MID$(MChoice(Cnt).V, 32, 12) = UBCustRec(1).LocMeters(MeterCnt).MTRNUM
        MID$(MChoice(Cnt).V, 50, 9) = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
        MID$(MChoice(Cnt).V, 61) = MKL$(RecBuff(Cnt))
      END IF
    NEXT
  NEXT
  
  FClose C1Handle
  
  IF MatchCnt = 0 THEN
    RecNo& = 0
  ELSEIF MatchCnt > 1 THEN
    QPrintRC "Sorting. . .  ", 12, 32, -1
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    IF Choice = 0 THEN
      Choice = 1                'Pre-load choice to highlight
    ELSEIF Choice > MatchCnt THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Customer/Owner               Meter No.         Location No."
    '--Find max menu width
    
    '--Center Menu within Screen
    
    Row = 4
    Col = ((80 - 60) \ 2) - 1
    
    IF CLSFlag THEN
      Row = 4
      BoxBot = 17               'limit the box length
      BlockClear
    ELSE
      Row = 6
      BoxBot = 14               'limit the box length to go no lower than line 20
      RestScrn ScrnArray()
    END IF
    
    LOCATE Row, Col, 0
    
    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Matched:" + STR$(MatchCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
    LOOP UNTIL RecNo& > 0
  ELSE
    RecNo& = CVL(MID$(MChoice(1).V, 61, 4))
  END IF
  
ExitMeterSearch:
  
  RestScrn ScrnArray()
  ERASE ScrnArray, UBCustRec, RecBuff, MChoice
  EXIT SUB

CheckEM2:
  FOR MeterCnt = 1 TO 7
    IF INSTR(UBCustRec(1).LocMeters(MeterCnt).MTRNUM, MeterNum$) > 0 THEN
      MatchCnt = MatchCnt + 1
      REDIM PRESERVE RecBuff(1 TO MatchCnt) AS LONG
      RecBuff(MatchCnt) = Cnt
    END IF
  NEXT
RETURN

END SUB

SUB Search4SAddr (SAddr$, RecNo&, CLSFlag%, ActiveOnly%)
  
  STATIC Choice
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  WPos = 1
  
  DisplayUBScrn "SHOWSCRH"
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  CustFileSize& = FileSize("UBCUST.DAT")
  NumCustRecs = CustFileSize& \ UBCustRecLen
  
  'REDIM MChoice(1 TO 1) AS FLen2

  REDIM RecBuff(1 TO 1) AS LONG
  
  FOpenS "UBCUST.DAT", C1Handle 'open data file
  
  'Find matching record
  
  MatchCnt = 0
  FOR Cnt = 1 TO NumCustRecs
    ShowSearchWheel 12, 44
    FGetRTA C1Handle, UBCustRec(1), CLNG(Cnt), UBCustRecLen
    IF NOT UBCustRec(1).DelFlag THEN
      IF (ActiveOnly = 0) OR ((ActiveOnly = True) AND ((UBCustRec(1).STATUS = "A"))) THEN
        GOSUB CheckLoadEM2
      ELSEIF (ActiveOnly = 1) AND (UBCustRec(1).STATUS = "I") THEN
        GOSUB CheckLoadEM2
      END IF
    END IF
  NEXT
  
  FClose C1Handle
  
  IF NOT DebugFlag THEN
    FreeMem& = FRE(-1)
    IF FreeMem& >= 65536 THEN
      FreeMem& = 65536
    END IF
    MemNeeded& = MatchCnt * 64&
    IF MemNeeded& > FreeMem& THEN
      'BEEP
      QPrintRC "Matched: " + QPStrI(MatchCnt), 24, 1, 15
      QPrintRC " Needed: " + QPStrL(MemNeeded&) + "  Over: " + QPStrL(MemNeeded& - FreeMem&), 25, 1, 15
      'MatchCnt = 0
      RecNo& = -1
      WaitForAction
      GOTO ExitSAddrSearch
    END IF
  END IF
  
  IF MatchCnt = 0 THEN
    GOTO ExitSAddrSearch
    RecNo& = -1
  ELSEIF MatchCnt > 1 THEN
    REDIM MChoice(1 TO MatchCnt) AS FLen2
    FOpenS "UBCUST.DAT", C1Handle               'open data file
    FOR Cnt = 1 TO MatchCnt
      FGetRTA C1Handle, UBCustRec(1), CLNG(RecBuff(Cnt)), UBCustRecLen
      Addr$ = LEFT$(QPTrim$(UBCustRec(1).SERVADDR), 25)
      BOOK$ = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
      LSET MChoice(Cnt).V = LEFT$(QPTrim$(UBCustRec(1).CUSTNAME), 20)
      MID$(MChoice(Cnt).V, 22, 25) = Addr$
      MID$(MChoice(Cnt).V, 50, 9) = BOOK$
      MID$(MChoice(Cnt).V, 61) = MKL$(RecBuff(Cnt))
      ShowSearchWheel 12, 44
    NEXT
    FClose C1Handle
    
    IF DCnt = 0 THEN
      RecNo& = 0
      GOTO ExitSAddrSearch
    END IF
    
    'FClose L1Handle
    
    QPrintRC "Sorting. . .   ", 12, 32, -1
    SortT MChoice(1), MatchCnt, 0, 64, 21, 25
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    IF Choice = 0 THEN
      Choice = 1                'Pre-load choice to highlight
    ELSEIF Choice > MatchCnt THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Customer                      Address"
    '--Find max menu width
    
    '--Center Menu within Screen
    
    Row = 4
    Col = ((80 - 60) \ 2) - 1
    
    IF CLSFlag THEN
      Row = 4
      BoxBot = 17               'limit the box length
      BlockClear
    ELSE
      Row = 6
      BoxBot = 14               'limit the box length to go no lower than line 20
      RestScrn ScrnArray()
    END IF
    
    LOCATE Row, Col, 0
    
    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Matched:" + STR$(MatchCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
    LOOP UNTIL RecNo& > 0
  ELSE
    RecNo& = RecBuff(MatchCnt)
  END IF
  
ExitSAddrSearch:
  
  'cls
  'Shell
  RestScrn ScrnArray()
  ERASE ScrnArray, UBCustRec, MChoice
  EXIT SUB

CheckLoadEM2:
  Cnted = Cnted + 1
  IF INSTR(UBCustRec(1).SERVADDR, SAddr$) > 0 THEN
    DCnt = DCnt + 1
    MatchCnt = MatchCnt + 1
    REDIM PRESERVE RecBuff(1 TO MatchCnt) AS LONG
    RecBuff(MatchCnt) = Cnt
  END IF
RETURN

END SUB

SUB ShowCustConsHist (CustRec&)
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  '  DisplayUBScrn "UBCUHIST"
  
  REDIM Metered(1 TO 15)
  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
  NumOfRevs = MaxRevsCnt
  FOR RevCnt = 1 TO 15
    RLen = LEN(QPTrim$(LEFT$(UBSetUpRec(1).Revenues(RevCnt).REVNAME, 14)))
    IF RLen >= 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    END IF
    IF UBSetUpRec(1).Revenues(RevCnt).UseMtr = "Y" THEN
      Metered(RevCnt) = True
    END IF
  NEXT
  
  
  REDIM MChoice(1 TO 1) AS FLen2
  '
  REDIM UBTranRec(1) AS UBTransRecType
  REDIM UBCustRec(1) AS NewUBCustRecType
  
  UBCustRecLen = LEN(UBCustRec(1))
  UBTranRecLen = LEN(UBTranRec(1))
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  GET UBFile, CustRec&, UBCustRec(1)
  CLOSE UBFile
  
  CurBal# = UBCustRec(1).CurrBalance
  PreBal# = UBCustRec(1).PrevBalance
  
  'Top:
  
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  
  PrevTranRec& = UBCustRec(1).LastTrans
  
  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
      GET UBTran, PrevTranRec&, UBTranRec(1)
      IF UBTranRec(1).TransType = TranUtilityBill OR UBTranRec(1).TransType = TranUtilityBill + 100 THEN
        FOR MtrCnt = 1 TO 7
          IF UBTranRec(1).MtrTypes(MtrCnt) > 0 THEN
            DCnt = DCnt + 1
            REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
            LSET MChoice(DCnt).V = Num2Date(UBTranRec(1).TransDate)
            SELECT CASE UBTranRec(1).MtrTypes(MtrCnt)
            CASE MtrWaterOnly
              MeterType$ = "Water"
            CASE MtrSewerOnly
              MeterType$ = "Sewer"
            CASE MtrCombined
              MeterType$ = "Combined"
            CASE MtrElectric
              MeterType$ = "Electric"
            CASE MtrDemand
              MeterType$ = "D Electric"
            CASE MtrGas
              MeterType$ = "Gas Meter"
            CASE MtrTouchRead
              MeterType$ = "Touch Read"
            CASE MtrLightsService
              MeterType$ = "L Service"
            END SELECT
            MID$(MChoice(DCnt).V, 13) = MeterType$
            MID$(MChoice(DCnt).V, 26) = FUsing$(STR$(UBTranRec(1).CurRead(MtrCnt)), "##########")
            MID$(MChoice(DCnt).V, 38) = FUsing$(STR$(UBTranRec(1).PrevRead(MtrCnt)), "##########")
            
            MeterConsp& = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
            IF MeterConsp& < 0 THEN
              MaxMeterAmt& = 10& ^ (LEN(STR$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
              MeterConsp& = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
            END IF
            MID$(MChoice(DCnt).V, 52) = FUsing$(STR$(MeterConsp&), "##########")
          END IF
        NEXT
      END IF
      PrevTranRec& = UBTranRec(1).PrevTrans
    LOOP
    
    CLOSE UBTran
    RestScrn TempScrn()
    MPaintBox 3, 5, 22, 75, 8
    
    'HistTop:
    
    MaxLen = 62 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    Balance$ = Title$
    LSET Title$ = " Trans Date   Meter Type      Current    Previous    Consumption"
    
    '--Find max menu width
    '--Center Menu within Screen
    
    Row = 4
    Col = 8
    
    Row = 6
    BoxBot = 17 'limit the box length to go no lower than line 20
    
    TitleBox BoxBot + 3, Col, MaxLen + 3, "Press <ESC> to continue.", Cnf
    
    QPrintRC Title$, Row - 1, Col, 112
    MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
    
    DO
      LOCATE Row, Col, 0
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RestScrn TempScrn()
        EXIT DO 'choice = 0
      END IF
    LOOP        'UNTIL EditLocRec& > 0
  ELSE
    CLOSE UBTran
    Ok = MsgBox%("UB.QSL", "NOCTRANS")
    RestScrn TempScrn()
  END IF
  
  RestScrn TempScrn()
  ERASE Metered, UBSetUpRec, MChoice
  ERASE TempScrn, UBTranRec, UBCustRec
  
  EXIT SUB
  
  
END SUB

SUB ShowCustHistory (CustRec&)
  
  REDIM TempScrn(0)
  SaveScrn TempScrn()
  
  DisplayUBScrn "UBCUHIST"
  
  REDIM RevText$(1 TO MaxRevsCnt)
  REDIM Metered(1 TO 15)
  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
  NumOfRevs = MaxRevsCnt
  FOR RevCnt = 1 TO 15
    RevText$(RevCnt) = LEFT$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME), 14)
    IF LEN(RevText$(RevCnt)) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    END IF
    IF UBSetUpRec(1).Revenues(RevCnt).UseMtr = "Y" THEN
      Metered(RevCnt) = True
    END IF
  NEXT
  
  REDIM MChoice(1 TO 1) AS FLen2
  
  REDIM UBTranRec(1) AS UBTransRecType
  REDIM UBCustRec(1) AS NewUBCustRecType
  
  UBCustRecLen = LEN(UBCustRec(1))
  UBTranRecLen = LEN(UBTranRec(1))
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  GET UBFile, CustRec&, UBCustRec(1)
  CLOSE UBFile
  
  CurBal# = UBCustRec(1).CurrBalance
  PreBal# = UBCustRec(1).PrevBalance
  
Top:
  
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  
  PrevTranRec& = UBCustRec(1).LastTrans
  
  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
      DCnt = DCnt + 1
      REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
      GET UBTran, PrevTranRec&, UBTranRec(1)
      LSET MChoice(DCnt).V = Num2Date(UBTranRec(1).TransDate)
      MID$(MChoice(DCnt).V, 15) = UBTranRec(1).TransDesc
      MID$(MChoice(DCnt).V, 37) = FUsing(STR$(UBTranRec(1).TransAmt), "#####.##")
      MID$(MChoice(DCnt).V, 50) = FUsing(STR$(UBTranRec(1).RunBalance), "#####.##")
      MID$(MChoice(DCnt).V, 61) = MKL$(PrevTranRec&)
      PrevTranRec& = UBTranRec(1).PrevTrans
    LOOP
    
    CLOSE UBTran
    
    'first sort the transactions into year
    '    SortT MChoice(1), DCnt, -1, 64, 6, 4
    '
    '    'this sorts each year by month & day
    '    First = 1
    '    DO
    '      FOR Cnt = First TO DCnt
    '        IF Cnt = DCnt THEN EXIT FOR
    '        IF MID$(MChoice(Cnt).V, 7, 4) = MID$(MChoice(Cnt + 1).V, 7, 4) THEN
    '          Last = Cnt
    '          DO
    '            Last = Last + 1
    '            IF Last > DCnt THEN EXIT DO
    '          LOOP WHILE MID$(MChoice(Cnt).V, 7, 4) = MID$(MChoice(Last).V, 7, 4)
    '          SortT MChoice(Cnt), Last - Cnt, -1, 64, 0, 64
    '          First = Last
    '          EXIT FOR
    '        END IF
    '      NEXT
    '    LOOP UNTIL Cnt >= DCnt
    
    RestScrn TempScrn()
    MPaintBox 3, 5, 22, 75, 8
    REDIM TempScrn2(0)
    SaveScrn TempScrn2()
    
HistTop:
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    Balance$ = Title$
    LSET Title$ = " Trans Date       Description        Trans Amt      Balance  "
    LSET Balance$ = " Balance:" + FUsing(STR$(CurBal# + PreBal#), ",#####.##") + "   Cur:" + FUsing(STR$(CurBal#), ",#####.##") + "  Prev:" + FUsing(STR$(PreBal#), ",#####.##")
    
    '--Find max menu width
    '--Center Menu within Screen
    
    Row = 4
    Col = ((80 - 60) \ 2) - 1
    
    Row = 6
    BoxBot = 17 'limit the box length to go no lower than line 20
    
    TitleBox BoxBot + 3, Col, MaxLen + 3, "       Press <ESC> to continue.", Cnf
    
    QPrintRC "Transactions:" + STR$(DCnt), BoxBot + 4, Col + 2, 15
    QPrintRC Balance$, Row - 2, Col, 112
    QPrintRC Title$, Row - 1, Col, 112
    MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
    'FirstTime = True
    DO
      LOCATE Row, Col, 0
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RestScrn TempScrn()
        EXIT DO 'choice = 0
      ELSEIF Ky$ = CHR$(13) THEN
        RestScrn TempScrn2()
        GOTO ShowTransDetail
      END IF
    LOOP        'UNTIL EditLocRec& > 0
  ELSE
    CLOSE UBTran
    Ok = MsgBox%("UB.QSL", "NOCTRANS")
    RestScrn TempScrn()
  END IF
  
  RestScrn TempScrn()
  ERASE RevText$, Metered, UBSetUpRec, MChoice
  ERASE TempScrn, UBTranRec, UBCustRec
  
  EXIT SUB
  
ShowTransDetail:
  CursorOff
  TransRecNum& = CVL(RIGHT$(MChoice(Choice).V, 4))
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  GET UBTran, TransRecNum&, UBTranRec(1)
  CLOSE UBTran
  
  DisplayUBScrn "TRDETAIL"
  
  QPrintRC Num2Date(UBTranRec(1).TransDate), 3, 23, 15
  
'CONST TranUtilityBill = 1          '   1=Utility bill
'CONST TranLateCharge = 2           '   2=late charge
'CONST TranReconnectFee = 3         '   3=reconnect fee
'CONST TranBillPayment = 4          '   4=Bill Payment
'CONST TranAppliedDeposit = 5       '   5=Applied Deposit
'CONST TranPenaltyCharge = 6        '   6=Penalty Charge
'CONST TranDepositPayment = 7       '   7=Deposit Payment
'CONST TranDraftPayment = 8         '   8=Draft Payment
'CONST TranRefundDeposit = 9       '    9=Refund Deposit
'CONST TranBeginBalance = 10        '  10=Beginning Balance
'CONST TranUpwardAdjustment = 11    '  11=Bill Adjustments
'CONST TranDownwardAdjustment = 12  '  12=Bill Adjustments
  
  SELECT CASE UBTranRec(1).TransType
  CASE TranUtilityBill, TranUtilityBill + 100
    TType$ = "Utility Bill"
  CASE TranLateCharge, TranReconnectFee, TranLateCharge + 100, TranReconnectFee + 100
    TType$ = "Penalty, Reconnect Fee"
  CASE TranBillPayment, TranBillPayment + 100
    TType$ = "Utility Payment"
  CASE TranPenaltyPayment
    TType$ = "Penalty Payment"
  CASE TranPenaltyCharge
    TType$ = "Penalty/Late Fee"

  CASE TranDepositPayment, TranDepositPayment + 100
    TType$ = "Deposit"

  CASE TranDraftPayment
    TType$ = "Draft Payment"
  CASE TranBeginBalance, TranBeginBalance + 100
    TType$ = "Beginning Balance"

  CASE 9
    TType$ = "Deposit Refund"

  CASE TranUpwardAdjustment
    TType$ = "Upward Billing Adjustment"
  CASE TranDownwardAdjustment
    TType$ = "Downward Billing Adjustment"
  CASE ELSE
    TType$ = STR$(UBTranRec(1).TransType) + " ???"
  END SELECT
  
  QPrintRC FUsing$(STR$(UBTranRec(1).TransAmt), "#####.##"), 4, 25, 15
  
  QPrintRC TType$, 4, 50, 15
  QPrintRC UBTranRec(1).TransDesc, 3, 50, 15
  
  FOR RevCnt = 1 TO NumOfRevs
    QPrintRC RevText$(RevCnt), RevCnt + 6, 8, 15
    QPrintRC FUsing$(STR$(UBTranRec(1).RevAmt(RevCnt)), "#####.##"), RevCnt + 6, 25, 15
    QPrintRC FUsing$(STR$(UBTranRec(1).TaxAmt(RevCnt)), "###.##"), RevCnt + 6, 36, 15
    '(Number$, Image$)
  NEXT
  
  FOR Cnt = 1 TO 7
    IF Metered(Cnt) THEN
      QPrintRC FUsing$(STR$(UBTranRec(1).CurRead(Cnt)), "#########"), Cnt + 6, 44, 15
      QPrintRC FUsing$(STR$(UBTranRec(1).PrevRead(Cnt)), "#########"), Cnt + 6, 55, 15
      IF QPTrim$(UBTranRec(1).ESTREAD(Cnt)) = "" THEN
        QPrintRC "N", Cnt + 6, 70, 15
      ELSE
        QPrintRC "Y", Cnt + 6, 70, 15
      END IF
    END IF
  NEXT
  
  WaitForAction
  RestScrn TempScrn2()
  GOTO HistTop
  
END SUB

SUB ShowWrkOrdHistory (CustRec&)
  
  REDIM TScrn1(0)
  SaveScrn TScrn1()
  
  DisplayUBScrn "UBCUHIST"
  
  REDIM MChoice(1 TO 1) AS FLen2
  
  REDIM WorkOrderRec(1) AS WorkOrderRecType
  WorkOrderRecLen = LEN(WorkOrderRec(1))
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  GET UBFile, CustRec&, UBCustRec(1)
  CLOSE UBFile

'Top:
  
  UBWrkOrd = FREEFILE
  OPEN "UBWRKORD.DAT" FOR RANDOM SHARED AS UBWrkOrd LEN = WorkOrderRecLen
  
  PrevTranRec& = UBCustRec(1).WOLastTrans
  
  IF PrevTranRec& > 0 THEN
    DO WHILE PrevTranRec& > 0
      DCnt = DCnt + 1
      REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
      GET UBWrkOrd, PrevTranRec&, WorkOrderRec(1)
      LSET MChoice(DCnt).V = FUsing(STR$(PrevTranRec&), "#####")
      MID$(MChoice(DCnt).V, 15) = Num2Date$(WorkOrderRec(1).ENTRYDATE)
      MID$(MChoice(DCnt).V, 32) = Num2Date$(WorkOrderRec(1).CompleteByDate)
      IF WorkOrderRec(1).CompletedDate <= 0 THEN
        MID$(MChoice(DCnt).V, 50) = "Open"
      ELSE
        MID$(MChoice(DCnt).V, 50) = Num2Date$(WorkOrderRec(1).CompletedDate)
      END IF
      MID$(MChoice(DCnt).V, 61) = MKL$(PrevTranRec&)
      PrevTranRec& = WorkOrderRec(1).PrevTransRec
    LOOP
    
    CLOSE UBWrkOrd
    
    RestScrn TScrn1()
    MPaintBox 3, 5, 22, 75, 8
    REDIM TScrn2(2000)
    SaveScrn TScrn2()
    
WOTop:
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    Balance$ = Title$
    LSET Balance$ = " Work Order       Entry           Complete         Completed    "
    LSET Title$ = "   Number          Date           By Date            Date       "
    
    '--Find max menu width
    '--Center Menu within Screen
    
    Row = 4
    Col = ((80 - 60) \ 2) - 1
    
    Row = 6
    BoxBot = 17 'limit the box length to go no lower than line 20
    
    TitleBox BoxBot + 3, Col, MaxLen + 3, "       Press <ESC> to continue.", Cnf
    
    QPrintRC "Transactions:" + STR$(DCnt), BoxBot + 4, Col + 2, 15
    QPrintRC Balance$, Row - 2, Col, 112
    QPrintRC Title$, Row - 1, Col, 112
    MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
    'FirstTime = True
    DO
      LOCATE Row, Col, 0
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RestScrn TScrn1()
        EXIT DO 'choice = 0
      ELSEIF Ky$ = CHR$(13) THEN
        RestScrn TScrn2()
        GOTO WOShowDetail
      END IF
    LOOP        'UNTIL EditLocRec& > 0
  ELSE
    CLOSE UBWrkOrd
    Ok = MsgBox%("UB.QSL", "NOCTRANS")
  END IF
  
  RestScrn TScrn1()
  ERASE MChoice
  ERASE UBCustRec, WorkOrderRec
  
  EXIT SUB
  
WOShowDetail:
  CursorOff
  TransRecNum& = CVL(RIGHT$(MChoice(Choice).V, 4))
  UBWrkOrd = FREEFILE
  OPEN "UBWRKORD.DAT" FOR RANDOM SHARED AS UBWrkOrd LEN = WorkOrderRecLen
  GET UBWrkOrd, TransRecNum&, WorkOrderRec(1)
  CLOSE UBWrkOrd
  
  LibFile2Scrn "UBWRKORD", "WODETAIL", MonoCode, Attribute%, ErrCode
  
  QPrintRC STR$(TransRecNum&), 4, 23, 15
  IF WorkOrderRec(1).CompletedDate <= 0 THEN
    TDate$ = "Open"
  ELSE
    TDate$ = Num2Date$(WorkOrderRec(1).CompletedDate)
  END IF
  QPrintRC Num2Date$(WorkOrderRec(1).ENTRYDATE), 5, 23, 15
  QPrintRC TDate$, 4, 63, 15
  QPrintRC Num2Date$(WorkOrderRec(1).CompleteByDate), 5, 63, 15
  
  FOR Cnt = 1 TO 6
    QPrintRC WorkOrderRec(1).OrdersText.Text(Cnt), Cnt + 7, 8, 15
    QPrintRC WorkOrderRec(1).RepliesText.Text(Cnt), Cnt + 14, 8, 15
  NEXT
  
  WaitForAction
  RestScrn TScrn2()
  GOTO WOTop
  
END SUB

