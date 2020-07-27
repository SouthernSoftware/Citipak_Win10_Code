DEFINT A-Z
DECLARE SUB CDisplayUBScrn (ScrnName$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB CShowProcessingScrn (RptTitle$)

DECLARE SUB MakeMowZipCodeIndex (IndexText$)
DECLARE SUB Search4911Addr (S911$, RecNo&, CLSFlag%, ActiveOnly%)
DECLARE FUNCTION CustHasMsg% (RecNo&)
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
DECLARE SUB ShowCursor ()
DECLARE SUB SaveScrn (array())
DECLARE SUB RestScrn (array())
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FGetA (Handle%, SEG Dest AS ANY, NumBytes%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)

DECLARE FUNCTION Chk4DupeBookSeqNum (Book$, SeqNum$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION QPValI% (Number$)

DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION FmtBook$ (Book$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION FindRateTbl% (RATECODE$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB UBLog (Text$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION ConvDateStr$ (d$)
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
DECLARE SUB KillFile (File2Kill$)
DECLARE FUNCTION GetCustMeterType% (UBCustRec() AS ANY, ThisMeter%)
DECLARE FUNCTION ChkBillFile% ()
DECLARE FUNCTION MakeMonth$ (TDate$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)



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

  DIM SHARED PctC(1) AS STRING * 3

  UBLog " IN: CReindex Utility Files"

  BlockClear
  'PRINT FRE(0)
  WaitForAction

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))              'Length of Cust Record Structure

  REDIM UBTransRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTransRec(1))             'Length of Tran Record Structure

  CShowProcessingScrn "Reading Customer Names"
  UBLog "BEGIN: Customer Name Reindex"
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  REDIM IdxBuff(1 TO NumOfRecs&) AS nUBCustReIndexRecType

  FOR Cnt = 1 TO NumOfRecs&
    GET UBFile, Cnt, UBCustRec(1)
    IdxBuff(Cnt).SearchName = UBCustRec(1).SEARCH
    IF UBCustRec(1).DelFlag THEN
      IdxBuff(Cnt).DelFlag = "Y"
    ELSE
      IdxBuff(Cnt).DelFlag = ""
    END IF
    IdxBuff(Cnt).Status = UBCustRec(1).Status
    IdxBuff(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumOfRecs&
  NEXT

  CLOSE UBFile

  QPrintRC " Sorting Customer Names", 11, 29, -1

  SortT IdxBuff(1), CINT(NumOfRecs&), 0, 16, 0, 10

  GOSUB ClearBlock
  QPrintRC "Writing Customer Index", 9, 30, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1

  KillFile "UBCUSTNM.IDX"
  UBFile = FREEFILE
  OPEN "UBCUSTNM.IDX" FOR RANDOM SHARED AS UBFile LEN = 4
  FOR Cnt = 1 TO NumOfRecs&
    PUT UBFile, Cnt, IdxBuff(Cnt).RecNum
    ShowPctComp Cnt, NumOfRecs&
  NEXT
  CLOSE UBFile

  GOSUB ClearBlock
  QPrintRC "Writing Customer Search Data", 9, 27, 126
  QPrintRC "Processing:    % Complete", 13, 28, -1

  KillFile "UBCUSTSN.DAT"
  UBFile = FREEFILE
  OPEN "UBCUSTSN.DAT" FOR RANDOM SHARED AS UBFile LEN = LEN(IdxBuff(1))
  FOR Cnt = 1 TO NumOfRecs&
    PUT UBFile, Cnt, IdxBuff(Cnt)
    ShowPctComp Cnt, NumOfRecs&
  NEXT
  CLOSE UBFile

  ERASE IdxBuff
  UBLog "FINISH: Customer Name Reindex"
  GOSUB ClearBlock

  QPrintRC "Reading Location Information", 9, 27, 126
  QPrintRC "Processing:    % Complete", 13, 28, -1
  UBLog "BEGIN: Book\Sequence Reindex"

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen

  REDIM LIdxBuff(1 TO NumOfRecs&) AS UBLocaReIndexRecType

  FOR Cnt = 1 TO NumOfRecs&
    GET UBFile, Cnt, UBCustRec(1)
    LIdxBuff(Cnt).Book = UBCustRec(1).Book
    LIdxBuff(Cnt).SeqNumb = UBCustRec(1).SeqNumb
    LIdxBuff(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumOfRecs&
  NEXT

  CLOSE UBFile

  QPrintRC " Sorting Locations Names", 11, 29, -1

  SortT LIdxBuff(1), CINT(NumOfRecs&), 0, 16, 0, 8
  'Array(1), NumElem, Dir, StructSize, MemOff, MemSize

  GOSUB ClearBlock
  QPrintRC "Writing Location Index", 9, 30, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1
  'here
  KillFile "UBCUSTBK.IDX"

  UBFile = FREEFILE
  OPEN "UBCUSTBK.IDX" FOR RANDOM SHARED AS UBFile LEN = 4

  FOR Cnt = 1 TO NumOfRecs&
    PUT UBFile, Cnt, LIdxBuff(Cnt).RecNum
    ShowPctComp Cnt, NumOfRecs&
  NEXT
  CLOSE UBFile

  UBLog "FINISH: Book\Sequence Reindex"
  REDIM BookSeq(1) AS BookSeqRecType

  KillFile "UBOOKSEQ.DAT"
  UBLog "BEGIN: Rebuild Book\Sequence List"
  BookHand = FREEFILE
  OPEN "UBOOKSEQ.DAT" FOR RANDOM SHARED AS BookHand LEN = 4
  FOR Cnt = 1 TO NumOfRecs&
    BkSq$ = QPTrim$(LIdxBuff(Cnt).Book) + QPTrim$(LIdxBuff(Cnt).SeqNumb)
    BookSeq(1).BookSeq = QPValL(BkSq$)
    PUT BookHand, Cnt, BookSeq(1)
  NEXT
  CLOSE BookHand
  UBLog "FINISH: Rebuild Book\Sequence List"

  ERASE LIdxBuff, BookSeq, IdxBuff
  ERASE UBCustRec, UBTransRec

  BlockClear

  CDisplayUBScrn "UPDATEOK"

  WaitForAction

ExitReindex:
  UBLog "OUT: Reindex Utility Files" + CRLF$
  RUN "ubmenu"

ClearBlock:
  HideCursor
  Blank$ = SPACE$(40)
  FOR Cnt = 8 TO 15
    QPrintRC Blank$, Cnt, 21, -1
  NEXT
  ShowCursor
RETURN

SUB CDisplayUBScrn (ScrnName$)
  LibFile2Scrn "UB", ScrnName$, MonoCode, Attribute%, ErrCode
END SUB

SUB CReindexSystem (PromptFlag)


END SUB

SUB CShowProcessingScrn (RptTitle$)
  TitleRow = 9
  TitleCol = 40 - (LEN(RptTitle$) \ 2) + 1
  CursorOff
  BlockClear
  DisplayUBScrn "PRORPT"
  HideCursor
  QPrintRC RptTitle$, TitleRow, TitleCol, 126
  QPrintRC "Processing:    % Completed.", 13, 28, Cnf.HiLite
  ShowCursor
END SUB

