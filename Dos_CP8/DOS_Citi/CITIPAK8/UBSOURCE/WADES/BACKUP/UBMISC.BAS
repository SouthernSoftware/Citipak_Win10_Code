DEFINT A-Z

DECLARE SUB RelinkWorkOrders ()
DECLARE SUB ubGetAcctStruct (GLFundLen%, GLAcctLen%, GLDetLen%)
DECLARE SUB UBLog (Text$)
DECLARE SUB OpenAcctIdx (AcctIdxFileNum%, NumIdxRecs%)
DECLARE SUB OpenAcctFile (AcctFileNum%, NumAccts%)
DECLARE SUB ReLinkHistory ()
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB PrintRateListing ()
DECLARE SUB ReIndexSystem ()
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB SmallPause ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB EnterNewBalance ()
DECLARE FUNCTION GetNumOfCust% ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB SetBeginningBalance ()
DECLARE SUB EnterEditBalance ()
DECLARE SUB PrintBalanceTrans ()
DECLARE SUB PostBalanceTrans ()
DECLARE SUB UBRateMaint ()
DECLARE SUB AddEditRateTbl (EdFlag%)
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB LoadRateTblRec (Form$(), TblNum%)
DECLARE SUB SaveRateTblRec (Form$(), RecNo%)
DECLARE SUB SelectRate2Edit (RecNo%)
DECLARE SUB EditUBSetUpRec ()
DECLARE SUB LoadSetUpRec (Form$(), OKFlag)
DECLARE SUB SaveSetUpRec (Form$())
DECLARE SUB LoadSysDraftRec (Form$())
DECLARE SUB EditBankDraftRec ()
DECLARE SUB SaveSysDraftRec (Form$())
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB WaitForAction ()
DECLARE SUB EditSysBankDraft ()
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB KillFile (File2Kill$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB QPrintRC (Text$, Row%, Col%, Kolor)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE FUNCTION GetAcctTitle$ (AcctRecNum)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB SortStr2 (BYVAL Address%, Size%, Direction%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Num2Date$ (DateNum)
  
'$INCLUDE: 'DefCnf.BI'
  
DECLARE FUNCTION ASCII (Character$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Number$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION QPStrL$ (Number&)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)
DECLARE FUNCTION GetNumRateRecs ()
DECLARE SUB ClearScrn ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE FUNCTION ubFmtAcct$ (AcctNum$, FundLen%, AcctLen%, DetLen%)
DECLARE FUNCTION ubStripAcct$ (AcctNum$)
DECLARE FUNCTION ubFindAcct% (AcctNum$)
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'ubtrans.bi'
  '$INCLUDE: 'PageInfo.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.BI'
  
  '$INCLUDE: 'GL.BI'
  ''$INCLUDE: 'AP.BI'
  
  CONST False = 0, True = NOT False

  CrLf$ = CHR$(13) + CHR$(10)

  'GOTO StartHere
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 7)
  
  MChoice$(1) = "Utility System Setup"
  MChoice$(2) = "Rate Code Maintenance"
  MChoice$(3) = "Bank Draft Setup"
  MChoice$(4) = "Relink Transactions History"
  MChoice$(5) = "Relink Work Order History"
  MChoice$(6) = "Reindex Utility Files "
  MChoice$(7) = "Exit to OS"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Miscellaneous Functions ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN
      Choice = 0
      ExitFlag = True
      EXIT DO   'choice = 0
    END IF
    
    SELECT CASE Choice
    CASE 1
      RUN "ubsetup"
    CASE 2
      UBRateMaint
    CASE 3
      EditBankDraftRec
    CASE 4
      ReLinkHistory
    CASE 5
      RelinkWorkOrders
    CASE 6
      ReIndexSystem
    CASE 7
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF WEnvTest AND INSTR(COMMAND$, "TEST") THEN
    HideCursor
    ClearScrn
    END
  ELSE
    RUN "ubmenu"
  END IF
  
StartHere:

SUB AddEditRateTbl (EdFlag)

  UBLog "ENTERED: Add\Edit Rate Code"
RateTop:
  REDIM ScrnArray(0)
  
  LibName$ = "UB"
  ScrnName$ = "UBNRATE"
  NumScrns = 1
  
  '--define the multi-choice fields
  SHARED Choice$()
  REDIM Choice$(0 TO 4, 0)
  
  '--Initialize the form name array
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
  
  IF EdFlag THEN
    SelectRate2Edit RecNo
    IF RecNo > 0 THEN
      LoadRateTblRec Form$(), RecNo
      UnPackBuffer 0, 0, Form$(), Fld()
      UBLog "EDITING RATE:" + STR$(RecNo) + "  " + Form$(1, 0)
    ELSE
      GOTO ExitEdit
    END IF
  ELSE
    RecNo = GetNumRateRecs + 1
    UBLog "ADDING NEW RATE:" + STR$(RecNo)
  END IF
  
  '--Set screen number to one and display screen
  Scr = 1
  BlockClear
  LibFile2Scrn LibName$, ScrnName$, MonoCode, Attribute%, ErrCode
  
  ShowCursor
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      SaveScrn ScrnArray()
      SaveRateTblRec Form$(), RecNo
      UBLog "UPDATED RATE:" + STR$(RecNo) + "  " + Form$(1, 0)
      ExitFlag = True
      RestScrn ScrnArray()
    CASE ESCKey
      UBLog "ABORTED EDIT:" + STR$(RecNo) + "  " + Form$(1, 0)
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 21   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 43 TO 54           '--Save Button
          PressButton -68, 21, 43, 54
        CASE 62 TO 72           '--cancel button
          PressButton 27, 21, 57, 70
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
  LOOP UNTIL ExitFlag
  'IF Frm(1).KeyCode = 27 THEN GOTO ExitEdit
  'IF NOT EdFlag GOTO ExitEdit
  'GOTO RateTop

ExitEdit:
  UBLog "EXITED: Add\Edit Rate Code" + CrLf$
  HideCursor
  EXIT SUB
  
END SUB

SUB EditBankDraftRec
  
  LibName$ = "UB"
  ScrnName$ = "UBDRAFT"
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
  
  LoadSysDraftRec Form$()
  CALL UnPackBuffer(0, 0, Form$(), Fld())
  
  '--Set screen number to one and display screen
  Scr = 1
  LibFile2Scrn LibName$, ScrnName$, MonoCode, Attribute%, ErrCode
  
  ShowCursor
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      SaveSysDraftRec Form$()
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 21   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 43 TO 54           '--Save Button
          PressButton -68, 21, 43, 54
        CASE 62 TO 72           '--cancel button
          PressButton 27, 21, 57, 70
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
    
  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag
  
  HideCursor
  
END SUB

SUB LoadRateTblRec (Form$(), TblNum)
  
  REDIM UBRateTblRec(1) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))
  
  Form$(0, 0) = SPACE$(UBRateTblRecLen)
  
  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  GET UBFile, TblNum, UBRateTblRec(1)
  CLOSE
  
  BCopy VARSEG(UBRateTblRec(1)), VARPTR(UBRateTblRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), UBRateTblRecLen, 0
  ERASE UBRateTblRec
  
END SUB

SUB LoadSysDraftRec (Form$())
  
  REDIM UBSysDraftRec(1) AS UBDraftRecType
  UBSysDraftRecLen = LEN(UBSysDraftRec(1))
  Form$(0, 0) = SPACE$(UBSysDraftRecLen)
  
  UBFile = FREEFILE
  OPEN "UBSDRAFT.DAT" FOR RANDOM SHARED AS UBFile LEN = UBSysDraftRecLen
  GET UBFile, 1, UBSysDraftRec(1)
  CLOSE
  
  BCopy VARSEG(UBSysDraftRec(1)), VARPTR(UBSysDraftRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), UBSysDraftRecLen, 0
  
  ERASE UBSysDraftRec
  
END SUB

SUB oShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&) STATIC
  IF NOT BeenDone THEN
    PctC2$ = SPACE$(3)
    BeenDone = True
  END IF
  RSET PctC2$ = QPStrL$(INT((RecNo& / NumOfRecs&) * 100))
  HideCursor
  QPrintRC PctC2$, 13, 40, Cnf.HiLite
  ShowCursor
END SUB

SUB PrintRateListing
  
  REDIM UBRateTblRec(1) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))
  
  NumRateRecs = FileSize("UBRATE.DAT") \ UBRateTblRecLen
  
  IF NumRateRecs = 0 THEN
    GOTO ExitRateListing
  END IF
  
  BlockClear
  ShowProcessingScrn "Rate Code Listing Report."
  
  REDIM StepText(1 TO 10) AS STRING * 40
  REDIM UBSetUp(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUp(1))
  UBFile = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM SHARED AS UBFile LEN = UBSetupLen
  IF LOF(UBFile) = 0 THEN
    OKFlag = False
  ELSE
    OKFlag = True
    GET UBFile, 1, UBSetUp(1)
  END IF
  CLOSE
  
  TownName$ = UBSetUp(1).UTILNAME
  Dash80$ = STRING$(78, "-")
  NumPrinted = 0
  
  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  RPTFile = FREEFILE
  OPEN "RATELIST.RPT" FOR OUTPUT AS RPTFile
  GOSUB PrintRateHeader
  FOR Cnt = 1 TO NumRateRecs
    GET UBFile, Cnt, UBRateTblRec(1)
    IF NumPrinted = 3 THEN
      PRINT #RPTFile, Dash80$
      PRINT #RPTFile, CHR$(12)
      GOSUB PrintRateHeader
    END IF
    PRINT #RPTFile, "       Rate Code:  "; UBRateTblRec(1).RATECODE
    PRINT #RPTFile, "     Description:  "; UBRateTblRec(1).RATEDESC
    PRINT #RPTFile, "  Minimum Charge:"; FUsing$(STR$(UBRateTblRec(1).MINAMT), "#######.##")
    PRINT #RPTFile, "   Minimum Units:"; FUsing$(STR$(UBRateTblRec(1).MINUNITS), "##########")
    PRINT #RPTFile, "       Discount%:"; FUsing$(STR$(UBRateTblRec(1).DiscPct), "#######")
    PRINT #RPTFile, "      [ Step ]        [ Beg Unit ]     [ Amount/Unit ]"
    FOR RCnt = 1 TO 10
      LSET StepText$(RCnt) = ""
      IF UBRateTblRec(1).TblBreaks(RCnt).UNITS > 0 THEN
        MID$(StepText$(RCnt), 8) = FUsing$(STR$(UBRateTblRec(1).TblBreaks(RCnt).UNITS), "########")
      END IF
      IF UBRateTblRec(1).TblBreaks(RCnt).UNITAMT > 0 THEN
        MID$(StepText$(RCnt), 25) = FUsing$(STR$(UBRateTblRec(1).TblBreaks(RCnt).UNITAMT), "####.######")
      END IF
    NEXT
    PRINT #RPTFile, "     First Break:"; StepText$(1)
    PRINT #RPTFile, "    Second Break:"; StepText$(2)
    PRINT #RPTFile, "     Third Break:"; StepText$(3)
    PRINT #RPTFile, "    Fourth Break:"; StepText$(4)
    PRINT #RPTFile, "     Fifth Break:"; StepText$(5)
    PRINT #RPTFile, "     Sixth Break:"; StepText$(6)
    PRINT #RPTFile, "   Seventh Break:"; StepText$(7)
    PRINT #RPTFile, "    Eighth Break:"; StepText$(8)
    PRINT #RPTFile, "     Ninth Break:"; StepText$(9)
    PRINT #RPTFile, "        All Over:"; StepText$(10)
    PRINT #RPTFile,
    NumPrinted = NumPrinted + 1
    ShowPctComp Cnt, NumRateRecs
  NEXT
  PRINT #RPTFile, Dash80$
  PRINT #RPTFile, CHR$(12)
  CLOSE
  
  ERASE UBSetUp, UBRateTblRec, StepText
  
  PrintRptFile "Rate Code Listing Report.", "RATELIST.RPT", 1, RetCode%, 1
  KillFile "RATELIST.RPT"
  GOTO ExitRateListing
  
PrintRateHeader:
  PageNo = PageNo + 1
  PRINT #RPTFile, "Utility Billing Rate Table Listing."
  PRINT #RPTFile, TownName$; TAB(70); "Page:"; PageNo
  PRINT #RPTFile, "Report Date: "; DATE$
  PRINT #RPTFile, Dash80$
  NumPrinted = 0
  RETURN
  
ExitRateListing:
END SUB

SUB ReIndexSystem
  
  UBLog "ENTERED: Reindex Utility Files"
  BlockClear
  Ok = MsgBox%("UB", "MUSTEXIT")
  
  SELECT CASE Ok
  CASE 2
    GOTO ExitReindex
  END SELECT
  
  'BlockClear
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))              'Length of Cust Record Structure
  
  REDIM UBTransRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTransRec(1))             'Length of Tran Record Structure
  
  ShowProcessingScrn "Reading Customer Names"
  UBLog "BEGIN: Customer Name Reindex"
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfRecs& = LOF(UBFile) \ UBCustRecLen
  
  REDIM IdxBuff(1 TO NumOfRecs&) AS UBCustReIndexRecType
  
  FOR Cnt = 1 TO NumOfRecs&
    GET UBFile, Cnt, UBCustRec(1)
    IdxBuff(Cnt).SearchName = UBCustRec(1).SEARCH
    IdxBuff(Cnt).First = QPTrim$(UBCustRec(1).CUSTNAME)
    IdxBuff(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumOfRecs&
  NEXT
  
  CLOSE UBFile
  
  QPrintRC " Sorting Customer Names", 11, 29, -1
  
  SortT IdxBuff(1), CINT(NumOfRecs&), 0, 16, 10, 2
  SortT IdxBuff(1), CINT(NumOfRecs&), 0, 16, 0, 10
  'Array(1), NumElem, Dir, StructSize, MemOff, MemSize
  
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
    LIdxBuff(Cnt).SEQNUMB = UBCustRec(1).SEQNUMB
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
    BookSeq(1).BookSeq = QPValL(LIdxBuff(Cnt).Book + LIdxBuff(Cnt).SEQNUMB)
    PUT BookHand, Cnt, BookSeq(1)
  NEXT
  CLOSE BookHand
  UBLog "FINISH: Rebuild Book\Sequence List"
  ERASE LIdxBuff, BookSeq
  
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
ExitReindex:
  UBLog "EXITED: Reindex Utility Files" + CrLf$
  EXIT SUB
  
ClearBlock:
  HideCursor
  Blank$ = SPACE$(40)
  FOR Cnt = 8 TO 15
    QPrintRC Blank$, Cnt, 21, -1
  NEXT
  ShowCursor
  RETURN
  
END SUB

SUB ReLinkHistory

  UBLog "ENTERED: Relink Transaction History"
  BlockClear
  Ok = MsgBox%("UB", "MUSTEXIT")
  
  SELECT CASE Ok
  CASE 2
    UBLog "ABORTED: Relink Transaction History"
    GOTO ExitRelink
  END SELECT
  
  'BlockClear
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))              'Length of Cust Record Structure
  
  REDIM UBTransRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTransRec(1))             'Length of Tran Record Structure
  
  ShowProcessingScrn "Relinking Transaction History"
  
  QPrintRC "        Pass: 1 of 2         ", 11, 26, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfCRecs& = LOF(UBFile) \ UBCustRecLen
  
  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen
  NumOfTRecs& = LOF(UBTran) \ UBTranRecLen

  UBLog "BEGIN: Pass 1 of 2"
  FOR CCnt& = 1 TO NumOfCRecs&
    GET UBFile, CCnt&, UBCustRec(1)
    UBCustRec(1).LastTrans = 0
    PUT UBFile, CCnt&, UBCustRec(1)
    ShowPctComp CCnt&, NumOfCRecs&
  NEXT
  UBLog "FINISH: Pass 1 of 2"

  QPrintRC "        Pass: 2 of 2         ", 11, 26, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1
  UBLog "BEGIN: Pass 2 of 2"
  FOR TCnt& = 1 TO NumOfTRecs&
    GET UBTran, TCnt&, UBTransRec(1)
    IF (UBTransRec(1).CustAcctNo > 0) AND (UBTransRec(1).CustAcctNo <= NumOfCRecs&) THEN
      GET UBFile, UBTransRec(1).CustAcctNo, UBCustRec(1)
      UBTransRec(1).PrevTrans = UBCustRec(1).LastTrans
      UBCustRec(1).LastTrans = TCnt&
      PUT UBFile, UBTransRec(1).CustAcctNo, UBCustRec(1)
      PUT UBTran, TCnt&, UBTransRec(1)
    END IF
    ShowPctCompL TCnt&, NumOfTRecs&
  NEXT
  CLOSE
  UBLog "FINISH: Pass 2 of 2"
  UBLog "RELINK: Transactions Completed."
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
ExitRelink:
  UBLog "EXITED: Relink Transaction History" + CrLf$

EXIT SUB
  
  
END SUB

SUB RelinkWorkOrders

  UBLog "ENTERED: Relink Workorder History"
  BlockClear
  Ok = MsgBox%("UB", "MUSTEXIT")

  SELECT CASE Ok
  CASE 2
    UBLog "ABORTED: Relink Workorder History"
    GOTO ExitWORelink
  END SELECT

  'BlockClear
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))              'Length of Cust Record Structure

  REDIM WorkOrderRec(1) AS WorkOrderRecType
  WorkOrderRecLen = LEN(WorkOrderRec(1))

  ShowProcessingScrn "Relinking Work Order History"

  QPrintRC "        Pass: 1 of 2         ", 11, 26, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1

  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBFile LEN = UBCustRecLen
  NumOfCRecs& = LOF(UBFile) \ UBCustRecLen

  UBWrkOrd = FREEFILE
  OPEN "UBWRKORD.DAT" FOR RANDOM SHARED AS UBWrkOrd LEN = WorkOrderRecLen
  NumOfWORecs& = LOF(UBWrkOrd) \ WorkOrderRecLen
  UBLog "BEGIN: Pass 1 of 2"
  FOR CCnt& = 1 TO NumOfCRecs&
    GET UBFile, CCnt&, UBCustRec(1)
    UBCustRec(1).WOLastTrans = 0
    PUT UBFile, CCnt&, UBCustRec(1)
    ShowPctComp CCnt&, NumOfCRecs&
  NEXT
  UBLog "FINISH: Pass 1 of 2"

  QPrintRC "        Pass: 2 of 2         ", 11, 26, -1
  QPrintRC "Processing:    % Complete", 13, 28, -1
  UBLog "BEGIN: Pass 2 of 2"
  FOR TCnt& = 1 TO NumOfWORecs&
    GET UBWrkOrd, TCnt&, WorkOrderRec(1)
    IF (WorkOrderRec(1).CustRec > 0) AND (WorkOrderRec(1).CustRec <= NumOfCRecs&) THEN
      GET UBFile, WorkOrderRec(1).CustRec, UBCustRec(1)
      WorkOrderRec(1).PrevTransRec = UBCustRec(1).WOLastTrans
      UBCustRec(1).WOLastTrans = TCnt&
      PUT UBFile, WorkOrderRec(1).CustRec, UBCustRec(1)
      PUT UBWrkOrd, TCnt&, WorkOrderRec(1)
    END IF
    ShowPctComp TCnt&, NumOfWORecs&
  NEXT
  CLOSE
  UBLog "FINISH: Pass 2 of 2"
  UBLog "RELINK: Workorders Completed."
  BlockClear
  DisplayUBScrn "UPDATEOK"
  WaitForAction

ExitWORelink:
  UBLog "EXITED: Relink Workorder History" + CrLf$
  EXIT SUB

END SUB

SUB SaveRateTblRec (Form$(), RecNo)
  
  REDIM UBRateTblRec(1) AS UBRateTblRecType
  
  UBRateTblRecLen = LEN(UBRateTblRec(1))
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBRateTblRec(1)), VARPTR(UBRateTblRec(1)), UBRateTblRecLen, 0
  
  UBFile = FREEFILE
  
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  PUT UBFile, RecNo, UBRateTblRec(1)
  CLOSE
  
  LibFile2Scrn "UB", "UPDATEOK", MonoCode, Attribute%, ErrCode
  WaitForAction
  
  ERASE UBRateTblRec
  
END SUB

SUB SaveSysDraftRec (Form$())
  
  REDIM UBSysDraftRec(1) AS UBDraftRecType
  UBSysDraftRecLen = LEN(UBSysDraftRec(1))
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBSysDraftRec(1)), VARPTR(UBSysDraftRec(1)), UBSysDraftRecLen, 0
  
  UBFile = FREEFILE
  OPEN "UBSDRAFT.DAT" FOR RANDOM SHARED AS UBFile LEN = UBSysDraftRecLen
  PUT UBFile, 1, UBSysDraftRec(1)
  CLOSE
  
  DisplayUBScrn "UPDATEOK"
  WaitForAction
  
  ERASE UBSysDraftRec
  
END SUB

SUB SelectRate2Edit (RecNo)
  
  STATIC Choice
  IF Choice = 0 THEN Choice = 1

  REDIM UBRateTblRec(1) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))
  NumOfRateRecs = GetNumRateRecs
  
  'add show a message
  IF NumOfRateRecs <= 0 THEN
    GOTO ExitRateSelect
  END IF
  
  REDIM MChoice$(1 TO NumOfRateRecs)
  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRateRecs
    GET UBFile, Cnt, UBRateTblRec(1)
    MChoice$(Cnt) = SPACE$(50)
    LSET MChoice$(Cnt) = UBRateTblRec(1).RATECODE
    MID$(MChoice$(Cnt), 6) = UBRateTblRec(1).RATEDESC
    MID$(MChoice$(Cnt), 49) = MKI$(Cnt)
  NEXT
  CLOSE
  
  SortStr2 VARPTR(MChoice$(1)), Cnt - 1, 0
  
  MaxLen = 47   'Set menu width to zero
  BoxBot = 16   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  'Choice = 1    'Pre-load choice to highlight
  
  TText$ = SPACE$(MaxLen + 4)
  LSET TText$ = " Code  Description"
  
  '--Center Menu within Screen
  Row = 6
  Col = 15
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    ShowCursor
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select ", Cnf
    QPrintRC TText$, Row - 1, Col, 112
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      RecNo = 0
      ExitFlag = True
    ELSE
      RecNo = CVI(MID$(MChoice$(Choice), 49, 2))
      ExitFlag = True
    END IF
    
  LOOP UNTIL ExitFlag
  
ExitRateSelect:
  
END SUB

SUB UBRateMaint
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 3)
  
  MChoice$(1) = "Add a New Rate Table"
  MChoice$(2) = "Edit an Existing Rate Table"
  MChoice$(3) = "Print Rate Table Listing"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  Help$ = "Customer Maintenance Menu"
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Rate Table Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      AddEditRateTbl False
    CASE 2
      AddEditRateTbl True
    CASE 3
      PrintRateListing
    END SELECT
  LOOP
  
END SUB

