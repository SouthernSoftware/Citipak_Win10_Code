  DEFINT A-Z
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
  
DECLARE SUB OPENDrawer ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB ShowTRDetail (RecNo&)
DECLARE SUB Search4Receipt ()
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB DisplayCMScrn (ScrnName$)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE FUNCTION GetDEFPort% ()
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE FUNCTION FindAcct% (AcctNum$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE FUNCTION Round# (ThisAmount#)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB HideCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB OpenMiscCodeFile (NumOfMiscRecs%, MCFile%)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintJournal (OperRecNumb%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPORT%, RetCode%, EntryPoint%)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB QPrintRC (t$, R%, c%, clr%)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB CursorOff ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB BlockClear ()
  
  TYPE FLen2
    v AS STRING * 64
  END TYPE
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ARv85.BI'
  
  '*** Citipak Include Files Needed ***
  '$INCLUDE: 'CMFILES.BI'
  '$INCLUDE: 'ubsetup.BI'
  
  TYPE Struct
    who AS STRING * 12
    RecNum AS LONG
  END TYPE
  
  CONST False = 0, True = NOT False
  
  DIM SHARED PctC(1) AS STRING * 3
  DIM SHARED ARCustRec(1) AS ARCustRecType
  DIM SHARED ARCatCodeRec(1) AS ARNewCatCodeRecType

SUB CMJourMenu (OperRecNumber)
  
  REDIM MChoice$(1 TO 3)
  
  MChoice$(1) = "Search for a Receipts"
  MChoice$(2) = "Print Receipts Journal"
  MChoice$(3) = "Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
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
  Col = ((80 - MaxLen) \ 2)
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    
    TitleBox 2, Col, MaxLen + 3, "Receipts Processing ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      Search4Receipt
    CASE 2
      PrintJournal OperRecNumber
    CASE 3
      HideCursor
      CLS
      END
    END SELECT
  LOOP
  
  
END SUB

SUB PrintJournal (OperRecNumb)
  
  REDIM TempScrn(0)
  
  SHARED User$
  LibName$ = "CM.QSL"
  ScrnName$ = "PRNTJOUR"        'form to edit
  DevFld = 6    'Set to output device field number
  
  SHARED Choice$()
  REDIM RevName$(10), TotalMiscRec$(200), TotalMiscDesc$(200), TotalMiscAmt#(200), MiscCodeGL$(200)
  REDIM TotalUtilRevAmt#(15)
  REDIM TotalDepRevAmt#(15)
  REDIM RevText$(15)
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  REDIM DistArray(1 TO 1) AS DistArrayType
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  REDIM Choice$(0 TO 6, 0 TO 2)
  
  Choice$(0, 0) = "1"           'Choices for Output Order
  Choice$(1, 0) = "Entry Order"
  Choice$(2, 0) = "Name"
  
  Choice$(0, 2) = "2"           'Choices for Journal
  Choice$(1, 2) = "ALL"
  Choice$(2, 2) = "Utility Payment"
  Choice$(3, 2) = "Misc Payment"
  Choice$(4, 2) = "Tax Payment"
  Choice$(5, 2) = "License Payment"
  Choice$(6, 2) = "Decal Payment"
  
  Choice$(0, 1) = STR$(DevFld)  'Choices for Output Device
  Choice$(1, 1) = "Screen"
  Choice$(2, 1) = "Prn#1"
  Choice$(3, 1) = "Prn#2"
  
  DisplayCMScrn ScrnName$
  
  Form$(3, 0) = DATE$
  Form$(4, 0) = DATE$
  
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM SHARED AS CMFile LEN = CMOperRecLen
  GET CMFile, OperRecNumb, CMOperRec(1)
  
  Form$(5, 0) = STR$(CMOperRec(1).OperatorNumber)
  CLOSE CMFile
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    SELECT CASE Frm(1).KeyCode  'Check the user's key presses
    CASE AltO
      OPENDrawer
    CASE F10Key 'SaveButton
      BegDate = Date2Num(Form$(3, 0))
      BegDate$ = Form$(3, 0)
      EndDate = Date2Num(Form$(4, 0))
      EndDate$ = Form$(4, 0)
      RecSource$ = UCASE$(LEFT$(Form$(2, 0), 1))
      IF RecSource$ = " " THEN RecSource$ = "A"
      OperatorNumber = VAL(Form$(5, 0))
      IF LEN(LTRIM$(RTRIM$(Form$(DevFld, 0)))) = 0 THEN
        dev$ = "S"
        LPTNo = 1
      ELSE
        dev$ = LEFT$(Form$(DevFld, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(DevFld, 0)), 1))
      END IF
      
      IF BegDate > EndDate THEN
        OK2Go = 0
        Frm(1).FldNo = 5
        Action = 1
      ELSE
        OK2Go = 1
      END IF
    CASE ESC    'ESC Key
      EXIT SUB
    END SELECT
    
  LOOP UNTIL OK2Go
  
  LOCATE , , 0
  
  IF LEFT$(Form$(1, 0), 2) = "En" THEN
    SortOrder$ = "Entry Order"
  ELSE
    SortOrder$ = "Name Order"
  END IF
  
  ERASE Form$, Fld, Choice$, Frm
  
  'End of Input
  '=====================================================
  'Start Report Processing
  
  ReportFile$ = "CMJOURNL.PRN"  'Report File Name
  Fmt1$ = "#####,#.##"
  Fmt2$ = "###,#.##"
  Fmt3$ = "$$######,#.##"
  Fmt4$ = "$$########,#.##"
  
  FF$ = CHR$(12)
  MaxLines = 53
  LineCnt = 0
  TotDr# = 0
  TotCr# = 0
  Max& = (FRE(-1) - 16000) \ 16
  Size = Max&
  
  Start = 1     'start at array element 1
  Dir = 0       'sort direction - use anything else for descending
  SSize = 16    'total size of each TYPE element
  MOff = 0      'offset into the TYPE for the key element
  MSize = 16    'size of the key element - coded as follows:
  
  '   -1 = integer
  '   -2 = long integer
  '   -3 = single precision
  '   -4 = double precision
  '   +N = TYPE array/fixed-length string of length N
  
  REDIM Array(1 TO Size) AS Struct
  REDIM CMTRRec(1) AS CMTransRecType            ' open transaction file
  CMTRRecLen = LEN(CMTRRec(1))
  TrHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CMTRRecLen
  TrNumRecs& = LOF(TrHandle) \ CMTRRecLen
  
  GOSUB GetReportInformation
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle
  
  REDIM OperRec(1) AS CMOperRecType             ' opens operatorfile
  OperRecLen = LEN(OperRec(1))
  OperFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM AS OperFile LEN = OperRecLen
  NumOperRecs = LOF(OperFile) / OperRecLen
  GET OperFile, OperRecNumb, OperRec(1)
  
  OpenMiscCodeFile NumOfMiscRecs, MCFile        ' opens misc code file
  REDIM MiscCodeRec(1) AS MiscCodeRecType
  
  PRINT #RptHandle, CHR$(27); CHR$(58);         ' oki 320 12 cpi
  
  GOSUB PrintRptHeader
  
  FOR Cnt = 1 TO Count
    GET TrHandle, Array(Cnt).RecNum, CMTRRec(1)
    IF OperatorNumber = 0 OR OperatorNumber = CMTRRec(1).TransOperNum THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintRptHeader
      END IF
      IF CMTRRec(1).TransDate >= BegDate AND CMTRRec(1).TransDate <= EndDate THEN
        TRType$ = ""
        SELECT CASE CMTRRec(1).TransSource
        CASE 1
          TRType$ = "Misc."
        CASE 27
          TRType$ = "UT-Dep."
        CASE 24
          TRType$ = "Util."
        CASE 30 TO 39
          TRType$ = "Tax"
        CASE 40 TO 49
          TRType$ = "Lic."
        END SELECT
        
        PRINT #RptHandle, Num2Date(CMTRRec(1).TransDate); TAB(12); TRType$; TAB(20); LEFT$(CMTRRec(1).TransName, 18);
        PRINT #RptHandle, TAB(40); USING Fmt1$; CMTRRec(1).TransCash; TAB(55); CMTRRec(1).TransCheck; TAB(70); CMTRRec(1).TransAmtOwed;
        
        TotalCash# = Round#(TotalCash# + CMTRRec(1).TransCash)
        TotalCheck# = Round#(TotalCheck# + CMTRRec(1).TransCheck)
        
        IF CMTRRec(1).TransSource >= 30 AND CMTRRec(1).TransSource <= 39 THEN
          TxRev# = 0
          FOR TRev = 1 TO 9
            TxRev# = Round#(TxRev# + CMTRRec(1).TransRevAmt(TRev))
          NEXT
          TotalAmount# = Round#(TotalAmount# + TxRev#)
        ELSEIF CMTRRec(1).TransSource >= 40 AND CMTRRec(1).TransSource <= 49 THEN
          IF CMTRRec(1).TransAmount < CMTRRec(1).TransAmtOwed THEN
            TotalAmount# = Round#(TotalAmount# + CMTRRec(1).TransAmount)
          ELSE
            TotalAmount# = Round#(TotalAmount# + CMTRRec(1).TransAmtOwed)
          END IF
        ELSEIF CMTRRec(1).TransSource = 1 THEN
          TotalAmount# = Round#(TotalAmount# + CMTRRec(1).TransAmtOwed)
        ELSE
          TotalAmount# = Round#(TotalAmount# + CMTRRec(1).TransCash + CMTRRec(1).TransCheck)
        END IF

        IF CMTRRec(1).TransSource = 27 THEN
          Change# = Round#((CMTRRec(1).TransCash + CMTRRec(1).TransCheck) - CMTRRec(1).TransAmtOwed)
        ELSE
          Change# = Round#(CMTRRec(1).TransAmount - CMTRRec(1).TransAmtOwed)
        END IF
        IF Change# < 0 THEN Change# = 0
        'IF Change# <> 0 THEN STOP
        TotChange# = Round#(TotChange# + Change#)
        
        PRINT #RptHandle, TAB(84); USING Fmt1$; Change#
        'PRINT #RptHandle, TAB(84); USING Fmt1$; ((CMTRRec(1).TransCheck + CMTRRec(1).TransCash) - CMTRRec(1).TransAmtOwed)
        
        TotalReceipts = TotalReceipts + 1
        LineCnt = LineCnt + 1
        
        IF CMTRRec(1).TransSource = 1 THEN
          'Second Line of Print is Misc Code Breakdown Dist.****************
          PrintMiscFlag = 0
          PRINT #RptHandle, "Description: "; CMTRRec(1).TransDesc
          FOR MCnt = 1 TO 5
            MiscRevAmt# = (CMTRRec(1).TransRevAmt(MCnt))
            MiscRevAmt# = Round#(MiscRevAmt#)
            IF MiscRevAmt# > 0 THEN
              ' If There Is an Amount in Misc Rev 1-5 then get code record number from 6-10
              IF CMTRRec(1).TransRevAmt(MCnt + 5) >= 1 THEN
                GET MCFile, CMTRRec(1).TransRevAmt(MCnt + 5), MiscCodeRec(1)
                PRINT #RptHandle, "Code BrkDwn:";
                PRINT #RptHandle, TAB(14); MiscCodeRec(1).MiscCode;
                PRINT #RptHandle, TAB(25); MiscCodeRec(1).Description;
                PRINT #RptHandle, TAB(55); USING Fmt1$; MiscRevAmt#
                PrintMiscFlag = 1
                LineCnt = LineCnt + 1
                GOSUB SubTotalMisc
              END IF
            END IF
          NEXT MCnt
          IF PrintMiscFlag = 1 THEN PRINT #RptHandle, STRING$(96, "-"): LineCnt = LineCnt + 1
          'End Misc Code Print on Second Line ****************************************
        END IF
        
        IF CMTRRec(1).TransSource >= 20 AND CMTRRec(1).TransSource <= 29 THEN
          IF CMTRRec(1).TransSource <> 27 THEN
            'Second Line of Print is Utility Breakdown Dist. *****************
            GOSUB GetRevenueSources
            IF NumOfRevs > 0 THEN
              PRINT #RptHandle, "Util BrkDwn:";
              FOR RCnt = 1 TO NumOfRevs STEP 2
                PRINT #RptHandle, TAB(15); RevText$(RCnt);
                PRINT #RptHandle, TAB(40); USING Fmt1$; CMTRRec(1).TransRevAmt(RCnt);
                PRINT #RptHandle, TAB(55); RevText$(RCnt + 1);
                PRINT #RptHandle, TAB(80); USING Fmt1$; CMTRRec(1).TransRevAmt(RCnt + 1)
                PrintUtilFlag = 1
                LineCnt = LineCnt + 1
              NEXT RCnt
              GOSUB SubTotalUtil
            END IF
            IF PrintUtilFlag = 1 THEN PRINT #RptHandle, STRING$(96, "-"): LineCnt = LineCnt + 1
            'End of Utility Print on Second Line *****************************
          ELSE
            GOSUB SubTotalDep
            PRINT #RptHandle, STRING$(96, "-")
            LineCnt = LineCnt + 1
          END IF
        END IF
        
        IF CMTRRec(1).TransSource >= 30 AND CMTRRec(1).TransSource <= 39 THEN
          'Second Line of Print is Tax Breakdown Dist.     *****************
          PRINT #RptHandle, "Tax BrkDwn:";
          PRINT #RptHandle, TAB(15); "Tax:"; USING Fmt1$; CMTRRec(1).TransRevAmt(1);
          PRINT #RptHandle, TAB(32); "Int.: "; USING Fmt2$; CMTRRec(1).TransRevAmt(2);
          PRINT #RptHandle, TAB(50); "Pen.: "; USING Fmt2$; CMTRRec(1).TransRevAmt(3);
          PRINT #RptHandle, TAB(65); "Strm: "; USING Fmt2$; CMTRRec(1).TransRevAmt(4)
          PRINT #RptHandle, TAB(5); " Past Tax: "; USING Fmt1$; CMTRRec(1).TransRevAmt(6);
          PRINT #RptHandle, TAB(32); "Int.: "; USING Fmt2$; CMTRRec(1).TransRevAmt(7);
          PRINT #RptHandle, TAB(50); "Pen.: "; USING Fmt2$; CMTRRec(1).TransRevAmt(8);
          PRINT #RptHandle, TAB(65); "Strm: "; USING Fmt2$; CMTRRec(1).TransRevAmt(9)
          PrintTaxFlag = 1
          LineCnt = LineCnt + 2
          GOSUB SubTotalTax
        END IF
        IF PrintTaxFlag = 1 THEN
          PRINT #RptHandle, STRING$(96, "-")
          LineCnt = LineCnt + 1
        END IF
        'End of Tax Print on Second Line *********************************
        
        IF CMTRRec(1).TransSource >= 40 AND CMTRRec(1).TransSource <= 49 THEN
          GOSUB SubTotalBL
        END IF
        IF CMTRRec(1).TransSource >= 50 AND CMTRRec(1).TransSource <= 59 THEN
          GOSUB SubTotalDC
        END IF
      END IF
    END IF
    
  NEXT Cnt
  
  GOSUB PrintRptEnding
  PRINT #RptHandle, CHR$(18);   ' oki 320 12 cpi
  
  CLOSE         'Close all open files now
  
  SELECT CASE dev$
  CASE "S"
    EntryPoint = 2
  CASE "P"
    EntryPoint = 5
  END SELECT
  
  ERASE RevName$, TotalMiscRec$, TotalMiscDesc$, TotalMiscAmt#
  ERASE TotalUtilRevAmt#, Frm, Form$, Fld, MiscCodeGL$
  ERASE Array, CMTRRec, RevText$, MiscCodeRec, UBSetUpRec
  ERASE DistArray, OperRec
  
  PrintRptFile Header$, ReportFile$, LPTNo, RetCode%, EntryPoint
  
  'KILL ReportFile$
  
  EXIT SUB
  
PrintRptHeader:
  page = page + 1
  PRINT #RptHandle, TAB(27); "Cash Receipts Journal : Cash Management System"
  PRINT #RptHandle, "Beginning Date: "; BegDate$
  PRINT #RptHandle, "   Ending Date: "; EndDate$
  PRINT #RptHandle, "      Operator: "; OperRec(1).OperatorName$; TAB(83); "Page #"; page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "   Date"; TAB(12); "Source"; TAB(20); "Description"; TAB(40); "Cash Amount"; TAB(55); "Check Amount"; TAB(70); "Amount Owed"; TAB(84); "Change Given"
  
  PRINT #RptHandle, STRING$(96, "=")
  LineCnt = 5
RETURN
  
PrintRptEnding:
  PRINT #RptHandle, FF$
  PRINT #RptHandle, "Totals Page for Operator # "; OperatorNumber
  PRINT #RptHandle, ""
  PRINT #RptHandle, "   Total Cash Received: "; USING Fmt3$; TotalCash#
  PRINT #RptHandle, " Total Checks Received: "; USING Fmt3$; TotalCheck#
  PRINT #RptHandle, "                      -----------------"
  PRINT #RptHandle, "  Total Money Received: "; USING Fmt3$; TotalCash# + TotalCheck#
  PRINT #RptHandle, "Amount Applied to Acct: "; USING Fmt3$; (TotalCash# + TotalCheck#) - TotChange#
  PRINT #RptHandle, "    Total Change Given: "; USING Fmt3$; TotChange#
  PRINT #RptHandle, ""
  PRINT #RptHandle, "   Bank Deposit Amount: "; USING Fmt3$; (TotalCash# + TotalCheck#) - TotChange#
  PRINT #RptHandle, "    Number of Receipts: "; USING "####,#"; TotalReceipts
  
  LineCnt = 12
  IF RecSource$ = "M" OR RecSource$ = "A" THEN
    GOSUB PrintTotalMisc
  END IF
  
  IF RecSource$ = "U" OR RecSource$ = "A" THEN
    GOSUB PrintTotalUtil
  END IF
  
  IF RecSource$ = "T" OR RecSource$ = "A" THEN
    GOSUB PrintTotalTax
  END IF
  
  IF RecSource$ = "L" OR RecSource$ = "A" THEN
    GOSUB PrintBLTotal          ' Not Active Yet!!!
  END IF
  
  IF RecSource$ = "D" OR RecSource$ = "A" THEN
    GOSUB PrintDCTotal
  END IF
  
  PRINT #RptHandle, FF$
RETURN
  
GetReportInformation:
  BegRecNo& = TrNumRecs& - 7500                ' Move back 7500 records to begin
  IF BegRecNo& < 1 THEN BegRecNo& = 1         ' Don't Allow Less Than 1
  IF OperatorNumber = 0 THEN BegRecNo& = 1
  
  FOR Cnt! = BegRecNo& TO TrNumRecs&
    GET TrHandle, Cnt!, CMTRRec(1)
    TransDate = CMTRRec(1).TransDate
    GoodRecordFlag = False
    
    'IF CMTRRec(1).TransOperNum = OperRecNumb AND (TransDate >= BegDate AND TransDate <= EndDate) THEN
    IF (TransDate >= BegDate AND TransDate <= EndDate) THEN
      
      IF RecSource$ = "A" THEN
        GoodRecordFlag = True
      END IF
      IF RecSource$ = "M" AND CMTRRec(1).TransSource = 1 THEN
        GoodRecordFlag = True
      END IF
      IF RecSource$ = "U" AND CMTRRec(1).TransSource >= 20 AND CMTRRec(1).TransSource <= 29 THEN
        GoodRecordFlag = True
      END IF
      IF RecSource$ = "T" AND CMTRRec(1).TransSource >= 30 AND CMTRRec(1).TransSource <= 39 THEN
        GoodRecordFlag = True
      END IF
      
      IF RecSource$ = "L" AND CMTRRec(1).TransSource >= 40 AND CMTRRec(1).TransSource <= 49 THEN
        GoodRecordFlag = True
      END IF
      
      IF RecSource$ = "D" AND CMTRRec(1).TransSource >= 50 AND CMTRRec(1).TransSource <= 59 THEN
        GoodRecordFlag = True
      END IF
      
    END IF
    
    IF GoodRecordFlag THEN
      Count = Count + 1
      IF Count > Size THEN
        SaveScrn TempScrn()
        DisplayUBScrn "ERRSCRN1"
        QPrintRC "TO MANY TRANSACTIONS!", 10, 30, -1
        UseW$ = "Will Display First:" + STR$(Count) + " Transactions."
        OffSet = ((80 - LEN(UseW$)) / 2)
        QPrintRC UseW$, 11, OffSet, -1
        QPrintRC "Press any key to continue!", 13, 28, -1
        WaitForAction
        RestScrn TempScrn()
        Count = Size
        EXIT FOR
      END IF
      
      'PrintHelp Help$ + " Count:" + STR$(Count)
      IF SortOrder$ = "Entry Order" THEN
        Array(Count).who = STR$(Cnt!)
      ELSE
        Array(Count).who = LEFT$(CMTRRec(1).TransName, 12)
      END IF
      Array(Count).RecNum = Cnt!
    END IF
    
  NEXT Cnt!
  
  SortT Array(Start), Count, Dir, SSize, MOff, MSize
RETURN
  
GetRevenueSources:
  
  NumOfRevs = MaxRevsCnt
  REDIM UBSetUpRec(1) AS UBSetupRecType
  REDIM DistArray(1 TO MaxRevsCnt) AS DistArrayType
  
  UBSetupLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
  
  FOR RevCnt = 1 TO MaxRevsCnt
    RevText$(RevCnt) = LEFT$(QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName), 14)
    DistArray(RevCnt).DistOrder = UBSetUpRec(1).Revenues(RevCnt).DistOr
    DistArray(RevCnt).DistCnt = RevCnt
    IF LEN(RevText$(RevCnt)) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    END IF
  NEXT
  
  REDIM PRESERVE DistArray(1 TO NumOfRevs) AS DistArrayType
  
  DO
    OutOfOrder = False          'assume it's sorted
    FOR X = 1 TO NumOfRevs - 1
      IF DistArray(X).DistOrder > DistArray(X + 1).DistOrder THEN
        SWAP DistArray(X), DistArray(X + 1)     'if we had to swap
        OutOfOrder = True       'we're not done yet
      END IF
    NEXT
  LOOP WHILE OutOfOrder
  
  TOWNNAME$ = UBSetUpRec(1).UTILNAME
RETURN
  
' ** old way
  
SubTotalUtil:
  FOR uCnt! = 1 TO NumOfRevs
    TotalUtilRevAmt#(uCnt!) = TotalUtilRevAmt#(uCnt!) + CMTRRec(1).TransRevAmt(uCnt!)
  NEXT uCnt!
RETURN
  
SubTotalDep:
  FOR DCnt = 1 TO 15
    TotalDepRevAmt#(DCnt) = Round#(TotalDepRevAmt#(DCnt) + CMTRRec(1).TransRevAmt(DCnt))
  NEXT
RETURN
  
PrintTotalUtil:
  PRINT #RptHandle, STRING$(96, "-")
  PRINT #RptHandle, "GRAND TOTAL Utilities Receipts Recap"
  IF LineCnt >= MaxLines - (NumOfRevs + 5) THEN
    PRINT #RptHandle, FF$
    PRINT #RptHandle, ""
    PRINT #RptHandle, "Totals Page for Operator # "; OperatorNumber
    PRINT #RptHandle, "Total Utilities Receipts Cont'd"
    PRINT #RptHandle, ""
    LineCnt = 4
  END IF
  FOR Cnt! = 1 TO NumOfRevs
    PRINT #RptHandle, RevText$(Cnt!); TAB(34); USING Fmt4$; TotalUtilRevAmt#(Cnt!)
    TotalUtilAmt# = TotalUtilAmt# + TotalUtilRevAmt#(Cnt!)
  NEXT Cnt!
  PRINT #RptHandle, "GRAND Total Utility Receipts ... "; USING Fmt4$; TotalUtilAmt#
  PRINT #RptHandle,
  
  TotalDepAmt# = 0
  FOR Cnt! = 1 TO 15
    'PRINT #RptHandle, RevText$(Cnt!); TAB(34); USING Fmt4$; TotalDepRevAmt#(Cnt!)
    TotalDepAmt# = Round#(TotalDepAmt# + TotalDepRevAmt#(Cnt!))
  NEXT Cnt!
  PRINT #RptHandle, "GRAND Total Utility Deposits ... "; USING Fmt4$; TotalDepAmt#
RETURN
  
SubTotalTax:
  TotCurTax# = TotCurTax# + CMTRRec(1).TransRevAmt(1)
  TotCurInt# = TotCurInt# + CMTRRec(1).TransRevAmt(2)
  TotCurPen# = TotCurPen# + CMTRRec(1).TransRevAmt(3)
  TotStrmFee# = TotStrmFee# + CMTRRec(1).TransRevAmt(4)
  TotPastTax# = TotPastTax# + CMTRRec(1).TransRevAmt(6)
  TotPastInt# = TotPastInt# + CMTRRec(1).TransRevAmt(7)
  TotPastPen# = TotPastPen# + CMTRRec(1).TransRevAmt(8)
  TotPastStrm# = TotPastStrm# + CMTRRec(1).TransRevAmt(9)
RETURN
  
PrintTotalTax:
  IF (TotCurTax# + TotCurInt# + TotCurPen# + TotPastTax# + TotPastInt# + TotPastPen# + TotStrmFee# + TotPastStm#) = 0 THEN RETURN
  PRINT #RptHandle, STRING$(96, "-")
  PRINT #RptHandle, "GRAND TOTAL Tax Receipts Recap"
  IF LineCnt >= MaxLines - 6 THEN
    PRINT #RptHandle, FF$
    PRINT #RptHandle, ""
    PRINT #RptHandle, "Totals Page for Operator # "; OperRec(1).OperatorNumber
    PRINT #RptHandle, "Total Tax Receipts Cont'd"
    PRINT #RptHandle, ""
    LineCnt = 4
  END IF
  PRINT #RptHandle, "Total Current Taxes Received ..... "; USING Fmt3$; TotCurTax#
  PRINT #RptHandle, "Total Current Interest Received .. "; USING Fmt3$; TotCurInt#
  PRINT #RptHandle, "Total Current Penalty Received ... "; USING Fmt3$; TotCurPen#
  PRINT #RptHandle, "Total Storm Fee Received ......... "; USING Fmt3$; TotStrmFee#
  PRINT #RptHandle, "Total Past Taxes Received ........ "; USING Fmt3$; TotPastTax#
  PRINT #RptHandle, "Total Past Interest Received ..... "; USING Fmt3$; TotPastInt#
  PRINT #RptHandle, "Total Past Penalty Received ...... "; USING Fmt3$; TotPastPen#
  PRINT #RptHandle, "Total Past Storm Fee Received .... "; USING Fmt3$; TotPastStrm#
  PRINT #RptHandle, "GRAND Total Tax Received ......... "; USING Fmt3$; (TotCurTax# + TotCurInt# + TotCurPen# + TotPastTax# + TotPastInt# + TotPastPen# + TotStrmFee# + TotPastStm#)
  PRINT #RptHandle, ""
  LineCnt = LineCnt + 8
RETURN
  
SubTotalBL:
  IF CMTRRec(1).TransAmount < CMTRRec(1).TransAmtOwed THEN
    TotalBLAmt# = TotalBLAmt# + CMTRRec(1).TransAmount
  ELSE
    TotalBLAmt# = TotalBLAmt# + CMTRRec(1).TransAmtOwed
  END IF
  TotalBLAmt# = Round#(TotalBLAmt#)
RETURN
  
PrintBLTotal:
  PRINT #RptHandle, STRING$(96, "-")
  PRINT #RptHandle, "GRAND TOTAL Business Licence Receipts Recap"
  PRINT #RptHandle, "GRAND Total Bus. Lic. Receipts .. "; TAB(35); USING "$$#######,#.##"; TotalBLAmt#
RETURN
  
SubTotalDC:
  TotalDCAmt# = TotalDCAmt# + CMTRRec(1).TransAmtOwed
  TotalDCAmt# = Round#(TotalDCAmt#)
RETURN
  
PrintDCTotal:
  PRINT #RptHandle, STRING$(96, "-")
  PRINT #RptHandle, "GRAND TOTAL Vehicle Decals Receipts Recap"
  PRINT #RptHandle, "GRAND Total Veh. Dec. Receipts .. "; TAB(35); USING "$$#######,#.##"; TotalDCAmt#
RETURN
  
SubTotalMisc:
  IF TotalMiscCnt! = 0 THEN
    TotalMiscCnt! = 1
    TotalMiscRec$(1) = MiscCodeRec(1).MiscCode
    TotalMiscDesc$(1) = MiscCodeRec(1).Description
    TotalMiscAmt#(1) = MiscRevAmt#
    MiscCodeGL$(1) = MiscCodeRec(1).GlAcctNumb
  ELSE
    FOR TCnt! = 1 TO TotalMiscCnt!
      IF MiscCodeRec(1).MiscCode = TotalMiscRec$(TCnt!) THEN
        TotalMiscAmt#(TCnt!) = TotalMiscAmt#(TCnt!) + MiscRevAmt#: RETURN
      END IF
    NEXT TCnt!
    TotalMiscCnt! = TotalMiscCnt! + 1
    TotalMiscRec$(TotalMiscCnt!) = MiscCodeRec(1).MiscCode
    TotalMiscDesc$(TotalMiscCnt!) = MiscCodeRec(1).Description
    TotalMiscAmt#(TotalMiscCnt!) = MiscRevAmt#
    MiscCodeGL$(TotalMiscCnt!) = MiscCodeRec(1).GlAcctNumb
  END IF
RETURN
  
PrintTotalMisc:
  PRINT #RptHandle, STRING$(96, "-")
  PRINT #RptHandle, "GRAND TOTAL Misc Receipts Recap"
  FOR Cnt! = 1 TO TotalMiscCnt!
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      PRINT #RptHandle, ""
      PRINT #RptHandle, "Totals Page for Operator # "; OperRec(1).OperatorNumber
      PRINT #RptHandle, "Total Misc Receipts Recap Cont'd"
      PRINT #RptHandle, ""
      LineCnt = 4
    END IF
    PRINT #RptHandle, TotalMiscDesc$(Cnt!);
    PRINT #RptHandle, TAB(35); USING "$$#######,#.##"; TotalMiscAmt#(Cnt!);
    PRINT #RptHandle, TAB(52); "GL# "; MiscCodeGL$(Cnt!)
    TotalMisc# = TotalMisc# + TotalMiscAmt#(Cnt!)
    LineCnt = LineCnt + 1
  NEXT Cnt!
  PRINT #RptHandle, "GRAND Total Misc Receipts .... "; TAB(35); USING "$$#######,#.##"; TotalMisc#
  LineCnt = LineCnt + 1
RETURN

OpenARFile:
  CLOSE ARFile
  ARCustRecLen = LEN(ARCustRec(1))
  ARFile = FREEFILE
  OPEN "ARCUST.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS ARFile LEN = ARCustRecLen
  NumOfArRecs = LOF(ARFile) \ ARCustRecLen
  CatFile = FREEFILE
  OPEN "ARCODE.DAT" FOR RANDOM AS CatFile LEN = CatCodeRecLen
  NumOfCatRecs = LOF(CatFile) \ CatCodeRecLen
  RETURN
  
END SUB

SUB PrintReceipt
  
  Fmt$ = "$$####,#.##"
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
  
  IF INSTR(UBSetUpRec(1).UTILNAME, "SPENCER") THEN
    SpencerFlag = True
  END IF
  Work$ = SPACE$(19)
  
  TheTime$ = TIME$
  IF LEFT$(TheTime$, 1) = "0" THEN
    Hour = VAL(MID$(TheTime$, 2, 1))
  ELSE
    Hour = VAL(MID$(TheTime$, 1, 2))
  END IF
  
  SELECT CASE Hour
  CASE IS > 12
    Hour = Hour - 12
    AmPm$ = "pm"
  CASE 1 TO 12
    AmPm$ = "am"
  CASE 0
    Hour = 12
    AmPm$ = "am"
  END SELECT
  SELECT CASE Hour
  CASE 1 TO 9
    Hour$ = "0" + QPTrim$(STR$(Hour))
  CASE ELSE
    Hour$ = QPTrim$(STR$(Hour))
  END SELECT
  TheTime$ = Hour$ + ":" + MID$(TheTime$, 4) + AmPm$
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
  TOWNNAME$ = UBSetUpRec(1).UTILNAME
  
  REDIM RMRec(1) AS RMReceiptRecType
  RMRecLen = LEN(RMRec(1))
  RMFile = FREEFILE
  OPEN "CMRECPT.DAT" FOR RANDOM SHARED AS RMFile LEN = RMRecLen
  GET RMFile, 1, RMRec(1)
  CLOSE RMFile
  
  RctFile = FREEFILE
  OPEN "recpt.prn" FOR OUTPUT SHARED AS #RctFile
  WIDTH #RctFile, 255
  
  PRINT #RctFile, CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)
  PRINT #RctFile, CHR$(7)
  PRINT #RctFile, TAB(12); TOWNNAME$
  PRINT #RctFile, TAB(12); "  MISC PAYMENT"
  PRINT #RctFile, TAB(13); "Date: "; RMRec(1).RecDate
  PRINT #RctFile,
  PRINT #RctFile, "CUSTOMER NAME & DESC. OF PAYMENT"
  PRINT #RctFile, RMRec(1).RecName
  PRINT #RctFile, RMRec(1).RecAddress
  PRINT #RctFile, RMRec(1).RecDesc
  PRINT #RctFile,
  PRINT #RctFile, "Total Owed: "; USING Fmt$; RMRec(1).RecAmtOwed
  PRINT #RctFile, "  Cash Amt: "; USING Fmt$; RMRec(1).RecCashAmt
  PRINT #RctFile, " Check Amt: "; USING Fmt$; RMRec(1).RecCheckAmt
  PRINT #RctFile, "Total Paid: "; USING Fmt$; RMRec(1).RecCashAmt + RMRec(1).RecCheckAmt
  PRINT #RctFile, "    Change: "; USING Fmt$; RMRec(1).RecChangeDue
  PRINT #RctFile,
  
  'put spencer trap here
  IF SpencerFlag THEN
    PRINT #RctFile, CHR$(27); "="; CHR$(2)      'only to display
    LSET Work$ = "THANK YOU!"
    PRINT #RctFile, Work$
    LSET Work$ = "Change:" + FUsing$(STR$(RMRec(1).RecChangeDue), "$$####.##")
    PRINT #RctFile, Work$; CHR$(10)
    PRINT #RctFile, CHR$(27); "="; CHR$(1)      'only to printer
  END IF
  
  PRINT #RctFile,
  PRINT #RctFile, "Operator: "; RMRec(1).RecOperator
  PRINT #RctFile, "Receipt#: "; USING "######"; RMRec(1).RecptNumber
  PRINT #RctFile,
  PRINT #RctFile, TAB(7); "T H A N K   Y O U !"
  PRINT #RctFile,
  PRINT #RctFile, "DATE: "; DATE$
  PRINT #RctFile, "TIME: "; TheTime$
  PRINT #RctFile,
  PRINT #RctFile,
  PRINT #RctFile,
  PRINT #RctFile,
  PRINT #RctFile,
  PRINT #RctFile,
  PRINT #RctFile,
  
  '27 "p" + 0   drawer 1
  '27 "p" + 1   drawer 2
  
  CLOSE #RctFile
  
  PrintRptFile Header$, "Recpt.PRN", GetDEFPort%, RetCode%, 5
  
  'SHELL "Type Recpt.PRN > lpt2:"
  
  '  EXIT SUB
  
END SUB

SUB Search4Receipt
  
  u$ = CHR$(24)
  d$ = CHR$(25)
  Escape$ = CHR$(27)
  
ResetEntry:
  LibName$ = "CM.QSL"
  ScrnName$ = "SCHRECPT"        'form to edit
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  REDIM CMTRRec(1) AS CMTransRecType            ' open transaction file
  CMTRRecLen = LEN(CMTRRec(1))
  NumOfRept& = FileSize&("CMTRANS.DAT") / CMTRRecLen
  FirstTime = True
  
Top:
  OK2Go = False
  Action = 1
  BlockClear
  DisplayCMScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(4, 0) = DATE$
      LSET Form$(5, 0) = "1"
      LSET Form$(6, 0) = QPTrim$(STR$(NumOfRept&))
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode  'Check the user's key presses
    CASE F10Key 'Do it
      GOSUB GetParms
      OK2Go = True
    CASE F5KEY
      GOTO ResetEntry
    CASE AltO
      OPENDrawer
      
    CASE ESC    'ESC Key
      EXIT SUB
    END SELECT
  LOOP UNTIL OK2Go
  
  Choice = 1
  
  CursorOff
  BlockClear
  
  ShowProcessingScrn "Searching Transactions"
  
  Max& = (FRE(-1) - 16000) \ 64
  IF Max& > 3500 THEN
    Max& = 3500
  END IF
  
  REDIM TRList(1 TO 1)    AS FLen2
  
  MatchCnt& = 0
  CMFile = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CMFile LEN = CMTRRecLen
  NumOfRecs& = LOF(CMFile) \ CMTRRecLen
  IF NumOfRecs& > 0 THEN
    FOR Cnt& = 1 TO NumOfRecs&
      GET CMFile, Cnt&, CMTRRec(1)
      GOSUB CheckForMatch
      IF GotOne THEN
        MatchCnt& = MatchCnt& + 1
        REDIM PRESERVE TRList(1 TO MatchCnt&) AS FLen2
        LSET TRList(MatchCnt&).v = Num2Date$(CMTRRec(1).TransDate)
        MID$(TRList(MatchCnt&).v, 12) = LEFT$(CMTRRec(1).TransName, 20)
        MID$(TRList(MatchCnt&).v, 34) = FUsing$(STR$(CMTRRec(1).TransAmount), "######.##")
        'Change# = Round#((CMTRRec(1).TransCheck + CMTRRec(1).TransCash) - CMTRRec(1).TransAmtOwed)
        'Change# = Round#((CMTRRec(1).TransCheck + CMTRRec(1).TransCash) - CMTRRec(1).TransAmount) 'tOwed
        
        SELECT CASE CMTRRec(1).TransSource
        CASE 1
          TRType$ = "Misc"
          Change# = Round#((CMTRRec(1).TransCheck + CMTRRec(1).TransCash) - CMTRRec(1).TransAmtOwed)
        CASE 20 TO 29
          TRType$ = "Util"
          TPaid# = 0
          FOR TRev = 1 TO 15
            TPaid# = Round#(TPaid# + CMTRRec(1).TransRevAmt(TRev))
          NEXT
          Change# = Round#((CMTRRec(1).TransCheck + CMTRRec(1).TransCash) - TPaid#)
        CASE 30 TO 39
          TRType$ = " Tax"
          Change# = Round#((CMTRRec(1).TransCheck + CMTRRec(1).TransCash) - CMTRRec(1).TransAmtOwed)
        CASE 40 TO 49
          TRType$ = " Lic"
          Change# = Round#((CMTRRec(1).TransCheck + CMTRRec(1).TransCash) - CMTRRec(1).TransAmtOwed)
        CASE 50 TO 60
          TRType$ = "Decl"
        CASE ELSE
          TRType$ = FUsing$(STR$(CMTRRec(1).TransSource), "####")
        END SELECT
        
        IF Change# < 0 THEN Change# = 0
        MID$(TRList(MatchCnt&).v, 44) = FUsing$(STR$(Change#), "######.##")
        MID$(TRList(MatchCnt&).v, 56) = TRType$
        MID$(TRList(MatchCnt&).v, 61) = MKL$(Cnt&)
        IF MatchCnt& >= Max& THEN
          BlockClear
          DisplayCMScrn "TOMANY"
          QPrintRC STR$(MatchCnt&) + " Transactions.", 11, 39, -1
          WaitForAction
          EXIT FOR
        END IF
      END IF
      ShowPctCompL Cnt&, NumOfRecs&
      QPrintRC "Found:" + STR$(MatchCnt&), 11, 33, -1
    NEXT
  END IF
  CLOSE
  
  IF MatchCnt& = 0 THEN
    BlockClear
    OK = MsgBox%("CM.QSL", "NOMATCH")           'action variable reset at goto top
    GOTO Top
  END IF
  
VertMenuEntry:
  MaxLen = 59
  Col = ((80 - MaxLen) \ 2) - 1
  BlockClear
  BoxBot = 16
  Action = 0
  
  Text$ = SPACE$(63)
  LSET Text$ = "  Tran Date      Customer             TR Amt    Change   Src"
  QPrintRC Text$, 4, 9, 112
  
  WazzWind BoxBot + 3, Col, BoxBot + 6, MaxLen + 3 + Col, 10, 4, True
  QPrintRC "  Use:  " + u$ + "-" + d$ + " to select.", BoxBot + 4, Col + 3, 15
  QPrintRC u$, BoxBot + 4, Col + 11, 14
  QPrintRC d$, BoxBot + 4, Col + 13, 14
  QPrintRC "Total: " + STR$(MatchCnt&), BoxBot + 5, Col + 3, 15
  QPrintRC "Press:   [ESC] to continue.", BoxBot + 4, Col + 33, 15
  QPrintRC "        [ENTER] for detail.", BoxBot + 5, Col + 33, 15
  QPrintRC "ESC", BoxBot + 4, Col + 43, 14
  QPrintRC "ENTER", BoxBot + 5, Col + 42, 14
  
  LOCATE 5, Col
  
  VertMenuT2 TRList(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
  IF Ky$ <> Escape$ THEN
    RecNo& = CVL(RIGHT$(TRList(Choice).v, 4))
    ShowTRDetail RecNo&
    GOTO VertMenuEntry
  END IF
  
  ERASE TRList
  Frm(1).FldNo = 1
  GOTO Top
  
CheckForMatch:
  GotOne = False
  IF UseCust THEN
    IF INSTR(CMTRRec(1).TransName, CustName$) > 0 THEN
      GotOne = True
    ELSE
      GotOne = False
      GOTO NoGood
    END IF
  END IF
  
  IF UseAcct THEN
    IF AcctNo& = CMTRRec(1).TransAcctNum THEN
      GotOne = True
    ELSE
      GotOne = False
      GOTO NoGood
    END IF
  END IF
  
  IF UseDFrom THEN
    IF FromDate <= CMTRRec(1).TransDate THEN
      GotOne = True
    ELSE
      GotOne = False
      GOTO NoGood
    END IF
  END IF
  
  IF UseDThru THEN
    IF ThruDate >= CMTRRec(1).TransDate THEN
      GotOne = True
    ELSE
      GotOne = False
      GOTO NoGood
    END IF
  END IF
  
  IF UseRFrom THEN
    IF FromRecp& <= Cnt& THEN
      GotOne = True
    ELSE
      GotOne = False
      GOTO NoGood
    END IF
  END IF
  
  IF UseRThru THEN
    IF ThruRecp& >= Cnt& THEN
      GotOne = True
    ELSE
      GotOne = False
      GOTO NoGood
    END IF
  END IF
  
NoGood:
  RETURN
  
GetParms:
  UseCust = False
  UseAcct = False
  UseDFrom = False
  UseDThru = False
  UseRFrom = False
  UseRThru = False
  
  CustName$ = QPTrim$(Form$(1, 0))
  IF LEN(CustName$) > 0 THEN
    UseCust = True
  END IF
  AcctNo& = Value(Form$(2, 0), ecode)
  IF AcctNo& > 0 THEN
    UseAcct = True
  END IF
  FromDate = Date2Num(Form$(3, 0))
  ThruDate = Date2Num(Form$(4, 0))
  IF FromDate > 0 THEN
    UseDFrom = True
  END IF
  IF ThruDate > 0 THEN
    UseDThru = True
  END IF
  FromRecp& = Value(Form$(5, 0), ecode)
  ThruRecp& = Value(Form$(6, 0), ecode)
  IF FromRecp& > 0 THEN
    UseRFrom = True
  END IF
  IF ThruRecp& > 0 THEN
    UseRThru = True
  END IF
  
  RETURN
  
END SUB

SUB ShowTRDetail (RecNo&)
  
  CursorOff

  Fmt1$ = "######.##"
  Fmt2$ = "######"
  
  REDIM CMTrans(1) AS CMTransRecType
  CMTRLen = LEN(CMTrans(1))
  CMFile = FREEFILE

  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS CMFile LEN = CMTRLen
  GET CMFile, RecNo&, CMTrans(1)
  'CMTrans(1).TransDate = 8345
  'PUT CMFile, RecNo&, CMTrans(1)
  'CMTrans(1).TransCheck = 0
  'CMTrans(1).TransAmount = CMTrans(1).TransCheck + CMTrans(1).TransCash

  CLOSE
  
  Change# = Round#((CMTrans(1).TransCheck + CMTrans(1).TransCash) - CMTrans(1).TransAmount)     'tOwed
  
  SELECT CASE CMTrans(1).TransSource
  CASE 1
    TRType$ = "Miscellaneous"
    MiscFlag = True
    GOSUB GetMiscCode
    Change# = Round#((CMTrans(1).TransCheck + CMTrans(1).TransCash) - CMTrans(1).TransAmtOwed)
  CASE 20 TO 29
    TRType$ = "Utility Billing"
    TPaid# = 0
    FOR TRev = 1 TO 15
      TPaid# = Round#(TPaid# + CMTrans(1).TransRevAmt(TRev))
    NEXT
    Change# = Round#((CMTrans(1).TransCheck + CMTrans(1).TransCash) - TPaid#)
  CASE 30 TO 39
    TRType$ = "Property Taxs"
    Change# = Round#((CMTrans(1).TransCheck + CMTrans(1).TransCash) - CMTrans(1).TransAmtOwed)
  CASE 40 TO 49
    TRType$ = "Business License"
    Change# = Round#((CMTrans(1).TransCheck + CMTrans(1).TransCash) - CMTrans(1).TransAmtOwed)
  CASE 50 TO 60
    DecalFlag = True
    TRType$ = "Vehicle Decals"
    CMTrans(1).TransRevAmt(1) = 0
  CASE ELSE
    TRType$ = FUsing$(STR$(CMTrans(1).TransSource), "####")
  END SELECT
  
  'Change# = Round#((CMTrans(1).TransCheck + CMTrans(1).TransCash) - CMTrans(1).TransAmtOwed)
  IF Change# < 0 THEN Change# = 0
  
  DisplayCMScrn "TRDETAIL"
  QPrintRC CMTrans(1).TransName, 5, 31, -1
  '  IF MiscFlag THEN
  '    QPrintRC Code$, 6, 31, -1
  '  ELSE
  QPrintRC CMTrans(1).TransDesc, 6, 31, -1
  '  END IF
  QPrintRC Num2Date(CMTrans(1).TransDate), 7, 31, -1
  QPrintRC TRType$, 8, 31, -1

  'QPrintRC CMTrans(1).TransDesc, 8, 1, -1
  QPrintRC FUsing$(STR$(CMTrans(1).TransAmount), Fmt1$), 10, 63, -1
  QPrintRC FUsing$(STR$(CMTrans(1).TransAmtOwed), Fmt1$), 11, 63, -1
  QPrintRC FUsing$(STR$(CMTrans(1).TransCash), Fmt1$), 12, 63, -1
  QPrintRC FUsing$(STR$(CMTrans(1).TransCheck), Fmt1$), 13, 63, -1
  QPrintRC FUsing$(STR$(Change#), Fmt1$), 14, 63, -1
  QPrintRC FUsing$(STR$(CMTrans(1).TransOperNum), Fmt2$), 15, 63, -1
  IF CMTrans(1).Trans2GL = "Y" THEN
    GL$ = "Y"
  ELSE
    GL$ = "N"
  END IF
  QPrintRC GL$, 16, 68, -1
  QPrintRC FUsing$(STR$(RecNo&), Fmt2$), 17, 63, -1
  
  FOR zz = 1 TO 8
    QPrintRC FUsing$(STR$(CMTrans(1).TransRevAmt(zz)), Fmt1$), zz + 10, 8, -1
  NEXT
  FOR zz = 9 TO 15
    QPrintRC FUsing$(STR$(CMTrans(1).TransRevAmt(zz)), Fmt1$), zz + 2, 26, -1
  NEXT
  WaitForAction

EXIT SUB
  
GetMiscCode:
  Code$ = ""
  REDIM MiscCodeRec(1) AS MiscCodeRecType
  MiscCodeRecLen = LEN(MiscCodeRec(1))
  FOR zz = 6 TO 10
    CodeRec& = CLNG(CMTrans(1).TransRevAmt(zz))
    IF CodeRec& > 0 THEN
      MCFile = FREEFILE
      OPEN "CMMISCCD.DAT" FOR RANDOM SHARED AS MCFile LEN = MiscCodeRecLen
      NumOfMiscRecs = LOF(MCFile) \ MiscCodeRecLen
      IF NumOfMiscRecs > CodeRec& THEN
        GET MCFile, CodeRec&, MiscCodeRec(1)
      END IF
      CLOSE
      Code$ = Code$ + QPTrim$(MiscCodeRec(1).MiscCode) + " "
      CMTrans(1).TransRevAmt(zz) = 0            'zero pointer to misccode record
    END IF
  NEXT
  RETURN
  
END SUB

