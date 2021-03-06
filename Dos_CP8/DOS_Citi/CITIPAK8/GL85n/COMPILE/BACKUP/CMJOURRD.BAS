DEFINT A-Z
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE FUNCTION Round# (ThisAmount#)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB HideCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB OpenMiscCodeFile (NumOfMiscRecs%, MCFile%)
DECLARE SUB PrintJournal (OperRecNumb%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPORT%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (t$, r%, C%, clr%)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION QPValI% (Number$)
DECLARE SUB CursorOff ()

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  
  '*** Citipak Include Files Needed ***
  '$INCLUDE: 'CMFILES.BI'
  '$INCLUDE: 'ubsetup.BI'
  
  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  
  CONST False = 0, True = NOT False

  ShowCursor
  PrintJournal 99
  HideCursor

SUB OpenMiscCodeFile (NumOfMiscRecs, MCFile)
  REDIM MiscCodeRec(1) AS MiscCodeRecType
  MiscCodeRecLen = LEN(MiscCodeRec(1))
  MCFile = FREEFILE
  OPEN "CMMISCCD.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS MCFile LEN = MiscCodeRecLen
  NumOfMiscRecs = LOF(MCFile) \ MiscCodeRecLen
END SUB

SUB PrintJournal (OperRecNumb)
  

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

  LibFile2Scrn "CM.QSL", "MENUBAK", MonoCode%, Attribute%, ErrorCode%
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%

  
  REDIM CMOperRec(1) AS CMOperRecType
  CMOperRecLen = LEN(CMOperRec(1))
  CMFile = FREEFILE
  OPEN "CMOPER.DAT" FOR RANDOM SHARED AS CMFile LEN = CMOperRecLen
  GET CMFile, OperRecNumb, CMOperRec(1)

  Fld(5).Protected = True

  'Form$(5, 0) = "ALL"
  CLOSE CMFile

  FirstTime = True
  Action = 1
  

  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      LSET Form$(3, 0) = DATE$
      LSET Form$(4, 0) = DATE$
      QPrintRC SPACE$(30), 13, 28, 15

      Action = 2
    END IF

    SELECT CASE Frm(1).KeyCode  'Check the user's key presses
      
    CASE F10Key 'SaveButton
      
      BegDate = Date2Num(Form$(3, 0))
      BegDate$ = Form$(3, 0)
      EndDate = Date2Num(Form$(4, 0))
      EndDate$ = Form$(4, 0)
      RecSource$ = UCASE$(LEFT$(Form$(2, 0), 1))

      'OperatorNumber = VAL(Form$(5, 0))
      
      IF LEN(LTRIM$(RTRIM$(Form$(DevFld, 0)))) = 0 THEN
        dev$ = "S"
        LPTNo = 1
      ELSE
        dev$ = LEFT$(Form$(DevFld, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(DevFld, 0)), 1))
      END IF
      
      IF BegDate > EndDate THEN
        Ok2Go = 0
        Frm(1).FldNo = 5
        Action = 1
      ELSE
        Ok2Go = 1
      END IF
      
    CASE ESC    'ESC Key
      EXIT SUB
      
    END SELECT
    
  LOOP UNTIL Ok2Go
  
  CursorOff
  
  IF LEFT$(Form$(1, 0), 2) = "En" THEN
    SortOrder$ = "Entry Order"
  ELSE
    SortOrder$ = "Name Order"
  END IF
  
  ERASE Form$, Fld, Choice$, Frm
  
  'End of Input
  '=====================================================
  'Start Report Processing
  
  ReportFile$ = "CMJOURRD.PRN"  'Report File Name
  Fmt1$ = "#####,#.##"
  Fmt2$ = "###,#.##"
  Fmt3$ = "$$######,#.##"
  Fmt4$ = "$$########,#.##"
  Fmt5$ = "$$#######,#.##"

  Dash$ = STRING$(96, "-")


  FF$ = CHR$(12)
  MaxLines = 53
  LineCnt = 0
  Size = 2500
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
  OPEN "CMTRANS.DAT" FOR RANDOM SHARED AS TrHandle LEN = CMTRRecLen
  TrNumRecs = LOF(TrHandle) \ CMTRRecLen
  
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

    'IF OperatorNumber = 0 OR OperatorNumber = CMTRRec(1).TransOperNum THEN
      IF LineCnt >= MaxLines THEN
        PRINT #RptHandle, FF$
        GOSUB PrintRptHeader
      END IF
      IF CMTRRec(1).TransDate >= BegDate AND CMTRRec(1).TransDate <= EndDate THEN
        PrintTaxFlag = False
        PrintMiscFlag = False
        PrintUtilFlag = False
        PRINT #RptHandle, Num2Date(CMTRRec(1).TransDate);
        IF CMTRRec(1).TransSource = 1 THEN
          PRINT #RptHandle, TAB(12); "Misc.";
        END IF
        
        IF CMTRRec(1).TransSource >= 20 AND CMTRRec(1).TransSource <= 29 THEN
          IF CMTRRec(1).TransSource = 27 THEN
            PRINT #RptHandle, TAB(12); "UT-Dep.";
          ELSE
            PRINT #RptHandle, TAB(12); "Util.";
          END IF
        END IF
        IF CMTRRec(1).TransSource >= 30 AND CMTRRec(1).TransSource <= 39 THEN
          PRINT #RptHandle, TAB(12); "Tax";
        END IF
        IF CMTRRec(1).TransSource >= 40 AND CMTRRec(1).TransSource <= 49 THEN
          PRINT #RptHandle, TAB(12); "Lic.";
        END IF
        IF CMTRRec(1).TransSource >= 50 AND CMTRRec(1).TransSource <= 59 THEN
          PRINT #RptHandle, TAB(12); "Decl.";
          GOSUB SubTotalDecal
        END IF
        
        PRINT #RptHandle, TAB(20); LEFT$(CMTRRec(1).TransName, 18);
        PRINT #RptHandle, TAB(40); USING Fmt1$; CMTRRec(1).TransCash;
        PRINT #RptHandle, TAB(55); USING Fmt1$; CMTRRec(1).TransCheck;
        PRINT #RptHandle, TAB(70); USING Fmt1$; CMTRRec(1).TransAmtOwed;

        IF CMTRRec(1).TransSource = 27 OR CMTRRec(1).TransSource = 51 THEN
          PRINT #RptHandle, TAB(84); USING Fmt1$; Round#((CMTRRec(1).TransCheck + CMTRRec(1).TransCash) - CMTRRec(1).TransAmtOwed)
        ELSE
          PRINT #RptHandle, TAB(84); USING Fmt1$; Round#(CMTRRec(1).TransAmount - CMTRRec(1).TransAmtOwed)
        END IF

        TotalCash# = Round#(TotalCash# + CMTRRec(1).TransCash)
        TotalCheck# = Round#(TotalCheck# + CMTRRec(1).TransCheck)
        TotalAmount# = Round#(TotalAmount# + CMTRRec(1).TransAmtOwed)
        TotalReceipts = TotalReceipts + 1
        LineCnt = LineCnt + 1
        
        IF CMTRRec(1).TransSource = 1 THEN
          'Second Line of Print is Misc Code Breakdown Dist.****************
          FOR MCnt = 1 TO 5
            MiscRevAmt# = Round#(CMTRRec(1).TransRevAmt(MCnt))
            IF MiscRevAmt# > 0 THEN
              ' If There Is an Amount in Misc Rev 1-5 then get code record number from 6-10
              IF CMTRRec(1).TransRevAmt(MCnt + 5) >= 1 THEN
                GET MCFile, CMTRRec(1).TransRevAmt(MCnt + 5), MiscCodeRec(1)
                PRINT #RptHandle, "Code BrkDwn:";
                PRINT #RptHandle, TAB(14); MiscCodeRec(1).MiscCode;
                PRINT #RptHandle, TAB(25); MiscCodeRec(1).Description;
                PRINT #RptHandle, TAB(55); USING Fmt1$; MiscRevAmt#
                PrintMiscFlag = True
                LineCnt = LineCnt + 1
                GOSUB SubTotalMisc
              END IF
            END IF
          NEXT MCnt
          IF PrintMiscFlag THEN
            PRINT #RptHandle, Dash$
            LineCnt = LineCnt + 1
          END IF
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
                PrintUtilFlag = True
                LineCnt = LineCnt + 1
              NEXT RCnt
              GOSUB SubTotalUtil
            END IF
            IF PrintUtilFlag THEN
              PRINT #RptHandle, Dash$
              LineCnt = LineCnt + 1
            END IF
            'End of Utility Print on Second Line *****************************
          ELSE
            'GOSUB GetRevenueSources
            GOSUB SubTotalDep
            'PRINT #RptHandle, Dash$
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
          PrintTaxFlag = True
          LineCnt = LineCnt + 2
          GOSUB SubTotalTax
        END IF
        IF PrintTaxFlag THEN
          PRINT #RptHandle, Dash$
          LineCnt = LineCnt + 1
        END IF
        'End of Tax Print on Second Line *********************************
        
        IF CMTRRec(1).TransSource >= 40 AND CMTRRec(1).TransSource <= 49 THEN
        GOSUB SubTotalBL
        END IF
      'END IF
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
  PRINT #RptHandle, "      Operator: ALL"; TAB(83); "Page #"; page
  PRINT #RptHandle, ""
  PRINT #RptHandle, "   Date"; TAB(12); "Source"; TAB(20); "Description"; TAB(40); "Cash Amount"; TAB(55); "Check Amount"; TAB(70); "Amount Owed"; TAB(84); "Change Given"
  PRINT #RptHandle, STRING$(96, "=")

  LineCnt = 5
RETURN
  
PrintRptEnding:
  PRINT #RptHandle, FF$
  PRINT #RptHandle, "Totals Page:"
  PRINT #RptHandle, ""
  PRINT #RptHandle, "   Total Cash Received: "; USING Fmt3$; TotalCash#
  PRINT #RptHandle, " Total Checks Received: "; USING Fmt3$; TotalCheck#
  PRINT #RptHandle, "                      -----------------"
  PRINT #RptHandle, "  Total Money Received: "; USING Fmt3$; Round#(TotalCash# + TotalCheck#)
  PRINT #RptHandle, "Amount Applied to Acct: "; USING Fmt3$; TotalAmount#
  PRINT #RptHandle, "    Total Change Given: "; USING Fmt3$; Round#((TotalCash# + TotalCheck#) - TotalAmount#)
  PRINT #RptHandle, ""
  PRINT #RptHandle, "   Bank Deposit Amount: "; USING Fmt3$; TotalAmount#
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
    GOSUB PrintBLTotal      ' Not Active Yet!!!
  END IF
  
  IF RecSource$ = "D" OR RecSource$ = "A" THEN
    GOSUB PrintDCTotal      ' Decals
  END IF
  
  PRINT #RptHandle, FF$
RETURN
  
GetReportInformation:
  BegRecNumb = TrNumRecs - 7500 ' Move back 7500 records to begin
  IF BegRecNumb < 1 THEN BegRecNumb = 1         ' Don't Allow Less Than 1
  
  FOR CCnt = BegRecNumb TO TrNumRecs
    GET TrHandle, CCnt, CMTRRec(1)
    TransDate = CMTRRec(1).TransDate
    GoodRecordFlag = False
    
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
      IF SortOrder$ = "Entry Order" THEN
        Array(Count).who = STR$(CCnt)
      ELSE
        Array(Count).who = LEFT$(CMTRRec(1).TransName, 14)
      END IF
      Array(Count).RecNum = CCnt
    END IF
    
  NEXT CCnt
  
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
  
  TownName$ = UBSetUpRec(1).UTILNAME
  RETURN

  REM ** old way

SubTotalUtil:
  FOR uCCnt = 1 TO NumOfRevs
    TotalUtilRevAmt#(uCCnt) = Round#(TotalUtilRevAmt#(uCCnt) + CMTRRec(1).TransRevAmt(uCCnt))
  NEXT uCCnt

RETURN

SubTotalDep:
  FOR dCnt = 1 TO 15
    TotalDepRevAmt#(dCnt) = Round#(TotalDepRevAmt#(dCnt) + CMTRRec(1).TransRevAmt(dCnt))
  NEXT
RETURN

SubTotalDecal:
  TotalDecalRevAmt# = Round#(TotalDecalRevAmt# + CMTRRec(1).TransRevAmt(2))
RETURN

PrintTotalUtil:
  PRINT #RptHandle, Dash$
  PRINT #RptHandle, "GRAND TOTAL Utilities Receipts Recap"
  IF LineCnt >= MaxLines - (NumOfRevs + 5) THEN
    PRINT #RptHandle, FF$
    PRINT #RptHandle, ""
    PRINT #RptHandle, "Totals Page:"
    PRINT #RptHandle, "Total Utilities Receipts Cont'd"
    PRINT #RptHandle, ""
    LineCnt = 4
  END IF
  FOR CCnt = 1 TO NumOfRevs
    PRINT #RptHandle, RevText$(CCnt); TAB(34); USING Fmt4$; TotalUtilRevAmt#(CCnt)
    TotalUtilAmt# = Round#(TotalUtilAmt# + TotalUtilRevAmt#(CCnt))
  NEXT CCnt
  PRINT #RptHandle, "GRAND Total Utility Receipts ... "; USING Fmt4$; TotalUtilAmt#
  PRINT #RptHandle,

  TotalDepAmt# = 0
  FOR CCnt = 1 TO 15
    'PRINT #RptHandle, RevText$(CCnt); TAB(34); USING Fmt4$; TotalDepRevAmt#(CCnt)
    TotalDepAmt# = Round#(TotalDepAmt# + TotalDepRevAmt#(CCnt))
  NEXT CCnt
  PRINT #RptHandle, "GRAND Total Utility Deposits ... "; USING Fmt4$; TotalDepAmt#

RETURN

SubTotalTax:
  TotCurTax# = Round#(TotCurTax# + CMTRRec(1).TransRevAmt(1))
  TotCurInt# = Round#(TotCurInt# + CMTRRec(1).TransRevAmt(2))
  TotCurPen# = Round#(TotCurPen# + CMTRRec(1).TransRevAmt(3))
  TotStrmFee# = Round#(TotStrmFee# + CMTRRec(1).TransRevAmt(4))
  TotPastTax# = Round#(TotPastTax# + CMTRRec(1).TransRevAmt(6))
  TotPastInt# = Round#(TotPastInt# + CMTRRec(1).TransRevAmt(7))
  TotPastPen# = Round#(TotPastPen# + CMTRRec(1).TransRevAmt(8))
  TotPastStrm# = Round#(TotPastStrm# + CMTRRec(1).TransRevAmt(9))
RETURN

PrintTotalTax:
  IF (TotCurTax# + TotCurInt# + TotCurPen# + TotPastTax# + TotPastInt# + TotPastPen# + TotStrmFee# + TotPastStm#) = 0 THEN RETURN
  PRINT #RptHandle, Dash$
  PRINT #RptHandle, "GRAND TOTAL Tax Receipts Recap"
  IF LineCnt >= MaxLines - 6 THEN
    PRINT #RptHandle, FF$
    PRINT #RptHandle, ""
    PRINT #RptHandle, "Totals Page:"
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
  PRINT #RptHandle, "GRAND Total Tax Received ......... "; USING Fmt3$; Round#(TotCurTax# + TotCurInt# + TotCurPen# + TotPastTax# + TotPastInt# + TotPastPen# + TotStrmFee# + TotPastStm#)
  PRINT #RptHandle, ""
  LineCnt = LineCnt + 8
  RETURN

SubTotalBL:
    TotalBLAmt# = Round#(TotalBLAmt# + CMTRRec(1).TransAmtOwed)
RETURN

PrintBLTotal:
  PRINT #RptHandle, Dash$
  PRINT #RptHandle, "GRAND TOTAL Business Licence Receipts Recap"
  PRINT #RptHandle, "GRAND Total Bus. Lic. Receipts .. "; TAB(35); USING Fmt5$; TotalBLAmt#
RETURN

SubTotalMisc:
  IF TotalMiscCnt! = 0 THEN
    TotalMiscCnt! = 1
    TotalMiscRec$(1) = MiscCodeRec(1).MiscCode
    TotalMiscDesc$(1) = MiscCodeRec(1).Description
    TotalMiscAmt#(1) = MiscRevAmt#
    MiscCodeGL$(1) = MiscCodeRec(1).GlAcctNumb
  ELSE
    FOR TCCnt = 1 TO TotalMiscCnt!
      IF MiscCodeRec(1).MiscCode = TotalMiscRec$(TCCnt) THEN
        TotalMiscAmt#(TCCnt) = Round#(TotalMiscAmt#(TCCnt) + MiscRevAmt#): RETURN
      END IF
    NEXT TCCnt
    
    TotalMiscCnt! = TotalMiscCnt! + 1
    TotalMiscRec$(TotalMiscCnt!) = MiscCodeRec(1).MiscCode
    TotalMiscDesc$(TotalMiscCnt!) = MiscCodeRec(1).Description
    TotalMiscAmt#(TotalMiscCnt!) = MiscRevAmt#
    MiscCodeGL$(TotalMiscCnt!) = MiscCodeRec(1).GlAcctNumb
    
  END IF
RETURN


PrintDCTotal:
  PRINT #RptHandle, Dash$
  PRINT #RptHandle, "GRAND TOTAL Decal Receipts Recap"
  IF LineCnt >= MaxLines - (NumOfRevs + 5) THEN
    PRINT #RptHandle, FF$
    PRINT #RptHandle, ""
    PRINT #RptHandle, "Totals Page:"
    PRINT #RptHandle, "Total Decal Receipts Cont'd"
    PRINT #RptHandle, ""
    LineCnt = 4
  END IF
  FOR CCnt = 1 TO NumOfRevs
    PRINT #RptHandle, RevText$(CCnt); TAB(34); USING Fmt4$; TotalUtilRevAmt#(CCnt)
    TotalUtilAmt# = Round#(TotalUtilAmt# + TotalUtilRevAmt#(CCnt))
  NEXT CCnt
  PRINT #RptHandle, "GRAND Total Decal Receipts ..... "; USING Fmt4$; TotalDecalRevAmt#
  PRINT #RptHandle,
RETURN

PrintTotalMisc:
  PRINT #RptHandle, Dash$
  PRINT #RptHandle, "GRAND TOTAL Misc Receipts Recap"
  FOR CCnt = 1 TO TotalMiscCnt!
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      PRINT #RptHandle, ""
      PRINT #RptHandle, "Totals Page:"
      PRINT #RptHandle, "Total Misc Receipts Recap Cont'd"
      PRINT #RptHandle, ""
      LineCnt = 4
    END IF
    PRINT #RptHandle, TotalMiscDesc$(CCnt);
    PRINT #RptHandle, TAB(35); USING Fmt5$; TotalMiscAmt#(CCnt);
    PRINT #RptHandle, TAB(52); "GL# "; MiscCodeGL$(CCnt)
    TotalMisc# = Round#(TotalMisc# + TotalMiscAmt#(CCnt))
    LineCnt = LineCnt + 1
  NEXT CCnt
  PRINT #RptHandle, "GRAND Total Misc Receipts .... "; TAB(35); USING Fmt5$; TotalMisc#
  LineCnt = LineCnt + 1
  RETURN
END SUB

