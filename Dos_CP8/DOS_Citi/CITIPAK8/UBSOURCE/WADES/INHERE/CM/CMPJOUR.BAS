  DEFINT A-Z
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
DECLARE SUB QPrintRC (t$, r%, C%, clr%)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB WaitForAction ()
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  
  '*** Citipak Include Files Needed ***
  '$INCLUDE: 'CMFILES.BI'
  ''$INCLUDE: 'ubtrans.BI'
  ''$INCLUDE: 'gl.bi'
  ''$INCLUDE: 'TAX.BI'
  '$INCLUDE: 'ubsetup.BI'
  
  TYPE Struct
    who AS STRING * 12
    RecNum AS LONG
  END TYPE
  
  
  CONST False = 0, True = NOT False
  
  'PrintJournal 1
  ' PrintJournal 99

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
      
    CASE F10Key 'SaveButton
      
      BegDate = Date2Num(Form$(3, 0))
      BegDate$ = Form$(3, 0)
      EndDate = Date2Num(Form$(4, 0))
      EndDate$ = Form$(4, 0)
      RecSource$ = UCASE$(LEFT$(Form$(2, 0), 1))
      
      OperatorNumber = VAL(Form$(5, 0))
      
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

  'SumLine$ = STRING$(13, "-")   'column summary line
  'DivLine$ = STRING$(77, "-")   'dashed line
  'DivLine2$ = STRING$(77, "=")  'Double Line

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
        TrType$ = ""
        SELECT CASE CMTRRec(1).TransSource
          CASE 1
            TrType$ = "Misc."
          CASE 27
            TrType$ = "UT-Dep."
          CASE 24
             TrType$ = "Util."
          CASE IS >= 30 <= 39
            TrType$ = "Tax"
          CASE IS >= 40 <= 49
            TrType$ = "Lic."
        END SELECT

        PRINT #RptHandle, Num2Date(CMTRRec(1).TransDate); TAB(12); TrType$; TAB(20); LEFT$(CMTRRec(1).TransName, 18);
        PRINT #RptHandle, TAB(40); USING Fmt1$; CMTRRec(1).TransCash; TAB(55); CMTRRec(1).TransCheck; TAB(70); CMTRRec(1).TransAmtOwed;

        IF CMTRRec(1).TransSource = 27 THEN
          PRINT #RptHandle, TAB(84); USING Fmt1$; ((CMTRRec(1).TransCheck + CMTRRec(1).TransCash) - CMTRRec(1).TransAmtOwed)
        ELSE
          PRINT #RptHandle, TAB(84); USING Fmt1$; (CMTRRec(1).TransAmount - CMTRRec(1).TransAmtOwed)
        END IF
        
        TotalCash# = TotalCash# + CMTRRec(1).TransCash
        TotalCash# = Round(TotalCash#)
        TotalCheck# = TotalCheck# + CMTRRec(1).TransCheck
        TotalCheck# = Round(TotalCheck#)
        TotalAmount# = TotalAmount# + CMTRRec(1).TransAmtOwed
        TotalAmount# = Round(TotalAmount#)
        TotalReceipts = TotalReceipts + 1
        LineCnt = LineCnt + 1
        
        IF CMTRRec(1).TransSource = 1 THEN
          'Second Line of Print is Misc Code Breakdown Dist.****************
          PrintMiscFlag = 0
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
        IF PrintTaxFlag = 1 THEN PRINT #RptHandle, STRING$(96, "-"): LineCnt = LineCnt + 1
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
  
  KILL ReportFile$
  
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
  PRINT #RptHandle, "Amount Applied to Acct: "; USING Fmt3$; TotalAmount#
  PRINT #RptHandle, "    Total Change Given: "; USING Fmt3$; (TotalCash# + TotalCheck#) - TotalAmount#
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
    GOSUB PrintBLTotal          ' Not Active Yet!!!
  END IF
  
  IF RecSource$ = "D" OR RecSource$ = "A" THEN
    GOSUB PrintDCTotal
  END IF
  
  PRINT #RptHandle, FF$
RETURN
  
GetReportInformation:
  BegRecNumb = TrNumRecs& - 7500                ' Move back 7500 records to begin
  IF BegRecNumb < 1 THEN BegRecNumb = 1         ' Don't Allow Less Than 1
  IF OperatorNumber = 0 THEN BegRecNumb = 1
  
  FOR Cnt! = BegRecNumb TO TrNumRecs&
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
  
  TownName$ = UBSetUpRec(1).UTILNAME
RETURN
  
  REM ** old way
  
SubTotalUtil:
  FOR uCnt! = 1 TO NumOfRevs
    TotalUtilRevAmt#(uCnt!) = TotalUtilRevAmt#(uCnt!) + CMTRRec(1).TransRevAmt(uCnt!)
  NEXT uCnt!
  
RETURN
  
SubTotalDep:
  FOR dCnt = 1 TO 15
    TotalDepRevAmt#(dCnt) = Round#(TotalDepRevAmt#(dCnt) + CMTRRec(1).TransRevAmt(dCnt))
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
  TotalBLAmt# = TotalBLAmt# + CMTRRec(1).TransAmtOwed
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
  TownName$ = UBSetUpRec(1).UTILNAME
  
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
  PRINT #RctFile, TAB(12); TownName$
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
    PRINT #RctFile, CHR$(27); "="; CHR$(2)           'only to display
    LSET Work$ = "THANK YOU!"
    PRINT #RctFile, Work$
    LSET Work$ = "Change:" + FUsing$(STR$(RMRec(1).RecChangeDue), "$$####.##")
    PRINT #RctFile, Work$; CHR$(10)
    PRINT #RctFile, CHR$(27); "="; CHR$(1)           'only to printer
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
  
  '27 "p" + 0   drawer 1
  '27 "p" + 1   drawer 2
  
  CLOSE #RctFile
  
  PrintRptFile Header$, "Recpt.PRN", GetDEFPort%, RetCode%, 5
  
  'SHELL "Type Recpt.PRN > lpt2:"
  
  '  EXIT SUB
  
END SUB

