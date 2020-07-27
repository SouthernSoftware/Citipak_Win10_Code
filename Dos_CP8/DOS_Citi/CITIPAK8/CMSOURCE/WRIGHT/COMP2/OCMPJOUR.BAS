DEFINT A-Z
DECLARE SUB OpenMiscCodeFile (NumOfMiscRecs%, MCFile%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
DECLARE FUNCTION FindAcct% (AcctNum$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE FUNCTION Round# (ThisAmount#)
DECLARE SUB FileView (FileName$, Ky, Action, FVI AS ANY, SEG array)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION MGetKey% (Row%, col%)
DECLARE FUNCTION Monitor% ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (t$, r%, C%, clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PrintLoading ()
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  
  '*** Citipak Include Files Needed ***
  '$INCLUDE: 'CMFILES.BI'
  '$INCLUDE: 'ubtrans.BI'
  ''$INCLUDE: 'gl.bi'
  '$INCLUDE: 'TAX.BI'
  '$INCLUDE: 'ubsetup.BI'

  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE
  
  
  CONST False = 0, True = NOT False

SUB PrintJournal (OperRecNumb)
  
  SHARED User$
  LibName$ = "CM.QSL"
  ScrnName$ = "PRNTJOUR"        'form to edit
  DevFld = 5    'Set to output device field number
  
  SHARED Choice$()
  REDIM RevName$(10), TotalMiscRec$(200), TotalMiscDesc$(200), TotalMiscAmt#(200)
  REDIM TotalUtilRevAmt#(15)
  
  
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
  
  
  REDIM Choice$(0 TO 5, 0 TO 2)
  
  
  Choice$(0, 0) = "1"           'Choices for Output Order
  Choice$(1, 0) = "Entry Order"
  Choice$(2, 0) = "Name"
  
  Choice$(0, 2) = "2"           'Choices for Journal
  Choice$(1, 2) = "ALL"
  Choice$(2, 2) = "Utility Payment"
  Choice$(3, 2) = "Misc Payment"
  Choice$(4, 2) = "Tax Payment"
  Choice$(5, 2) = "License Payment"
  
  Choice$(0, 1) = STR$(DevFld)  'Choices for Output Device
  Choice$(1, 1) = "Screen"
  Choice$(2, 1) = "Prn#1"
  Choice$(3, 1) = "Prn#2"
  
  LibFile2Scrn "CM.QSL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
  Form$(3, 0) = DATE$
  Form$(4, 0) = DATE$
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
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  MaxLines = 53
  LineCnt = 0
  TotDr# = 0
  TotCr# = 0
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
  
  REDIM array(1 TO Size) AS Struct
  REDIM CMTRRec(1) AS CMTransRecType            ' open transaction file
  CMTRRecLen = LEN(CMTRRec(1))
  TrHandle = FREEFILE
  OPEN "CMTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS TrHandle LEN = CMTRRecLen
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
  
  FOR Cnt = 1 TO TrNumRecs
    'GET TrHandle, array(Cnt).RecNum, CMTRRec(1)
    GET TrHandle, Cnt, CMTRRec(1)
    IF LineCnt >= MaxLines THEN
      PRINT #RptHandle, FF$
      GOSUB PrintRptHeader
    END IF
    IF CMTRRec(1).TransDate >= BegDate AND CMTRRec(1).TransDate <= EndDate THEN
    PRINT #RptHandle, Num2Date(CMTRRec(1).TransDate);

    IF CMTRRec(1).TransSource = 1 THEN
      PRINT #RptHandle, TAB(12); "Misc.";
    END IF
    IF CMTRRec(1).TransSource >= 20 AND CMTRRec(1).TransSource <= 29 THEN
      PRINT #RptHandle, TAB(12); "Util.";
    END IF
    IF CMTRRec(1).TransSource >= 30 AND CMTRRec(1).TransSource <= 39 THEN
      PRINT #RptHandle, TAB(12); "Tax";
    END IF
    IF CMTRRec(1).TransSource >= 40 AND CMTRRec(1).TransSource <= 49 THEN
      PRINT #RptHandle, TAB(12); "Lic.";
    END IF
    
    PRINT #RptHandle, TAB(20); LEFT$(CMTRRec(1).TransName, 18);
    PRINT #RptHandle, TAB(40); USING "#####,#.##"; CMTRRec(1).TransCash;
    PRINT #RptHandle, TAB(55); USING "#####,#.##"; CMTRRec(1).TransCheck;
    PRINT #RptHandle, TAB(70); USING "#####,#.##"; CMTRRec(1).TransAmtOwed;
    PRINT #RptHandle, TAB(84); USING "#####,#.##"; (CMTRRec(1).TransAmount - CMTRRec(1).TransAmtOwed)
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
            PRINT #RptHandle, TAB(55); USING "#####,#.##"; MiscRevAmt#
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

      'Second Line of Print is Utility Breakdown Dist. *****************
      GOSUB GetRevenueSources
      IF TotalRevenues > 0 THEN
        PRINT #RptHandle, "Util BrkDwn:";
        FOR RCnt = 1 TO TotalRevenues STEP 2
          PRINT #RptHandle, TAB(15); RevName$(RCnt);
          PRINT #RptHandle, TAB(40); USING "#####,#.##"; CMTRRec(1).TransRevAmt(RCnt);
          PRINT #RptHandle, TAB(55); RevName$(RCnt + 1);
          PRINT #RptHandle, TAB(80); USING "#####,#.##"; CMTRRec(1).TransRevAmt(RCnt + 1)
          PrintUtilFlag = 1
          LineCnt = LineCnt + 1
        NEXT RCnt
        GOSUB SubTotalUtil
      END IF
      IF PrintUtilFlag = 1 THEN PRINT #RptHandle, STRING$(96, "-"): LineCnt = LineCnt + 1
      'End of Utility Print on Second Line *****************************
    END IF
    
    IF CMTRRec(1).TransSource >= 30 AND CMTRRec(1).TransSource <= 39 THEN
      'Second Line of Print is Tax Breakdown Dist.     *****************
      PRINT #RptHandle, "Tax BrkDwn:";
      PRINT #RptHandle, TAB(15); "Tax:"; USING "#####,#.##"; CMTRRec(1).TransRevAmt(1);
      PRINT #RptHandle, TAB(32); "Int.: "; USING "###,#.##"; CMTRRec(1).TransRevAmt(2);
      PRINT #RptHandle, TAB(50); "Pen.: "; USING "###,#.##"; CMTRRec(1).TransRevAmt(3);
      PRINT #RptHandle, TAB(65); "Strm: "; USING "###,#.##"; CMTRRec(1).TransRevAmt(4)
      PRINT #RptHandle, TAB(5); " Past Tax: "; USING "#####,#.##"; CMTRRec(1).TransRevAmt(6);
      PRINT #RptHandle, TAB(32); "Int.: "; USING "###,#.##"; CMTRRec(1).TransRevAmt(7);
      PRINT #RptHandle, TAB(50); "Pen.: "; USING "###,#.##"; CMTRRec(1).TransRevAmt(8);
      PRINT #RptHandle, TAB(65); "Strm: "; USING "###,#.##"; CMTRRec(1).TransRevAmt(9)
      PrintTaxFlag = 1
      LineCnt = LineCnt + 2
      GOSUB SubTotalTax
    END IF
    IF PrintTaxFlag = 1 THEN PRINT #RptHandle, STRING$(96, "-"): LineCnt = LineCnt + 1
    'End of Tax Print on Second Line *********************************
    
    IF CMTRRec(1).TransSource >= 40 AND CMTRRec(1).TransSource <= 49 THEN
      'Second Line of Print is License Breakdown Dist. *****************
      'End of License Print on Second Line *****************************
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
  ERASE TotalUtilRevAmt#, Frm, Form$, Fld
  ERASE array, CMTRRec
  
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
  PRINT #RptHandle, "Totals Page for Operator # "; OperRec(1).OperatorNumber
  PRINT #RptHandle, ""
  PRINT #RptHandle, "   Total Cash Received: "; USING "$$######,#.##"; TotalCash#
  PRINT #RptHandle, " Total Checks Received: "; USING "$$######,#.##"; TotalCheck#
  PRINT #RptHandle, "                      -----------------"
  PRINT #RptHandle, "  Total Money Received: "; USING "$$######,#.##"; TotalCash# + TotalCheck#
  PRINT #RptHandle, "Amount Applied to Acct: "; USING "$$######,#.##"; TotalAmount#
  PRINT #RptHandle, "    Total Change Given: "; USING "$$######,#.##"; (TotalCash# + TotalCheck#) - TotalAmount#
  PRINT #RptHandle, ""
  PRINT #RptHandle, "   Bank Deposit Amount: "; USING "$$######,#.##"; TotalAmount#
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

  'IF RecSource$ = "L" OR RecSource$ = "A" THEN
  'GOSUB PrintTotalLicense      ' Not Active Yet!!!
  'END IF
  
  
  PRINT #RptHandle, FF$
  RETURN
  
  
GetReportInformation:
  BegRecNumb = TrNumRecs - 7500 ' Move back 7500 records to begin
  IF BegRecNumb < 1 THEN BegRecNumb = 1         ' Don't Allow Less Than 1
  
  FOR Cnt! = BegRecNumb TO TrNumRecs
    GET TrHandle, Cnt!, CMTRRec(1)
    TransDate = CMTRRec(1).TransDate
    GoodRecordFlag = False
    
    IF CMTRRec(1).TransOperNum = OperRecNumb AND (TransDate >= BegDate AND TransDate <= EndDate) THEN

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
      END IF

      IF GoodRecordFlag THEN
        Count = Count + 1
        IF SortOrder$ = "Entry Order" THEN
          array(Count).who = STR$(Cnt!)
        ELSE
          array(Count).who = LEFT$(CMTRRec(1).TransName, 14)
        END IF
        array(Count).RecNum = Cnt!
      END IF

  NEXT Cnt!
  
  SortT array(Start), Count, Dir, SSize, MOff, MSize
  RETURN
  
GetRevenueSources:
  RHandle = FREEFILE
  OPEN "wbinfo.dat" FOR RANDOM ACCESS READ WRITE SHARED AS RHandle LEN = 512
  FIELD RHandle, 30 AS dispname$, 20 AS DEFCITY$, 2 AS defst$, 5 AS defzip$
  FIELD RHandle, 57 AS du1$, 2 AS REVTrack$
  FIELD RHandle, 59 AS du2$, 22 AS Rev1Name$, 1 AS rev1code$
  FIELD RHandle, 82 AS du3$, 22 AS Rev2Name$, 1 AS rev2code$
  FIELD RHandle, 105 AS du4$, 22 AS Rev3Name$, 1 AS rev3code$
  FIELD RHandle, 128 AS du5$, 22 AS Rev4Name$, 1 AS rev4code$
  FIELD RHandle, 151 AS du6$, 22 AS Rev5Name$, 1 AS rev5code$
  FIELD RHandle, 174 AS du7$, 22 AS Rev6Name$, 1 AS rev6code$
  FIELD RHandle, 197 AS du8$, 22 AS Rev7Name$, 1 AS rev7code$
  FIELD RHandle, 220 AS du9$, 22 AS Rev8Name$, 1 AS rev8code$
  FIELD RHandle, 243 AS du10$, 269 AS blank$
  GET RHandle, 1
  TotalRevenues = CVI(REVTrack$)
  IF TotalRevenues = 0 THEN CLOSE RHandle: RETURN
  RevName$(1) = QPTrim$(Rev1Name$)
  RevName$(2) = QPTrim$(Rev2Name$)
  RevName$(3) = QPTrim$(Rev3Name$)
  RevName$(4) = QPTrim$(Rev4Name$)
  RevName$(5) = QPTrim$(Rev5Name$)
  RevName$(6) = QPTrim$(Rev6Name$)
  RevName$(7) = QPTrim$(Rev7Name$)
  RevName$(8) = QPTrim$(Rev8Name$)
  CLOSE RHandle
  RETURN
  
SubTotalUtil:
  FOR uCnt! = 1 TO TotalRevenues
    TotalUtilRevAmt#(uCnt!) = TotalUtilRevAmt#(uCnt!) + CMTRRec(1).TransRevAmt(uCnt!)
  NEXT uCnt!
  RETURN
PrintTotalUtil:
  PRINT #RptHandle, STRING$(96, "-")
  PRINT #RptHandle, "GRAND TOTAL Utilities Receipts Recap"
  IF LineCnt >= MaxLines - (TotalRevenues + 5) THEN
    PRINT #RptHandle, FF$
    PRINT #RptHandle, ""
    PRINT #RptHandle, "Totals Page for Operator # "; OperRec(1).OperatorNumber
    PRINT #RptHandle, "Total Utilities Receipts Cont'd"
    PRINT #RptHandle, ""
    LineCnt = 4
  END IF
  FOR Cnt! = 1 TO TotalRevenues
    PRINT #RptHandle, RevName$(Cnt!); TAB(34); USING "$$########,#.##"; TotalUtilRevAmt#(Cnt!)
    TotalUtilAmt# = TotalUtilAmt# + TotalUtilRevAmt#(Cnt!)
  NEXT Cnt!
  PRINT #RptHandle, "GRAND Total Utility Receipts ... "; USING "$$########,#.##"; TotalUtilAmt#
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
  PRINT #RptHandle, "Total Current Taxes Received ..... "; USING "$$######,#.##"; TotCurTax#
  PRINT #RptHandle, "Total Current Interest Received .. "; USING "$$######,#.##"; TotCurInt#
  PRINT #RptHandle, "Total Current Penalty Received ... "; USING "$$######,#.##"; TotCurPen#
  PRINT #RptHandle, "Total Storm Fee Received ......... "; USING "$$######,#.##"; TotStrmFee#
  PRINT #RptHandle, "Total Past Taxes Received ........ "; USING "$$######,#.##"; TotPastTax#
  PRINT #RptHandle, "Total Past Interest Received ..... "; USING "$$######,#.##"; TotPastInt#
  PRINT #RptHandle, "Total Past Penalty Received ...... "; USING "$$######,#.##"; TotPastPen#
  PRINT #RptHandle, "Total Past Storm Fee Received .... "; USING "$$######,#.##"; TotPastStrm#
  PRINT #RptHandle, "GRAND Total Tax Received ......... "; USING "$$######,#.##"; (TotCurTax# + TotCurInt# + TotCurPen# + TotPastTax# + TotPastInt# + TotPastPen# + TotStrmFee# + TotPastStm#)
  PRINT #RptHandle, ""
  LineCnt = LineCnt + 8
  RETURN
  
SubTotalMisc:
  IF TotalMiscCnt! = 0 THEN
    TotalMiscCnt! = 1
    TotalMiscRec$(1) = MiscCodeRec(1).MiscCode
    TotalMiscDesc$(1) = MiscCodeRec(1).Description
    TotalMiscAmt#(1) = MiscRevAmt#
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
    PRINT #RptHandle, TAB(35); USING "$$#######,#.##"; TotalMiscAmt#(Cnt!)
    TotalMisc# = TotalMisc# + TotalMiscAmt#(Cnt!)
    LineCnt = LineCnt + 1
  NEXT Cnt!
  PRINT #RptHandle, "GRAND Total Misc Receipts .... "; TAB(35); USING "$$#######,#.##"; TotalMisc#
  LineCnt = LineCnt + 1
  RETURN
END SUB

SUB PrintReceipt

  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUpRec(1))
  FGetAH "UBSETUP.DAT", UBSetUpRec(1), UBSetupLen, 1            'load it
  TownName$ = UBSetUpRec(1).UTILNAME

  RMFile = FREEFILE
  REDIM RMRec(1) AS RMReceiptRecType
  RMRecLen = LEN(RMRec(1))
  OPEN "CMRECPT.DAT" FOR RANDOM AS RMFile LEN = RMRecLen
  LOCK #RMFile, 1
  GET RMFile, 1, RMRec(1)
  
  LPRINT
  LPRINT TAB(12); TownName$
  LPRINT TAB(12); "  MISC PAYMENT"
  LPRINT TAB(13); "Date: "; RMRec(1).RecDate
  LPRINT
  LPRINT "CUSTOMER NAME & DESC. OF PAYMENT"
  LPRINT RMRec(1).RecName
  LPRINT RMRec(1).RecAddress
  LPRINT RMRec(1).RecDesc
  LPRINT
  LPRINT "Total Owed: "; USING "$$####,#.##"; RMRec(1).RecAmtOwed
  LPRINT "Total Paid: "; USING "$$####,#.##"; RMRec(1).RecCashAmt + RMRec(1).RecCheckAmt
  LPRINT "Change Due: "; USING "$$####,#.##"; RMRec(1).RecChangeDue
  LPRINT
  LPRINT
  LPRINT "Operator: "; RMRec(1).RecOperator
  LPRINT "Receipt#: "; USING "######"; RMRec(1).RecptNumber
  LPRINT
  LPRINT TAB(7); "T H A N K   Y O U !"
  LPRINT
  LPRINT
  LPRINT
  LPRINT
  LPRINT
  LPRINT
  LPRINT CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)
  CLOSE RMFile
  EXIT SUB
  
END SUB

