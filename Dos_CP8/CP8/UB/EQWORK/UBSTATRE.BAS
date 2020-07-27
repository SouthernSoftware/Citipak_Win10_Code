  DEFINT A-Z
DECLARE FUNCTION QPSTRL$ (LONGNO&)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB CustbyRate ()
DECLARE SUB RateAnalysisMenu ()
DECLARE SUB AnalysisByRATE ()
DECLARE SUB AnalysisByREVENUE ()
DECLARE SUB ConsumpUnitStep ()
DECLARE SUB UsageByRateCode ()
DECLARE SUB UsageByZoneCode ()
DECLARE SUB UsageByBook ()
DECLARE SUB RateAnalysis ()
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB BlockClear ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PressButton (BYVAL Key2Stuff%, BYVAL MouseActionRow%, BYVAL LeftKeyCol%, BYVAL RightKeyCol%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE SUB WaitForAction ()
DECLARE SUB KillFile (FileName$)
  '$INCLUDE: 'DefCnf.BI'
  
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Date2Num% (AnyDate$)
DECLARE FUNCTION Num2Date$ (AnyDate%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB QPrintRC (PctC$, Row, Col, Kolor)
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
  'DECLARE SUB VendorMenu ()
DECLARE SUB HideCursor ()
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ClearScrn ()
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'PageInfo.BI'                  'Form Page info
  
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'Newcust.bi'
  ''$INCLUDE: 'ubLoc.bi'
  '$INCLUDE: 'ubrate.bi'
  '$INCLUDE: 'ubTrans.bi'
  ''$INCLUDE: 'ubmeter.BI'
  
  CONST False = 0, True = NOT False

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 6)
  
  MChoice$(1) = "Customer Usage by Rate Code"
  MChoice$(2) = "Customer Usage by Book (Route)"
  MChoice$(3) = "Customer Count by Rate Code"
  MChoice$(4) = "Exit to OS"
  'MChoice$(4) = "Rate Analysis"
  'MChoice$(5) = "Exit to OS"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 1
  Help$ = "Statistical Reports Menu"
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 2, Col, MaxLen + 3, "Statistical Reports Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      ConsumpUnitStep
    CASE 2
      UsageByBook
    CASE 3
      CustbyRate
      'CASE 4
      '  RateAnalysisMenu
    CASE 4
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
  
  RUN "ubmenu"

SUB AnalysisByRATE
  
  SHARED Choice$()
  
  REDIM TempScrn(0)
  REDIM RevenueName$(15)
  REDIM RateCode$(100), RateDesc$(100), RateTrans(100), RateUse#(100), RateRev#(100)
  
  LibName$ = "UB"
  ScrnName$ = "ANRATRPT"
  MActionRow = 17
  
  Dash80$ = STRING$(80, "-")
  
  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  Action = 1
  Frm(1).StayOnField = True
  
  NumOfRates = GetNumRateRecs%
  REDIM UBRateTbls(1 TO NumOfRates) AS UBRateTblRecType
  REDIM Choice$(0 TO NumOfRates + 1, 0)
  
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  Choice$(1, 0) = "All-Rate Codes"
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRates
    GET RateFile, Cnt, UBRateTbls(Cnt)
    Choice$(Cnt + 1, 0) = UCASE$(UBRateTbls(Cnt).RateCode + " " + UBRateTbls(Cnt).RateDesc)
  NEXT
  CLOSE
  
  Choice$(0, 0) = "3"
  
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB AnRateCheckReqFields
      IF ReqFldsOK THEN
        ExitFlag = True
      ELSE
        Action = 2
      END IF
    CASE EscKey
      AbortFlag = True
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE MActionRow           'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 28 TO 41           '--Cancel button
          PressButton 27, MActionRow, 28, 41
        CASE 42 TO 55           '--F10 button
          PressButton F10Key, MActionRow, 42, 55
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
  LOOP UNTIL ExitFlag
  
  
  
  
  
  
  
  
  IF AbortFlag THEN GOTO ExitAnRate:
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetUp(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetUp(1))
  
  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))
  
  
  UBRpt = FREEFILE
  OPEN "UBRATEAN.RPT" FOR OUTPUT AS UBRpt
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBTrans LEN = UBTransRecLen
  
  NumOfRecs! = LOF(UBTrans) / UBTransRecLen
  
  UBSetUp = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBSetUp LEN = UBSetUpRecLen
  IF LOF(UBSetUp) / UBSetUpRecLen = 0 THEN
    TownName$ = "Undefined"
    ' Set Revenue Names to Nothing
    FOR RCnt = 1 TO 15
      RevenueName$(RCnt) = "Not Set"
    NEXT RCnt
  ELSE
    GET UBSetUp, 1, UBSetUp(1)
    FOR RCnt = 1 TO 15
      RevenueName$(RCnt) = QPTrim$(UBSetUp(1).Revenues(RCnt).REVNAME)
    NEXT RCnt
    RCnt = 1
    DO WHILE RCnt <= 15
      IF RevenueName$(RCnt) = "" THEN
        MaxRevenue = RCnt - 1
        EXIT DO
      END IF
      RCnt = RCnt + 1
    LOOP
    
    TownName$ = UBSetUp(1).UTILNAME
    TownLen = LEN(RTRIM$(TownName$))
    TabStop = 40 - (TownLen / 2)
    IF TabStop < 1 THEN TabStop = 1
  END IF
  
  CLOSE UBSetUp
  
  GOSUB DoAnalByRateHeader
  
  
  
  
  REM delete and finish ******
  BlockClear
  ShowProcessingScrn "Analysis by Rate"
  
  
  CLOSE
  SLEEP 5
  
  SaveScrn TempScrn()
  DisplayUBScrn "NOTENGH"
  WaitForAction
  RestScrn TempScrn()
  Frm(1).FldNo = 1
  GOTO ExitAnRate
  
  REM *** Take Out to Here
  GOSUB GetReportData
  
  
  FOR Cnt = 1 TO RateCnt
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB DoAnalByRateHeader
    END IF
    
    '*************************************
    
    '   Main body of Printing goes here
    
    '   PRIN TAB(60); "Tot Rev"; TAB(70); "Avg Rev"
    
    RateDesc$(1) = RateCode$(RCnt)
    RateTrans(1) = 1
    RateUse#(1) = CalcUsage#
    RateRev#(1) = RevenueTotal#
    
    PRINT #UBRpt, RateDesc$(Cnt);
    ' PRINT #UBRpt,
    PRINT #UBRpt, USING "#####"; TAB(30); RateTrans(Cnt);
    PRINT #UBRpt, TAB(40); USING "#######,#"; RateUse#(Cnt);
    IF RateTrans(Cnt) > 0 THEN
      PRINT #UBRpt, TAB(50); USING "#####,#.##"; RateUse#(Cnt) / RateTrans(Cnt);
    END IF
    PRINT #UBRpt, TAB(60); USING "#####,#.##"; RateRev#(Cnt);
    IF RateTrans(Cnt) > 0 THEN
      PRINT #UBRpt, TAB(70); USING "#####,#.##"; RateRev#(Cnt) / RateTrans(Cnt)
    ELSE
      PRINT #UBRpt, ""
    END IF
    LineCnt = LineCnt + 1
    
    '*************************************
    
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    
    ShowPctComp Cnt, RateCnt
    
    
  NEXT
  
  GOSUB DoAnalByRateFooter:
  
  
  CLOSE
  
  
  
  IF NOT AbortFlag THEN
    PrintRptFile "Rate Analysis by Rate", "UBRATEAN.RPT", 1, RetCode, EntryPoint
  END IF
  
  KillFile "UBRATEAN.RPT"
  GOTO ExitAnRate:
  
GetReportData:
  RatetoFind$ = Form$(3, 0)
  IF LEFT$(RatetoFind$, 3) = "All" THEN
    RatetoFind$ = "ALL"
  ELSE
    RatetoFind$ = RTRIM$(RatetoFind$)
  END IF
  
  FOR Cnt! = 1 TO NumOfRecs!
    ShowPctCompL CLNG(Cnt!), CLNG(NumOfRecs!)
    
    GET UBTrans, Cnt!, UBTransRec(1)
    IF UBTransRec(1).TransDate < StartDate THEN
      BadCount = BadCount + 1
      IF BadCount > 500 THEN
        RETURN
      END IF
    END IF
    IF (UBTransRec(1).TransDate >= StartDate AND UBTransRec(1).TransDate <= Enddate) THEN
      IF UBTransRec(1).TransType = 1 THEN
        'rem rate code and check against what was asked for
        'GET UBLoca, UBTransRec(1).CustLocation, UBLocaRec(1)
        'find all rate codes
        FOR RCnt = 1 TO MaxRevenue
          'RateCode$(RCnt) = UBLocaRec(1).Serv(RCnt).RateCode
        NEXT RCnt
        
        'if Rate Codes = All then Store All of Them for printing else Findit
        FOR RCnt = 1 TO MaxRevenue
          IF RatetoFind$ = "ALL" THEN
            GOSUB CalcUse
            GOSUB GetRev
            IF RateCnt = 0 THEN
              RateCnt = 1
              RateDesc$(1) = RateCode$(RCnt)
              RateTrans(1) = 1
              RateUse#(1) = CalcUsage#
              RateRev#(1) = RevenueTotal#
            ELSE
              FOR LL = 1 TO RateCnt
                IF RateDesc$(LL) = RateCode$(RCnt) THEN
                  RateTrans(LL) = RateTrans(LL) + 1
                  RateUse#(LL) = RateUse#(LL) + CalcUsage#
                  RateRev#(LL) = RateRev#(LL) + RevenueTotal#
                  NoProcess = 1
                END IF
              NEXT LL
              IF NoProcess = 0 THEN
                RateCnt = RateCnt + 1
                RateDesc$(RateCnt) = RateCode$(RCnt)
                RateTrans(RateCnt) = 1
                RateUse#(RateCnt) = CalcUsage#
                RateRev#(RateCnt) = RevenueTotal#
              END IF
            END IF
            EXIT FOR
            
          ELSE
            
          END IF
          
          
          
          
          
          
        NEXT RCnt
        
        
        
        
      END IF
    END IF
    
  NEXT Cnt!
  RETURN
  
CalcUse:
  ' late put in check by revenue
  CalcUsage# = UBTransRec(1).CurRead(1) - UBTransRec(1).PrevRead(1)
  RETURN
  
GetRev:
  ' this section needs work also
  RevenueTotal# = RevenueTotal# + UBTransRec(1).TransAmt
  RETURN
  
DoAnalByRateHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop + 4); TownName$
  PRINT #UBRpt, TAB(31); "Rate Code Analysis by Rate"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$;
  PRINT #UBRpt,
  PRINT #UBRpt, "Code"; TAB(8); "Description"; TAB(30); "# Trans"; TAB(40); "Consump"; TAB(50); "Avg Con"; TAB(60); "Tot Rev"; TAB(70); "Avg Rev"
  PRINT #UBRpt, Dash80$
  LineCnt = 6
  RETURN
DoAnalByRateFooter:
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Totals:"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, CHR$(12);
  RETURN
  
  
  
  
AnRateCheckReqFields:
  BadDate = False
  StartDate = Date2Num(Form$(1, 0))
  Enddate = Date2Num(Form$(2, 0))
  RateCode$ = QPTrim$(Form$(3, 0))
  
  IF StartDate < 0 OR Enddate < 0 THEN
    Frm(1).FldNo = 1
    BadDate = True
  ELSEIF StartDate > Enddate THEN
    BadDate = True
  END IF
  
  IF BadDate OR LEN(RateCode$) = 0 THEN
    SaveScrn TempScrn()
    DisplayUBScrn "BADRDATE"
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 1
    Action = 2
  ELSE
    ReqFldsOK = True
  END IF
  
  IF Choice$(1, 0) <> RateCode$ THEN
    FOR Cnt = 2 TO NumOfRates
      IF INSTR(Choice$(Cnt, 0), RateCode$) THEN
        WhatRate = Cnt - 1
        EXIT FOR
      END IF
    NEXT
  ELSE
    WhatRate = 0
  END IF
  
  RETURN
  
ExitAnRate:
  
END SUB

SUB AnalysisByREVENUE
  REDIM TempScrn(0)
  
  LibName$ = "UB"
  ScrnName$ = "ANREVRPT"
  MActionRow = 17
  
  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  Action = 1
  Frm(1).StayOnField = True
  
  '--Set screen number to one and display screen
  SHARED Choice$()
  
  NumOfRevs = MaxRevsCnt
  REDIM Choice$(0 TO MaxRevsCnt + 1, 0)
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  Choice$(1, 0) = "All-Revenues"
  
  'REDIM Choice$(0 TO MaxRevsCnt, 0)
  FOR Cnt = 1 TO MaxRevsCnt     'find last active revenue
    IF UBSetUpRec(1).Revenues(Cnt).USERATE = "Y" THEN
      TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(Cnt).REVNAME)
      IF LEN(TempRev$) = 0 THEN
        NumOfRevs = Cnt - 1     'set actual number of revenues
        EXIT FOR
      ELSE
        Choice$(Cnt + 1, 0) = UCASE$(TempRev$)
        'build revenue description line
      END IF
    END IF
  NEXT
  
  'REDIM PRESERVE Choice$(0 TO NumOfRevs, 0)
  
  Choice$(0, 0) = "3"
  
  
  BlockClear
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB AnReveCheckReqFields
      IF ReqFldsOK THEN
        ExitFlag = True
      ELSE
        Action = 2
      END IF
    CASE EscKey
      AbortFlag = True
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE MActionRow           'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 28 TO 41           '--Cancel button
          PressButton 27, MActionRow, 28, 41
        CASE 42 TO 55           '--F10 button
          PressButton F10Key, MActionRow, 42, 55
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN GOTO ExitAnReve:
  BlockClear
  ShowProcessingScrn "Analysis by Revenue"
  
  
  CLOSE
  SLEEP 5
  
  SaveScrn TempScrn()
  DisplayUBScrn "NOTENGH"
  WaitForAction
  RestScrn TempScrn()
  Frm(1).FldNo = 1
  EXIT SUB
  
  
  
  
  
  GOTO ExitAnReve:
  
AnReveCheckReqFields:
  BadDate = False
  StartDate = Date2Num(Form$(1, 0))
  Enddate = Date2Num(Form$(2, 0))
  Revenue$ = QPTrim$(Form$(3, 0))
  
  IF StartDate < 0 OR Enddate < 0 THEN
    Frm(1).FldNo = 1
    BadDate = True
  ELSEIF StartDate > Enddate THEN
    BadDate = True
  END IF
  
  
  IF BadDate OR LEN(Revenue$) = 0 THEN
    SaveScrn TempScrn()
    DisplayUBScrn "BADRINFO"
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 1
    Action = 2
  ELSE
    ReqFldsOK = True
  END IF
  
  IF Choice$(1, 0) <> Revenue$ THEN
    FOR Cnt = 2 TO NumOfRevs
      IF INSTR(Choice$(Cnt, 0), Revenue$) THEN
        WhatRevenue = Cnt - 1
        EXIT FOR
      END IF
    NEXT
  ELSE
    WhatRevenue = 0
  END IF
  
  
  RETURN
  
ExitAnReve:
  
END SUB

SUB ConsumpUnitStep
  
  SHARED Choice$()
  REDIM TblBreak&(10), TotalConsp#(10), TotalCust(10)
  
  REDIM TempScrn(0)
  MaxLines = 56
  Dash80$ = STRING$(80, "-")
  
  LibName$ = "UB"
  
  
  ScrnName$ = "CONUSTEP"
  'Add
  MActionRow = 18
  
  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  Action = 1
  Frm(1).StayOnField = True
  
  NumOfRates = GetNumRateRecs%
  REDIM UBRateTbls(1 TO NumOfRates) AS UBRateTblRecType
  REDIM Choice$(0 TO NumOfRates + 10, 1)
  
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRates
    GET RateFile, Cnt, UBRateTbls(Cnt)
    Choice$(Cnt, 0) = UCASE$(UBRateTbls(Cnt).RateCode + " " + UBRateTbls(Cnt).RateDesc) + " " + STR$(Cnt)
  NEXT
  CLOSE
  
  Choice$(0, 0) = "3"
  
  Choice$(0, 1) = "4"
  Choice$(1, 1) = "Water Only"
  Choice$(2, 1) = "Sewer Only"
  Choice$(3, 1) = "Both Water/Sewer"
  Choice$(4, 1) = "Electric"
  Choice$(5, 1) = "Demand Elec"
  Choice$(6, 1) = "Gas"
  Choice$(7, 1) = "Touch Read"

  
  FirstTime = True
  
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    'IF FirstTime THEN
    '  FirstTime = False
    '  LSET Form$(1, 0) = "01-01-1998"
    '  LSET Form$(2, 0) = "04-30-1998"
    '  LSET Form$(3, 0) = "W0"
    '  LSET Form$(4, 0) = Choice$(3, 1)
    '  Action = 1
    'END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      
      GOSUB ConsCheckReqFields:
      IF ReqFldsOK THEN
        ExitFlag = True
      ELSE
        Action = 2
      END IF
    CASE EscKey
      AbortFlag = True
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE MActionRow           'Look for the f10 or esc button
        
        SELECT CASE Frm(1).MCol
        CASE 28 TO 41           '--Cancel button
          PressButton 27, MActionRow, 28, 41
        CASE 42 TO 55           '--F10 button
          PressButton F10Key, MActionRow, 42, 55
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
  LOOP UNTIL ExitFlag
  IF AbortFlag THEN GOTO ExitConsStep
  
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Lnt = 1 TO NumOfRates
    WhatRate$ = QPTrim$(LEFT$(Form$(3, 0), 4))
    ThisRate$ = QPTrim$(LEFT$(Choice$(Lnt, 0), 4))

    IF ThisRate$ = WhatRate$ THEN
      GET RateFile, Lnt, UBRateTbls(1)
      RateCode$ = QPTrim$(UBRateTbls(1).RateCode)

      FOR Tnt = 1 TO 10
        IF UBRateTbls(Lnt).TblBreaks(Tnt).Units >= 0 THEN
          IF UBRateTbls(Lnt).TblBreaks(Tnt).Units > 0 THEN
            Greater = True
          END IF
          IF (UBRateTbls(Lnt).TblBreaks(Tnt).Units = 0) AND (Greater = True) THEN
            MaxStep = Tnt
            TblBreak&(Tnt) = 99999999
            EXIT FOR
          END IF

          TblBreak&(Tnt) = UBRateTbls(Lnt).TblBreaks(Tnt).Units

          MaxStep = Tnt
        ELSE
          EXIT FOR
        END IF
      NEXT Tnt
    END IF
  NEXT Lnt
  CLOSE
  
  Mtype$ = RTRIM$(Form$(4, 0))
  IF Mtype$ = "Water Only" THEN Mtype = 1: Mtype1 = 3
  IF Mtype$ = "Sewer Only" THEN Mtype = 2: Mtype1 = 3
  IF Mtype$ = "Both Water/Sewe" THEN Mtype = 3: Mtype1 = 1
  IF Mtype$ = "Electric" THEN Mtype = 4: Mtype1 = 9
  IF Mtype$ = "Demand Elec" THEN Mtype = 5: Mtype1 = 9
  IF Mtype$ = "Gas" THEN Mtype = 6: Mtype1 = 9
  IF Mtype$ = "Touch Read" THEN Mtype = 7: Mtype1 = 7
  IF Mtype = 0 THEN
    Mtype = 7
  END IF
  
  BegDate = Date2Num%(Form$(1, 0))
  Enddate = Date2Num%(Form$(2, 0))
  
  IndexName$ = BookIndexFile
  NumOfRecs = FileSize(IndexName$) \ 4
  REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
  FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
  
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetUpRec(1))
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTrans LEN = UBTransRecLen
  NumOfRecords! = LOF(UBTrans) / UBTransRecLen
  
  UBRpt = FREEFILE
  OPEN "UBBKCNSP.RPT" FOR OUTPUT AS UBRpt
  
  BlockClear
  ShowProcessingScrn "Consumption by Book Report"
  
  REM Report Goes Here
  UBSetUp = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBSetUp LEN = UBSetUpRecLen
  IF LOF(UBSetUp) / UBSetUpRecLen = 0 THEN
    TownName$ = "Undefined"
  ELSE
    GET UBSetUp, 1, UBSetUpRec(1)
    TownName$ = UBSetUpRec(1).UTILNAME
    TownLen = LEN(RTRIM$(TownName$))
    TabStop = 40 - (TownLen / 2)
    IF TabStop < 1 THEN TabStop = 1
  END IF
  CLOSE UBSetUp
  
  
  BlockClear
  ShowProcessingScrn "Consumption by Unit Step Report"
  
  GOSUB DoUnitStepHeader
  
  FOR Cnt! = 1 TO NumOfRecords!
    GET UBTrans, Cnt!, UBTransRec(1)
    IF UBTransRec(1).TransDate >= BegDate AND UBTransRec(1).TransDate <= Enddate THEN
      IF UBTransRec(1).TransType = 1 OR UBTransRec(1).TransType = 101 THEN
        'Set Valid Flag to Off to Check Each Customer For Valid Rate Code
        ValidCustomer = 0
        IF LineCnt > MaxLines THEN
          PRINT #UBRpt, CHR$(12)
          GOSUB DoUnitStepHeader
        END IF
        
        'IF Valid Customer Then Get the Customer and Check to See if Rate Applies
        CustomerRecord = UBTransRec(1).CustAcctNo
        
        IF CustomerRecord > 0 THEN
          GET UBCust, CustomerRecord, UBCustRec(1)
          FOR Snt = 1 TO 15
            IF QPTrim$(UBCustRec(1).Serv(Snt).RateCode) = RateCode$ THEN
              MtrType$ = UBCustRec(1).Serv(Snt).RMtrType
              ValidCustomer = 1
              EXIT FOR
            END IF
          NEXT Snt
        END IF
        'IF Customer Has a Rate Code Match then Analyze Meter Type to Get Usage
        IF ValidCustomer = 1 THEN
          FOR MCnt = 1 TO 7
            IF UBTransRec(1).MtrTypes(MCnt) = Mtype OR UBTransRec(1).MtrTypes(MCnt) = Mtype1 THEN
              ' Get Usage Here
              MeterConsp& = MeterConsp& + UBTransRec(1).CurRead(MCnt) - UBTransRec(1).PrevRead(MCnt)
              IF MeterConsp& < 0 THEN
                MaxMeterAmt& = 10& ^ (LEN(STR$(UBTransRec(1).PrevRead(MCnt))) - 1)
                MeterConsp& = (MaxMeterAmt& - UBTransRec(1).PrevRead(MCnt)) + UBTransRec(1).CurRead(MCnt)
              END IF

              MeterConsp& = MeterConsp& * 100

              TMeterConsp# = TMeterConsp# + MeterConsp&
              MeterConsp& = 0
            END IF
          NEXT MCnt
        END IF

        'IF TMeterConsp# > 0 AND TMeterConsp# < 500 THEN
        '  STOP
        'END IF

        IF (TMeterConsp# >= 0) AND (ValidCustomer = 1) THEN

          'IF TMeterConsp# = 0 THEN
          '  'STOP
          '  ZeroCnt = ZeroCnt + 1
          'END IF
          
          NonUpdated = 1        'Set Flag to Let Me Know When this Cust Cons Used
          FOR LL = 1 TO MaxStep
            IF TMeterConsp# >= TblBreak&(LL - 1) AND TMeterConsp# <= TblBreak&(LL) THEN
              TotalConsp#(LL) = TotalConsp#(LL) + TMeterConsp#
              TotalCust(LL) = TotalCust(LL) + 1
              NonUpdated = 0
              EXIT FOR
            END IF
          NEXT LL
          IF NonUpdated = 1 THEN
            TotalConsp#(MaxStep) = TotalConsp#(MaxStep) + TMeterConsp#
            TotalCust(MaxStep) = TotalCust(MaxStep) + 1
          END IF
        END IF
        TMeterConsp# = 0
        
        IF AskAbandonPrint% THEN
          AbortFlag = True
          EXIT FOR
        END IF
      END IF
    END IF

    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF

    ShowPctCompL CLNG(Cnt!), CLNG(NumOfRecords!)
    
  NEXT
  
  GOSUB DoUnitStepFooter:
  
  CLOSE
  
  ERASE TblBreak&, TotalConsp#, TotalCust
  
  IF NOT AbortFlag THEN
    PrintRptFile "Consumption by Step", "UBBKCNSP.RPT", 1, RetCode, EntryPoint
  END IF
  
  'KillFile "UBBKCNSP.RPT"
  EXIT SUB
  
DoUnitStepHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(29); "Consumption by RateCode"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    For Rate Code: "; Form$(3, 0)
  PRINT #UBRpt, "Period Beginning : "; Form$(1, 0)
  PRINT #UBRpt, "    Period Ending: "; Form$(2, 0)
  PRINT #UBRpt, "      Report Date: "; DATE$
  PRINT #UBRpt, ""
  PRINT #UBRpt, Dash80$
  LineCnt = 6
  RETURN
  
DoUnitStepFooter:
  TblBreak&(MaxStep) = 99999999
  'TblBreak&(MaxStep + 1) = 99999999

  FOR LL = 1 TO MaxStep
    PRINT #UBRpt, "Step # "; LL;
    PRINT #UBRpt, TAB(12); "From "; TblBreak&(LL - 1); " to "; TblBreak&(LL)
    PRINT #UBRpt, "Consumption = "; USING "#########,#"; TotalConsp#(LL);
    PRINT #UBRpt, "  # of Cust = "; USING "#####,#"; TotalCust(LL);
    IF TotalCust(LL) > 0 THEN
      PRINT #UBRpt, "  Avg Use= "; USING "#####,#.##"; TotalConsp#(LL) / TotalCust(LL)
    ELSE
      PRINT #UBRpt, ""
    END IF
    PRINT #UBRpt, Dash80$
    
    
    
  NEXT LL
  PRINT #UBRpt, CHR$(12);
  RETURN
  GOTO ExitConsStep
  
ConsCheckReqFields:
  BadDate = False
  StartDate = Date2Num(Form$(1, 0))
  Enddate = Date2Num(Form$(2, 0))
  
  IF StartDate < 0 OR Enddate < 0 THEN
    Frm(1).FldNo = 1
    BadDate = True
  ELSEIF StartDate > Enddate THEN
    BadDate = True
  END IF
  
  IF BadDate THEN
    SaveScrn TempScrn()
    DisplayUBScrn "BADRDATE"
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 1
    Action = 2
  ELSE
    ReqFldsOK = True
  END IF
  
  
  RETURN
  
ExitConsStep:
  CLOSE
  
END SUB

SUB CustbyRate
  'here
  
  NumOfRateRecs = GetNumRateRecs
  
  REDIM UBRateTblRec(1 TO NumOfRateRecs) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))
  
  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRateRecs
    GET UBFile, Cnt, UBRateTblRec(Cnt)
  NEXT
  CLOSE
  
  REDIM CustCnt(1 TO NumOfRateRecs) AS INTEGER
  
  '***************
  MaxLines = 55
  PageNo = 0
  Dash80$ = STRING$(80, "-")
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetUp(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetUp(1))
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  NumOfRecs = LOF(UBCust) \ UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBCBRATE.RPT" FOR OUTPUT AS UBRpt
  
  UBSetUp = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBSetUp LEN = UBSetUpRecLen
  IF LOF(UBSetUp) / UBSetUpRecLen = 0 THEN
    TownName$ = "Undefined"
  ELSE
    GET UBSetUp, 1, UBSetUp(1)
    TownName$ = UBSetUp(1).UTILNAME
    TownLen = LEN(RTRIM$(TownName$))
    TabStop = 40 - (TownLen / 2)
    IF TabStop < 1 THEN TabStop = 1
  END IF
  CLOSE UBSetUp
  
  
  BlockClear
  ShowProcessingScrn "Customer By Rate Code Report."
  
  GOSUB DOCustRateHeader
  
  FOR Cnt = 1 TO NumOfRecs
    GET UBCust, Cnt, UBCustRec(1)
    IF UBCustRec(1).STATUS = "A" THEN
      'IF LineCnt > MaxLines THEN
      '  PRINT #UBRpt, CHR$(12)
      '  GOSUB DOCustRateHeader
      'END IF
      FOR RCnt = 1 TO 15
        '*************************************
        '   Main body of Printing goes here
        ThisCode$ = QPTrim$(UBCustRec(1).Serv(RCnt).RateCode)
        IF LEN(ThisCode$) > 0 THEN
          FOR ZCnt = 1 TO NumOfRateRecs
            IF QPTrim$(UBRateTblRec(ZCnt).RateCode) = ThisCode$ THEN
              CustCnt(ZCnt) = CustCnt(ZCnt) + 1
              EXIT FOR
            END IF
          NEXT
        END IF
      NEXT
      'LineCnt = LineCnt + 1
      CustomerCnt = CustomerCnt + 1
      '*************************************
      IF AskAbandonPrint% THEN
        AbortFlag = True
        EXIT FOR
      END IF
    END IF
    ShowPctComp Cnt, NumOfRecs
  NEXT
  FOR Cnt = 1 TO NumOfRateRecs
    PRINT #UBRpt, UBRateTblRec(Cnt).RateCode; TAB(12); UBRateTblRec(Cnt).RateDesc; TAB(61); USING "####"; CustCnt(Cnt)
  NEXT
  
  GOSUB DOCustRateFooter
  
  CLOSE
  
  ERASE UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Customer By Rate Code Report.", "UBCBRATE.RPT", 1, RetCode, EntryPoint
  END IF
  
  'KillFile "UBCOLIST.RPT"
  
ExitCustRateReport:
  
  EXIT SUB
  
DOCustRateHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(26); "Customers By Rate Code Report"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$;
  PRINT #UBRpt, ""
  PRINT #UBRpt, "Rate"; TAB(12); "Description"; TAB(60); "Count"                '; TAB(43); "Service Address"; TAB(72); "Balance"
  PRINT #UBRpt, Dash80$
  LineCnt = 6
  RETURN
  
DOCustRateFooter:
  PRINT #UBRpt, ""
  PRINT #UBRpt, "Total Customers Count: "; USING "####,#"; CustomerCnt
  RETURN
  
  
END SUB

SUB nShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&) STATIC
  Pct$ = SPACE$(3)
  RSET Pct$ = QPSTRL$(INT((RecNo& / NumOfRecs&) * 100))
  QPrintRC Pct$, 13, 40, Cnf.HiLite
END SUB

SUB RateAnalysisMenu
  
  REDIM MChoice$(1 TO 2)
  
  MChoice$(1) = "Rate Analysis by RATE CODE"
  MChoice$(2) = "Rate Analysis by REVENUE SOURCE"
  
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
  Help$ = "Statistical Reports Menu"
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Statistical Reports Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      AnalysisByRATE
    CASE 2
      AnalysisByREVENUE
    END SELECT
  LOOP
  
END SUB

SUB UsageByBook
  SHARED Choice$()
  REDIM Book(999), TotalConsp#(999), TotalCust(999)
  
  REDIM TempScrn(0)
  MaxLines = 56
  Dash80$ = STRING$(80, "-")
  
  LibName$ = "UB"
  
  
  ScrnName$ = "BKCONRTP"
  'Add
  MActionRow = 18
  
  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  Action = 1
  Frm(1).StayOnField = True
  
  NumOfRates = GetNumRateRecs%
  REDIM UBRateTbls(1 TO NumOfRates) AS UBRateTblRecType
  REDIM Choice$(0 TO NumOfRates + 10, 1)
  
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRates
    GET RateFile, Cnt, UBRateTbls(Cnt)
    Choice$(Cnt, 0) = UCASE$(UBRateTbls(Cnt).RateCode + " " + UBRateTbls(Cnt).RateDesc) + " " + STR$(Cnt)
  NEXT
  CLOSE
  
  Choice$(0, 0) = "3"
  
  Choice$(0, 1) = "4"
  Choice$(1, 1) = "Water Only"
  Choice$(2, 1) = "Sewer Only"
  Choice$(3, 1) = "Both Water/Sewer"
  Choice$(4, 1) = "Electric"
  Choice$(5, 1) = "Demand Elec"
  Choice$(6, 1) = "Gas"
  
  
  
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      
      GOSUB ConsBookConsFields:
      IF ReqFldsOK THEN
        ExitFlag = True
      ELSE
        Action = 2
      END IF
    CASE EscKey
      AbortFlag = True
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE MActionRow           'Look for the f10 or esc button
        
        SELECT CASE Frm(1).MCol
        CASE 28 TO 41           '--Cancel button
          PressButton 27, MActionRow, 28, 41
        CASE 42 TO 55           '--F10 button
          PressButton F10Key, MActionRow, 42, 55
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
  LOOP UNTIL ExitFlag
  IF AbortFlag THEN GOTO ExitBookConsStep
  
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  
  Mtype$ = RTRIM$(Form$(4, 0))
  IF Mtype$ = "Water Only" THEN Mtype = 1: Mtype1 = 3
  IF Mtype$ = "Sewer Only" THEN Mtype = 2: Mtype1 = 3
  IF Mtype$ = "Both Water/Sewe" THEN Mtype = 3: Mtype1 = 1
  IF Mtype$ = "Electric" THEN
    Mtype = 4
    Mtype1 = 9
  END IF
  IF Mtype$ = "Demand Elec" THEN Mtype = 5: Mtype1 = 9
  IF Mtype$ = "Gas" THEN Mtype = 6: Mtype1 = 9
  
  BegDate = Date2Num%(Form$(1, 0))
  Enddate = Date2Num%(Form$(2, 0))
  
  IndexName$ = BookIndexFile
  NumOfRecs = FileSize(IndexName$) \ 4
  REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
  FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
  
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))
  
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetUpRec(1))
  
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTrans LEN = UBTransRecLen
  NumOfRecords! = LOF(UBTrans) / UBTransRecLen
  
  UBRpt = FREEFILE
  OPEN "UBBOOKCN.RPT" FOR OUTPUT AS UBRpt
  
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Lnt = 1 TO NumOfRates
    IF LEFT$(Form$(3, 0), 5) = LEFT$(Choice$(Lnt, 0), 5) THEN
      GET RateFile, Lnt, UBRateTbls(1)
      RateCode$ = QPTrim$(UBRateTbls(1).RateCode)
      EXIT FOR
    END IF
  NEXT Lnt
  CLOSE RateFile

  'RateCode$ =

  BlockClear
  ShowProcessingScrn "Consumption by Book Report"
  
  REM Report Goes Here
  UBSetUp = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBSetUp LEN = UBSetUpRecLen
  IF LOF(UBSetUp) / UBSetUpRecLen = 0 THEN
    TownName$ = "Undefined"
  ELSE
    GET UBSetUp, 1, UBSetUpRec(1)
    TownName$ = UBSetUpRec(1).UTILNAME
    TownLen = LEN(RTRIM$(TownName$))
    TabStop = 40 - (TownLen / 2)
    IF TabStop < 1 THEN TabStop = 1
  END IF
  CLOSE UBSetUp
  
  
  BlockClear
  ShowProcessingScrn "Consumption by Book(Route) Report"
  
  GOSUB DoBookConsHeader
  
  FOR Cnt! = 1 TO NumOfRecords!
    GET UBTrans, Cnt!, UBTransRec(1)
    IF UBTransRec(1).TransDate >= BegDate AND UBTransRec(1).TransDate <= Enddate THEN
      IF UBTransRec(1).TransType = 1 OR UBTransRec(1).TransType = 101 THEN
        'Set Valid Flag to Off to Check Each Customer For Valid Rate Code
        ValidCustomer = 0
        'IF Valid Customer Then Get the Customer and Check to See if Rate Applies
        CustomerRecord = UBTransRec(1).CustAcctNo
        
        IF CustomerRecord > 0 THEN
          GET UBCust, CustomerRecord, UBCustRec(1)

          FOR Snt = 1 TO 15
            IF QPTrim$(UBCustRec(1).Serv(Snt).RateCode) = RateCode$ THEN
              MtrType$ = UBCustRec(1).Serv(Snt).RMtrType
              ValidCustomer = 1
              EXIT FOR
            END IF
          NEXT Snt
        END IF
        
        'IF Customer Has a Rate Code Match then Analyze Meter Type to Get Usage
        IF ValidCustomer = 1 THEN
          FOR MCnt = 1 TO 7
            IF (UBTransRec(1).MtrTypes(MCnt) = Mtype OR UBTransRec(1).MtrTypes(MCnt) = Mtype1) OR Mtype = 0 THEN
              ' Get Usage Here
              MeterConsp& = MeterConsp& + UBTransRec(1).CurRead(MCnt) - UBTransRec(1).PrevRead(MCnt)
              IF MeterConsp& < 0 THEN
                MaxMeterAmt& = 10& ^ (LEN(STR$(UBTransRec(1).PrevRead(MCnt))) - 1)
                MeterConsp& = (MaxMeterAmt& - UBTransRec(1).PrevRead(MCnt)) + UBTransRec(1).CurRead(MCnt)
              END IF
              MeterConsp& = MeterConsp& * 100
              TMeterConsp# = TMeterConsp# + MeterConsp&
              MeterConsp& = 0
            END IF
          NEXT MCnt
        END IF
        
        IF TMeterConsp# > 0 THEN
          NonUpdated = 1        'Set Flag to Let Me Know When this Cust Cons Used

          IF BookCnt = 0 THEN
            BookCnt = 1
            Book(BookCnt) = VAL(UBCustRec(1).Book)
            TotalConsp#(1) = TotalConsp#(1) + TMeterConsp#
            TotalCust(1) = TotalCust(1) + 1
            NonUpdated = 0
          END IF
          
          IF NonUpdated = 1 THEN
            FOR LL = 1 TO BookCnt
              IF VAL(UBCustRec(1).Book) = Book(LL) THEN
                TotalConsp#(LL) = TotalConsp#(LL) + TMeterConsp#
                TotalCust(LL) = TotalCust(LL) + 1
                NonUpdated = 0
                EXIT FOR
              END IF
            NEXT LL
          END IF
          
          IF NonUpdated = 1 THEN
            BookCnt = BookCnt + 1
            Book(BookCnt) = VAL(UBCustRec(1).Book)
            TotalConsp#(BookCnt) = TMeterConsp#
            TotalCust(BookCnt) = 1
            NonUpdated = 0
          END IF
        END IF
        
        TMeterConsp# = 0
        
        IF AskAbandonPrint% THEN
          AbortFlag = True
          EXIT FOR
        END IF
      END IF
    END IF
    ShowPctCompL CLNG(Cnt!), CLNG(NumOfRecords!)
    
  NEXT
  
  GOSUB DoBookConsFooter:
  
  CLOSE
  
  ERASE Book, TotalConsp#, TotalCust
  
  IF NOT AbortFlag THEN
    PrintRptFile "Consumption by Step", "UBBOOKCN.RPT", 1, RetCode, EntryPoint
  END IF
  
  KillFile "UBBOOKCN.RPT"
  EXIT SUB
  
DoBookConsHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(29); "Consumption by Book(Route)"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    For Rate Code: "; Form$(3, 0)
  PRINT #UBRpt, "Period Beginning : "; Form$(1, 0)
  PRINT #UBRpt, "    Period Ending: "; Form$(2, 0)
  PRINT #UBRpt, "      Report Date: "; DATE$
  PRINT #UBRpt, " "
  PRINT #UBRpt, "Book Number"; TAB(15); "Customers"; TAB(30); "Tot Consumption"; TAB(50); "Avg Cons/Cust"
  PRINT #UBRpt, Dash80$
  LineCnt = 6
  RETURN
  
DoBookConsFooter:
  
  GOSUB SortBooks
  
  FOR LL = 1 TO BookCnt
    
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB DoBookConsHeader
    END IF
    
    PRINT #UBRpt, USING "####"; Book(LL);
    PRINT #UBRpt, TAB(15); USING "#####,#"; TotalCust(LL);
    PRINT #UBRpt, TAB(30); USING "#########,#"; TotalConsp#(LL);
    IF TotalCust(LL) > 0 THEN
      PRINT #UBRpt, TAB(50); USING "#####,#.##"; TotalConsp#(LL) / TotalCust(LL)
    ELSE
      PRINT #UBRpt, ""
    END IF
    NumofBooks = NumofBooks + 1
    TotalCustomers# = TotalCustomers# + TotalCust(LL)
    TotalCons# = TotalCons# + TotalConsp#(LL)
    
  NEXT LL
  LineCnt = LineCnt + 1
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, USING "####"; NumofBooks;
  PRINT #UBRpt, TAB(15); USING "#####,#"; TotalCustomers#;
  PRINT #UBRpt, TAB(30); USING "#########,#"; TotalCons#;
  IF TotalCustomers# > 0 THEN
    PRINT #UBRpt, TAB(50); USING "#####,#.##"; TotalCons# / TotalCustomers#;
  END IF
  
  PRINT #UBRpt, TAB(70); "* TOTALS "
  
  
  PRINT #UBRpt, CHR$(12);
  RETURN
  GOTO ExitBookConsStep
  
ConsBookConsFields:
  BadDate = False
  StartDate = Date2Num(Form$(1, 0))
  Enddate = Date2Num(Form$(2, 0))
  
  IF StartDate < 0 OR Enddate < 0 THEN
    Frm(1).FldNo = 1
    BadDate = True
  ELSEIF StartDate > Enddate THEN
    BadDate = True
  END IF
  
  IF BadDate THEN
    SaveScrn TempScrn()
    DisplayUBScrn "BADRDATE"
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 1
    Action = 2
  ELSE
    ReqFldsOK = True
  END IF
  RETURN
SortBooks:
26000         REM sort
  Count = BookCnt
26020 m = BookCnt
26030 m = INT(m / 2)
26040 IF m = 0 THEN 26190
26050 FOR ST = 1 TO m
26060 i = ST
26070 j = ST + m
26080 SW = 0
26090 IF Book(i) <= Book(j) THEN 26120
26100 SW = 1
26110 SWAP Book(i), Book(j)
  SWAP TotalCust(i), TotalCust(j)
  SWAP TotalConsp#(i), TotalConsp#(j)
26120 i = j
26130 j = j + m
26140 IF j <= Count THEN 26090
26150 IF SW = 0 THEN 26170
26160 GOTO 26060
26170 NEXT ST
26180 GOTO 26030
26190 RETURN
  
  
ExitBookConsStep:
  CLOSE
END SUB

SUB UsageByRateCode
  
  SHARED Choice$()
  
  REDIM TempScrn(0)
  
  LibName$ = "UB"
  
  ScrnName$ = "RTCODRPT"
  'Add
  MActionRow = 17
  
  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  Action = 1
  Frm(1).StayOnField = True
  NumOfRates = GetNumRateRecs%
  REDIM UBRateTbls(1 TO NumOfRates) AS UBRateTblRecType
  REDIM Choice$(0 TO NumOfRates + 1, 1)
  
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  Choice$(1, 0) = "All-Revenues"
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRates
    GET RateFile, Cnt, UBRateTbls(Cnt)
    Choice$(Cnt + 1, 0) = UCASE$(UBRateTbls(Cnt).RateCode + " " + UBRateTbls(Cnt).RateDesc)
  NEXT
  CLOSE
  
  Choice$(0, 0) = "3"
  
  
  Choice$(0, 0) = "3"
  Choice$(0, 1) = "4"
  Choice$(1, 1) = "WATER/SEWER"
  Choice$(2, 1) = "ELECTRIC"
  Choice$(3, 1) = "GAS"
  
  
  
  
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB RateCheckReqFields:
      IF ReqFldsOK THEN
        ExitFlag = True
      ELSE
        Action = 2
      END IF
    CASE EscKey
      AbortFlag = True
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE MActionRow           'Look for the f10 or esc button
        
        SELECT CASE Frm(1).MCol
        CASE 28 TO 41           '--Cancel button
          PressButton 27, MActionRow, 28, 41
        CASE 42 TO 55           '--F10 button
          PressButton F10Key, MActionRow, 42, 55
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
  LOOP UNTIL ExitFlag
  IF AbortFlag THEN GOTO ExitCustByRate
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  REDIM RevDesc(1 TO MaxRevsCnt) AS STRING * 12
  FOR Cnt = 1 TO MaxRevsCnt     'find last active revenue
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(Cnt).REVNAME)
    IF LEN(TempRev$) = 0 THEN
      NumOfRevs = Cnt - 1       'set actual number of revenues
      EXIT FOR
    ELSE
      LSET RevDesc(Cnt) = UCASE$(TempRev$)
      'build revenue description line
    END IF
  NEXT
  
  NumOfRates = GetNumRateRecs%
  REDIM UBRateTbls(1 TO NumOfRates) AS UBRateTblRecType
  REDIM RateConsump(1 TO NumOfRates) AS LONG
  
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRates
    GET RateFile, Cnt, UBRateTbls(Cnt)
  NEXT
  CLOSE
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))
  
  REDIM RevTotals(1 TO NumOfRevs)  AS DOUBLE    'holds each revenues total amt
  REDIM ConsumpTot(1 TO NumOfRevs, 1 TO 2) AS DOUBLE            'holds each revenues total amt
  REDIM BookConsump(1 TO 1) AS BookConsumpType
  
  'This is the BookConsumptionType construct.
  'type BookConsumpType
  '    Book             AS INTEGER    'holds this book number
  '    CustCnt          AS LONG       'holds the cust count for this book
  '    Consump(1 TO 15) AS LONG       'Holds the consump amts for each revenue
  '    RevAmt(1 TO 15)  AS DOUBLE      'Holds total revenues for this book
  
  IndexName$ = BookIndexFile
  NumOfRecs = FileSize(IndexName$) \ 4
  REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
  FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBBYRTSP.RPT" FOR OUTPUT AS UBRpt
  
  
  BlockClear
  ShowProcessingScrn "Customer Consumption by Rate Report"
  
  CLOSE
  
  SLEEP 4
  
  SaveScrn TempScrn()
  DisplayUBScrn "NOTENGH"
  WaitForAction
  RestScrn TempScrn()
  Frm(1).FldNo = 1
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  GOTO ExitCustByRate
  
RateCheckReqFields:
  BadDate = False
  StartDate = Date2Num(Form$(1, 0))
  Enddate = Date2Num(Form$(2, 0))
  
  IF StartDate < 0 OR Enddate < 0 THEN
    Frm(1).FldNo = 1
    BadDate = True
  ELSEIF StartDate > Enddate THEN
    BadDate = True
  END IF
  
  IF BadDate THEN
    SaveScrn TempScrn()
    DisplayUBScrn "BADRDATE"
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 1
    Action = 2
  ELSE
    ReqFldsOK = True
  END IF
  
  
  RETURN
  
ExitCustByRate:
  
END SUB

SUB UsageByZoneCode
  SHARED Choice$()
  
  REDIM TempScrn(0)
  
  LibName$ = "UB"
  ScrnName$ = "ZOCODRPT"
  MActionRow = 17
  
  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  Action = 1
  Frm(1).StayOnField = True
  
  '--Set screen number to one and display screen
  
  REDIM Choice$(0 TO 3, 0)
  Choice$(0, 0) = "4"
  Choice$(1, 0) = "WATER/SEWER"
  Choice$(2, 0) = "ELECTRIC"
  Choice$(3, 0) = "GAS"
  
  
  
  BlockClear
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB ZoneCheckReqFields
      IF ReqFldsOK THEN
        ExitFlag = True
      ELSE
        Action = 2
      END IF
    CASE EscKey
      AbortFlag = True
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE MActionRow           'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 28 TO 41           '--Cancel button
          PressButton 27, MActionRow, 28, 41
        CASE 42 TO 55           '--F10 button
          PressButton F10Key, MActionRow, 42, 55
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN GOTO ExitCustByZone
  
  CLOSE
  
  ShowProcessingScrn "Customer Consumption by Zone Report"
  
  CLOSE
  
  SLEEP 4
  
  SaveScrn TempScrn()
  DisplayUBScrn "NOTENGH"
  WaitForAction
  RestScrn TempScrn()
  Frm(1).FldNo = 1
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  
  GOTO ExitCustByZone
  
ZoneCheckReqFields:
  BadDate = False
  StartDate = Date2Num(Form$(1, 0))
  Enddate = Date2Num(Form$(2, 0))
  
  IF StartDate < 0 OR Enddate < 0 THEN
    Frm(1).FldNo = 1
    BadDate = True
  ELSEIF StartDate > Enddate THEN
    BadDate = True
  END IF
  
  IF BadDate THEN
    SaveScrn TempScrn()
    DisplayUBScrn "BADRDATE"
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 1
    Action = 2
  ELSE
    ReqFldsOK = True
  END IF
  
  
  RETURN
  
ExitCustByZone:
  
END SUB

