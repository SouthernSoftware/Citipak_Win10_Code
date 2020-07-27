  DEFINT A-Z
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
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumofRecs%)
DECLARE SUB ClearScrn ()
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'PageInfo.BI'                  'Form Page info
  
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ubcust.bi'
  '$INCLUDE: 'ubLoc.bi'
  '$INCLUDE: 'ubrate.bi'
  '$INCLUDE: 'ubTrans.bi'
  ''$INCLUDE: 'ubmeter.BI'
  
  CONST False = 0, True = NOT False
  '--Dim the choice array to the number of menu items
  REDIM Mchoice$(1 TO 6)
  
  Mchoice$(1) = "Consumption by Unit Step"
  Mchoice$(2) = "Customer Usage by Book (Route)"
  Mchoice$(3) = "Customer Usage by Rate Code"
  Mchoice$(4) = "Customer Usage by Zone Code"
  Mchoice$(5) = "Rate Analysis"
  Mchoice$(6) = "Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 20   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(Mchoice$)
    TLen = LEN(Mchoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(Mchoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 2
  Help$ = "Statistical Reports Menu"
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Statistical Reports Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      ConsumpUnitStep
    CASE 2
      UsageByBook
    CASE 3
      UsageByRateCode
    CASE 4
      UsageByZoneCode
    CASE 5
      RateAnalysisMenu
    CASE 6
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
  REDIM Ratecode$(100), RateDesc$(100), RateTrans(100), RateUse#(100), RateRev#(100)
  
  
  LibName$ = "UB"
  ScrnName$ = "ANRATRPT"
  MActionRow = 17
  
  Dash80$ = STRING$(80, "-")
  
  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), form$(), ErrCode
  
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
    Choice$(Cnt + 1, 0) = UCASE$(UBRateTbls(Cnt).Ratecode + " " + UBRateTbls(Cnt).RateDesc)
  NEXT
  CLOSE
  
  Choice$(0, 0) = "3"
  
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm form$(), Fld(), Frm(1), Cnf, Action
    
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
  
  REDIM UBCustRec(1) AS UBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetUp(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetUp(1))
  
  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))
  
  REDIM UBLocaRec(1) AS UBLocationRecType
  UBLocaRecLen = LEN(UBLocaRec(1))
  
  
  UBRpt = FREEFILE
  OPEN "UBRATEAN.RPT" FOR OUTPUT AS UBRpt
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBLoca = FREEFILE
  OPEN "UBLOCA.DAT" FOR RANDOM SHARED AS UBLoca LEN = UBLocaRecLen
  
  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS UBTrans LEN = UBTransRecLen
  
  NumofRecs! = LOF(UBTrans) / UBTransRecLen
  
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
    
    RateDesc$(1) = Ratecode$(RCnt)
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
  RatetoFind$ = form$(3, 0)
  IF LEFT$(RatetoFind$, 3) = "All" THEN
    RatetoFind$ = "ALL"
  ELSE
    RatetoFind$ = RTRIM$(RatetoFind$)
  END IF
  
  FOR Cnt! = 1 TO NumofRecs!
    ShowPctComp Cnt!, NumofRecs!
    
    GET UBTrans, Cnt!, UBTransRec(1)
    IF UBTransRec(1).TransDate < StartDate THEN
      BadCount = BadCount + 1
      IF BadCount > 500 THEN
        RETURN
      END IF
    END IF
    IF (UBTransRec(1).TransDate >= StartDate AND UBTransRec(1).TransDate <= EndDate) THEN
      IF UBTransRec(1).TransType = 1 THEN
        'rem rate code and check against what was asked for
        GET UBLoca, UBTransRec(1).CustLocation, UBLocaRec(1)
        'find all rate codes
        FOR RCnt = 1 TO MaxRevenue
          Ratecode$(RCnt) = UBLocaRec(1).Serv(RCnt).Ratecode
        NEXT RCnt
        'if Rate Codes = All then Store All of Them for printing else Findit
        FOR RCnt = 1 TO MaxRevenue
          IF RatetoFind$ = "ALL" THEN
            GOSUB CalcUse
            GOSUB GetRev
            IF RateCnt = 0 THEN
              RateCnt = 1
              RateDesc$(1) = Ratecode$(RCnt)
              RateTrans(1) = 1
              RateUse#(1) = CalcUsage#
              RateRev#(1) = RevenueTotal#
            ELSE
              FOR ll = 1 TO RateCnt
                IF RateDesc$(ll) = Ratecode$(RCnt) THEN
                  RateTrans(ll) = RateTrans(ll) + 1
                  RateUse#(ll) = RateUse#(ll) + CalcUsage#
                  RateRev#(ll) = RateRev#(ll) + RevenueTotal#
                  NoProcess = 1
                END IF
              NEXT ll
              IF NoProcess = 0 THEN
                RateCnt = RateCnt + 1
                RateDesc$(RateCnt) = Ratecode$(RCnt)
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
  StartDate = Date2Num(form$(1, 0))
  EndDate = Date2Num(form$(2, 0))
  Ratecode$ = QPTrim$(form$(3, 0))
  
  IF StartDate < 0 OR EndDate < 0 THEN
    Frm(1).FldNo = 1
    BadDate = True
  ELSEIF StartDate > EndDate THEN
    BadDate = True
  END IF
  
  IF BadDate OR LEN(Ratecode$) = 0 THEN
    SaveScrn TempScrn()
    DisplayUBScrn "BADRDATE"
    WaitForAction
    RestScrn TempScrn()
    Frm(1).FldNo = 1
    Action = 2
  ELSE
    ReqFldsOK = True
  END IF
  
  IF Choice$(1, 0) <> Ratecode$ THEN
    FOR Cnt = 2 TO NumOfRates
      IF INSTR(Choice$(Cnt, 0), Ratecode$) THEN
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
  REDIM form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), form$(), ErrCode
  
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
    
    EditForm form$(), Fld(), Frm(1), Cnf, Action
    
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
  StartDate = Date2Num(form$(1, 0))
  EndDate = Date2Num(form$(2, 0))
  Revenue$ = QPTrim$(form$(3, 0))
  
  IF StartDate < 0 OR EndDate < 0 THEN
    Frm(1).FldNo = 1
    BadDate = True
  ELSEIF StartDate > EndDate THEN
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
  
  REDIM TempScrn(0)
  
  LibName$ = "UB"
  
  ScrnName$ = "CONUSTEP"
  'Add
  MActionRow = 18
  
  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), form$(), ErrCode
  
  Action = 1
  Frm(1).StayOnField = True
  
  NumOfRates = GetNumRateRecs%
  REDIM UBRateTbls(1 TO NumOfRates) AS UBRateTblRecType
  REDIM Choice$(0 TO NumOfRates + 1, 1)
  
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRates
    GET RateFile, Cnt, UBRateTbls(Cnt)
    Choice$(Cnt, 0) = UCASE$(UBRateTbls(Cnt).Ratecode + " " + UBRateTbls(Cnt).RateDesc)
  NEXT
  CLOSE
  
  Choice$(0, 0) = "3"

  Choice$(0, 1) = "4"
  Choice$(1, 1) = "WATER/SEWER"
  Choice$(2, 1) = "ELECTRIC"
  Choice$(3, 1) = "GAS"
  
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm form$(), Fld(), Frm(1), Cnf, Action
    
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
      IF INSTR(RevDesc(Cnt), "ELECTRIC") THEN
        DoFuelAdjFlag = True
      END IF
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
  
  REDIM UBCustRec(1) AS UBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBLocaRec(1) AS UBLocationRecType
  UBLocaRecLen = LEN(UBLocaRec(1))
  
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
  NumofRecs = FileSize(IndexName$) \ 4
  REDIM IndexArray(1 TO NumofRecs) AS UBCustIndexRecType
  FGetAH IndexName$, IndexArray(1), 4, NumofRecs
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBBKCNSP.RPT" FOR OUTPUT AS UBRpt
  
  BlockClear
  ShowProcessingScrn "Consumption by Book Report"
  
  CLOSE
  SLEEP 5
  
  SaveScrn TempScrn()
  DisplayUBScrn "NOTENGH"
  WaitForAction
  RestScrn TempScrn()
  Frm(1).FldNo = 1
  
  GOTO ExitConsStep
  
ConsCheckReqFields:
  BadDate = False
  StartDate = Date2Num(form$(1, 0))
  EndDate = Date2Num(form$(2, 0))
  
  IF StartDate < 0 OR EndDate < 0 THEN
    Frm(1).FldNo = 1
    BadDate = True
  ELSEIF StartDate > EndDate THEN
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

SUB RateAnalysisMenu
  
  REDIM Mchoice$(1 TO 2)
  
  Mchoice$(1) = "Rate Analysis by RATE CODE"
  Mchoice$(2) = "Rate Analysis by REVENUE SOURCE"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 20   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(Mchoice$)
    TLen = LEN(Mchoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(Mchoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  Help$ = "Statistical Reports Menu"
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Statistical Reports Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu Mchoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
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
  
  REDIM TempScrn(0)
  
  LibName$ = "UB"
  
  ScrnName$ = "BKCONRPT"
  MActionRow = 17
  
  '--define the multi-choice fields
  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), form$(), ErrCode
  
  
  
  
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
    
    EditForm form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB BookCheckReqFields:
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
  
  IF AbortFlag THEN GOTO ExitCustByBook
  
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
      IF INSTR(RevDesc(Cnt), "ELECTRIC") THEN
        DoFuelAdjFlag = True
      END IF
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
  
  REDIM UBCustRec(1) AS UBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBLocaRec(1) AS UBLocationRecType
  UBLocaRecLen = LEN(UBLocaRec(1))
  
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
  NumofRecs = FileSize(IndexName$) \ 4
  REDIM IndexArray(1 TO NumofRecs) AS UBCustIndexRecType
  FGetAH IndexName$, IndexArray(1), 4, NumofRecs
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBBKCNSP.RPT" FOR OUTPUT AS UBRpt
  
  
  BlockClear
  ShowProcessingScrn "Customer Consumption by Book Report"
  
  CLOSE
  '  WaitForAction
  
  CLOSE
  SLEEP 5
  
  SaveScrn TempScrn()
  DisplayUBScrn "NOTENGH"
  WaitForAction
  RestScrn TempScrn()
  Frm(1).FldNo = 1
  
  
  
  
  
  GOTO ExitCustByBook
  
BookCheckReqFields:
  BadDate = False
  StartDate = Date2Num(form$(1, 0))
  EndDate = Date2Num(form$(2, 0))
  
  IF StartDate < 0 OR EndDate < 0 THEN
    Frm(1).FldNo = 1
    BadDate = True
  ELSEIF StartDate > EndDate THEN
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
  
ExitCustByBook:
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
  REDIM form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), form$(), ErrCode
  
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
    Choice$(Cnt + 1, 0) = UCASE$(UBRateTbls(Cnt).Ratecode + " " + UBRateTbls(Cnt).RateDesc)
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
    
    EditForm form$(), Fld(), Frm(1), Cnf, Action
    
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
  
  REDIM UBCustRec(1) AS UBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBLocaRec(1) AS UBLocationRecType
  UBLocaRecLen = LEN(UBLocaRec(1))
  
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
  NumofRecs = FileSize(IndexName$) \ 4
  REDIM IndexArray(1 TO NumofRecs) AS UBCustIndexRecType
  FGetAH IndexName$, IndexArray(1), 4, NumofRecs
  
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
  StartDate = Date2Num(form$(1, 0))
  EndDate = Date2Num(form$(2, 0))
  
  IF StartDate < 0 OR EndDate < 0 THEN
    Frm(1).FldNo = 1
    BadDate = True
  ELSEIF StartDate > EndDate THEN
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
  REDIM form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), form$(), ErrCode
  
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
    
    EditForm form$(), Fld(), Frm(1), Cnf, Action
    
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
  StartDate = Date2Num(form$(1, 0))
  EndDate = Date2Num(form$(2, 0))
  
  IF StartDate < 0 OR EndDate < 0 THEN
    Frm(1).FldNo = 1
    BadDate = True
  ELSEIF StartDate > EndDate THEN
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

