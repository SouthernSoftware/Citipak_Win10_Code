DEFINT A-Z
DECLARE SUB UsageByType ()
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
  '$INCLUDE: 'ubrate.bi'
  '$INCLUDE: 'ubTrans.bi'
  
  CONST False = 0, True = NOT False
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 6)
  
  MChoice$(1) = "Water Consumption by 1000's Step"
  
  MChoice$(2) = "Exit to OS"
  
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
  Help$ = "Consumption Report Menu"
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 2, Col, MaxLen + 3, "Consumption Report Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      UsageByType
    CASE 2
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
  RUN "ubmenu"

SUB UsageByType
  
  SHARED Choice$()
  MaxStep = 101

  REDIM TblBreak&(MaxStep), TotalConsp#(MaxStep), TotalCust(MaxStep), AvgCon#(MaxStep)
  
  REDIM TempScrn(0)
  MaxLines = 56
  Dash80$ = STRING$(80, "-")
  
  LibName$ = "UBlong"
  ScrnName$ = "CONCUTYP"

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
  
  FirstTime = True
  
  '--Set screen number to one and display screen
  
  BlockClear
  
  LibFile2Scrn LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%

  QPrintRC "Consumption By 1000's Step", 5, 23, -1
  QPrintRC " ]ÄÄÄ", 5, 49, 10

  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      GOSUB TypeCheckFlds:
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
  IF AbortFlag THEN GOTO ExitTypeStep
  
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Lnt = 1 TO NumOfRates
    WhatRate$ = QPTrim$(LEFT$(Form$(3, 0), 4))
    ThisRate$ = QPTrim$(LEFT$(Choice$(Lnt, 0), 4))
    IF ThisRate$ = WhatRate$ THEN
      GET RateFile, Lnt, UBRateTbls(1)
      RateCode$ = QPTrim$(UBRateTbls(1).RateCode)
      
      'FOR Tnt = 1 TO 10
      '  IF UBRateTbls(Lnt).TblBreaks(Tnt).Units >= 0 THEN
      '    IF UBRateTbls(Lnt).TblBreaks(Tnt).Units > 0 THEN
      '     Greater = True
      '    END IF
      '    IF (UBRateTbls(Lnt).TblBreaks(Tnt).Units = 0) AND (Greater = True) THEN
      '      MaxStep = Tnt
      '      TblBreak&(Tnt) = 99999999
      '      EXIT FOR
      '    END IF
      '
      '    TblBreak&(Tnt) = UBRateTbls(Lnt).TblBreaks(Tnt).Units
      '    MaxStep = Tnt
      '  ELSE
      '    IF MaxStep = 1 THEN
      '      MaxStep = 2
      '      TblBreak&(MaxStep) = 99999999
      '    END IF
      '    EXIT FOR
      '  END IF
      'NEXT Tnt
    END IF
  NEXT Lnt
  CLOSE

  FOR zz = 1 TO MaxStep - 1
    TblBreak&(zz) = (0& + zz) * 1000
  NEXT

'  TblBreak&(1) = 3000
'  TblBreak&(2) = 5000
'  TblBreak&(3) = 7000
'  TblBreak&(4) = 10000
'  TblBreak&(5) = 999999999
  'MaxStep = 101

  IF MaxStep = 0 THEN
    MaxStep = 1
    TblBreak&(MaxStep) = 999999999
  END IF

  TblBreak&(MaxStep) = 999999999

  BegDate = Date2Num%(Form$(1, 0))
  Enddate = Date2Num%(Form$(2, 0))
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetUpRec(1))
  
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
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  NumOfCust& = LOF(UBCust) / UBCustRecLen
  
  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTrans LEN = UBTransRecLen
  
  UBRpt = FREEFILE
  OPEN "UBCUTYPE.RPT" FOR OUTPUT AS UBRpt
  
  REM Report Goes Here
  
  BlockClear
  ShowProcessingScrn "Water Consumption by 1000's Step"
  
  GOSUB DoTypeStepHeader
  
  FOR Cnt& = 1 TO NumOfCust&
    GET UBCust, Cnt&, UBCustRec(1)
    CType$ = UCASE$(QPTrim$(UBCustRec(1).CustType))
    FOR Serv = 1 TO 15
      IF INSTR(UBCustRec(1).Serv(Serv).RateCode, WhatRate$) > 0 THEN
        GOTO UseThisCust
      END IF
    NEXT
    GOTO NotThisCust
UseThisCust:
    IF CType$ = ChkType$ THEN
      TranRec& = UBCustRec(1).LastTrans
      TRCnt = 0
      DO WHILE TranRec& > 0
        GET UBTrans, TranRec&, UBTransRec(1)
        IF UBTransRec(1).TransDate >= BegDate AND UBTransRec(1).TransDate <= Enddate THEN
          IF UBTransRec(1).TransType = 1 OR UBTransRec(1).TransType = 101 THEN
            TRCnt = TRCnt + 1
            Got1 = False
            FOR MCnt = 1 TO 7
              SELECT CASE UBTransRec(1).MtrTypes(MCnt)
              CASE 0, 1, 2, 3, 7
                Got1 = True
                MeterConsp& = MeterConsp& + UBTransRec(1).CurRead(MCnt) - UBTransRec(1).PrevRead(MCnt)
                IF MeterConsp& < 0 THEN
                  MaxMeterAmt& = 10& ^ (LEN(STR$(UBTransRec(1).PrevRead(MCnt))) - 1)
                  MeterConsp& = (MaxMeterAmt& - UBTransRec(1).PrevRead(MCnt)) + UBTransRec(1).CurRead(MCnt)
                END IF
                'MeterConsp& = MeterConsp& * 100
                TMeterConsp# = TMeterConsp# + MeterConsp&
                MeterConsp& = 0
              END SELECT
            NEXT MCnt
        '    IF NOT Got1 THEN STOP
          END IF
        END IF
        TranRec& = UBTransRec(1).PrevTrans
      LOOP

      IF (TMeterConsp# > 0) THEN
        AvgConsp# = TMeterConsp# / TRCnt
        NonUpdated = 1    'Set Flag to Let Me Know When this Cust Cons Used
        FOR LL = 1 TO MaxStep
          IF AvgConsp# >= TblBreak&(LL - 1) AND AvgConsp# <= TblBreak&(LL) THEN
            TotalConsp#(LL) = TotalConsp#(LL) + TMeterConsp#
            AvgCon#(LL) = AvgCon#(LL) + AvgConsp#
            TotalCust(LL) = TotalCust(LL) + 1
            NonUpdated = 0
            EXIT FOR
          END IF
        NEXT
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
    
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
NotThisCust:
    ShowPctCompL Cnt&, NumOfCust&
  NEXT
  
  GOSUB DoTypeStepFooter:
  
  CLOSE
  
  ERASE TblBreak&, TotalConsp#, TotalCust
  
  IF NOT AbortFlag THEN
    PrintRptFile "Consumption by Customer Type", "UBCUTYPE.RPT", 1, RetCode, EntryPoint
  END IF
  
  'KillFile "UBBKCNSP.RPT"
  EXIT SUB
  
DoTypeStepHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(29); "Consumption by Customer Type"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    Customer Type: "; ChkType$
  IF WhatRate$ > "" THEN
    PRINT #UBRpt, "    For Rate Code: "; Form$(3, 0)
  END IF
  PRINT #UBRpt, " Period Beginning: "; Form$(1, 0)
  PRINT #UBRpt, "    Period Ending: "; Form$(2, 0)
  PRINT #UBRpt, "      Report Date: "; DATE$
  PRINT #UBRpt, ""
  PRINT #UBRpt, Dash80$
  LineCnt = 6
  RETURN
  
DoTypeStepFooter:
  'TblBreak&(MaxStep) = 99999999
  'TblBreak&(MaxStep + 1) = 99999999
  
  FOR LL = 1 TO MaxStep
    IF TotalConsp#(LL) > 0 THEN

      IF MaxStep > 1 THEN
        PRINT #UBRpt, "Step # "; LL;
      END IF
      IF LL = MaxStep AND MaxStep > 1 THEN
        PRINT #UBRpt, TAB(10); "From "; TblBreak&(LL - 1); " to "; TblBreak&(LL); ">"
      ELSEIF MaxStep > 1 THEN
        PRINT #UBRpt, TAB(10); "From "; TblBreak&(LL - 1); " to "; TblBreak&(LL)
      END IF
      PRINT #UBRpt, "Consumption = "; USING "#########,#"; TotalConsp#(LL);
      PRINT #UBRpt, "  # of Cust = "; USING "#####,#"; TotalCust(LL);
      IF TotalCust(LL) > 0 THEN
        PRINT #UBRpt, "  Avg Use= "; USING "#######,#.##"; AvgCon#(LL) / TotalCust(LL)
      ELSE
        PRINT #UBRpt, ""
      END IF
      PRINT #UBRpt, Dash80$
    END IF
  NEXT LL
  PRINT #UBRpt, CHR$(12);
  RETURN
  GOTO ExitTypeStep
  
TypeCheckFlds:
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

  ChkType$ = UCASE$(QPTrim$(Form$(4, 0)))
  
RETURN
  
ExitTypeStep:
  CLOSE

END SUB

