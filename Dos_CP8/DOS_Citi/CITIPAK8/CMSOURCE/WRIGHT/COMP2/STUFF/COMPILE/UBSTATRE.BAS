DEFINT A-Z
DECLARE SUB ConByRateDate ()
DECLARE FUNCTION FUsing$ (Number$, Fmt$)
DECLARE SUB IrrConSteps ()
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
DECLARE SUB TotalUsageByDate ()
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
DECLARE FUNCTION Round# (DblNum#)
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
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ClearScrn ()

  '$INCLUDE: 'DefCnf.BI'
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
  
  MChoice$(1) = "Consumption by Rate Code"
  MChoice$(2) = "Irrigation Consumption by Rate Code"
  MChoice$(3) = "Customer Count by Rate Code"
  MChoice$(4) = "Consumption by Rate & Date"
  MChoice$(5) = "Exit to OS"
  
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
  Col = ((80 - MaxLen) \ 2)

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
      IrrConSteps
    CASE 3
      CustbyRate
    CASE 4
      ConByRateDate
    CASE 5
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
  IF INSTR(COMMAND$, "TEST") THEN
    HideCursor
    ClearScrn
    END
  END IF
  
  RUN "ubmenu"

SUB ConByRateDate

  SHARED Choice$()

  MaxLines = 59
  PageNo = 0
  Dash80$ = STRING$(103, "-")

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTransRec(1) AS UBTransRecType
  UBTransRecLen = LEN(UBTransRec(1))

  ThruDate = Date2Num%(DATE$)
  FromDate = ThruDate - 365
  Thru$ = DATE$
  From$ = Num2Date$(FromDate)

'*****************

  LibName$ = "UB"
  ScrnName$ = "CONRTDAT"

  '--Get the total number of fields from all pages
  NumFlds = NumFlds + LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  '--Increment StartEl to next

  REDIM UBRateTblRec(1) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))
  NumOfRateRecs = GetNumRateRecs

  'add show a message
  IF NumOfRateRecs <= 0 THEN
    GOTO oExitCustByRate
  END IF

  REDIM Choice$(0 TO NumOfRateRecs, 0 TO 2)
  Choice$(0, 0) = "3"
  Choice$(0, 1) = "4"
  Choice$(0, 2) = "7"
  Choice$(1, 1) = "YES"
  Choice$(2, 1) = "NO"

  Choice$(1, 2) = "Combined"
  Choice$(2, 2) = "Water"
  Choice$(3, 2) = "Sewer"
  Choice$(4, 2) = "Irrigation"

  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRateRecs
    GET UBFile, Cnt, UBRateTblRec(1)
    Choice$(Cnt, 0) = SPACE$(30)
    LSET Choice$(Cnt, 0) = UBRateTblRec(1).RATECODE
    MID$(Choice$(Cnt, 0), 6, 24) = UBRateTblRec(1).RATEDESC
  NEXT
  CLOSE

  RateRec = 0
  PrnRate = False

  Action = 1
  FirstTime = True

  BlockClear

  DisplayUBScrn ScrnName$

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = From$
      LSET Form$(2, 0) = Thru$
      LSET Form$(4, 0) = "YES"
      LSET Form$(5, 0) = "0"
      LSET Form$(6, 0) = "999999999999"
      Action = 1
    END IF

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key 'F10
      OkFlag = True
      WhatRate$ = QPTrim$(LEFT$(Form$(3, 0), 4))
      FromDate = Date2Num(Form$(1, 0))
      ThruDate = Date2Num(Form$(2, 0))
      ActiveFlag = QPTrim$(Form$(4, 0)) = "YES"
      MinUnits# = VAL(Form$(5, 0))
      MaxUnits# = VAL(Form$(6, 0))
      RptType$ = QPTrim$(Form$(7, 0))
      IF LEN(RptType$) = 0 THEN
        OkFlag = False
        Frm(1).FldNo = 7
        Action = 1
      ELSE
        IF INSTR(RptType$, "Irr") > 0 THEN
          FromCnt = 5
          ThruCnt = 5
        ELSEIF INSTR(RptType$, "Wat") > 0 THEN
          FromCnt = 1
          ThruCnt = 1
          WaterOnly = True
        ELSEIF INSTR(RptType$, "Sew") > 0 THEN
          FromCnt = 2
          ThruCnt = 2
          SewerOnly = True
        ELSE
          FromCnt = 1
          ThruCnt = 2
          Combined = True
        END IF
      END IF
      
      IF LEN(WhatRate$) = 0 THEN
        OkFlag = False
      END IF
      IF FromDate < -32000 OR ThruDate < -32000 OR FromDate > ThruDate THEN
        OkFlag = False
      END IF
    CASE EscKey
      ExitFlag = True
    END SELECT

  '--Check screen page
  LOOP UNTIL ExitFlag OR OkFlag

  IF ExitFlag THEN
    GOTO oExitCustByRate
  END IF
  
'****************

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  NumOfCust& = LOF(UBCust) / UBCustRecLen

  UBRpt = FREEFILE
  OPEN "UBCSBYRT.RPT" FOR OUTPUT AS UBRpt

  UBTrans = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTrans LEN = UBTransRecLen

  BlockClear
  ShowProcessingScrn "Consumption By Rate & Date Range."

  GOSUB CustByRateHeader

  FOR Cnt = 1 TO NumOfCust&
    GET UBCust, Cnt, UBCustRec(1)
    IF ActiveFlag THEN
      IF UBCustRec(1).Status <> "A" THEN
        GOTO ActiveOnlySkip
      END IF
    END IF
    IF UBCustRec(1).DelFlag = 0 THEN
      Valid = False
      Cubic = False
      Multi& = 0
      FOR SCnt = FromCnt TO ThruCnt
        ThisRate$ = QPTrim$(UBCustRec(1).Serv(SCnt).RATECODE)
        IF WhatRate$ = ThisRate$ THEN
          IF SewerOnly THEN
            TRate$ = QPTrim$(UBCustRec(1).Serv(1).RATECODE)
            IF LEN(TRate$) > 0 THEN
              GOTO NotThisOne
            END IF
          ELSEIF WaterOnly THEN
            TRate$ = QPTrim$(UBCustRec(1).Serv(2).RATECODE)
            IF LEN(TRate$) > 0 THEN
              GOTO NotThisOne
            END IF
          ELSEIF Combined THEN
            TRate1$ = QPTrim$(UBCustRec(1).Serv(1).RATECODE)
            TRate2$ = QPTrim$(UBCustRec(1).Serv(2).RATECODE)
            IF LEN(TRate1$) = 0 OR LEN(TRate2$) = 0 THEN
              GOTO NotThisOne
            ELSE

            END IF
          END IF
          MtrType$ = UBCustRec(1).Serv(SCnt).RMtrType
          FOR MCnt = 1 TO 7
            IF UBCustRec(1).LocMeters(MCnt).MtrType = MtrType$ THEN
              Multi& = UBCustRec(1).LocMeters(MCnt).MtrMulti
              IF Multi& = 0 THEN Multi& = 1
              IF UBCustRec(1).LocMeters(MCnt).MTRUnit = "C" THEN
                Cubic = True
              END IF
              EXIT FOR
            END IF
          NEXT
          SELECT CASE MtrType$
          CASE "W"
            MType = 1
          CASE "S"
            MType = 2
          CASE "C"
            MType = 3
          CASE "T"
            MType = 7
          CASE ELSE
            MType = 4
          END SELECT
          Valid = True
          EXIT FOR
        END IF
      NEXT
NotThisOne:
      IF Valid THEN
        CustTRCnt = 0
        CustConTot# = 0
        MeterConsp# = 0
        ThisTrans& = UBCustRec(1).LastTrans
        DO WHILE ThisTrans& > 0
          GET UBTrans, ThisTrans&, UBTransRec(1)
          'IF TRDate < FromDate THEN EXIT DO
          IF UBTransRec(1).TransType = 1 THEN
            TRDate = UBTransRec(1).TransDate
            IF TRDate >= FromDate AND TRDate <= ThruDate THEN
              CustTRCnt = CustTRCnt + 1
              FOR MCnt = 1 TO 7
                IF UBTransRec(1).MtrTypes(MCnt) = MType THEN
                  CurrRead# = UBTransRec(1).CurRead(MCnt)
                  PrevRead# = UBTransRec(1).PrevRead(MCnt)
                  ThisConsp& = CurrRead# - PrevRead#
                  IF ThisConsp& < 0 THEN
                    MaxMeterAmt# = 10& ^ (LEN(STR$(PrevRead#)) - 1)
                    ThisConsp& = (MaxMeterAmt# - PrevRead#) + CurrRead#
                  END IF
                  IF Cubic THEN
                    ThisConsp& = Round(ThisConsp& * 7.481)
                  END IF
                  ThisConsp& = Round#(ThisConsp& * Multi&)
                  IF ThisConsp& >= MinUnits# AND ThisConsp& <= MaxUnits# THEN
                    MeterConsp# = Round#(MeterConsp# + ThisConsp&)
                  ELSE
                    CustTRCnt = CustTRCnt - 1
                  END IF
                  EXIT FOR
                END IF
              NEXT
            END IF
          END IF
          ThisTrans& = UBTransRec(1).PrevTrans
        LOOP
        IF CustTRCnt > 0 THEN
          PRINT #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; " ";
          PRINT #UBRpt, FUsing$(STR$(Cnt), "#####"); "  "; LEFT$(UBCustRec(1).CUSTNAME, 25);
          PRINT #UBRpt, TAB(44); LEFT$(UBCustRec(1).SERVADDR, 25);
          PRINT #UBRpt, TAB(70); FUsing$(STR$(MeterConsp#), "##########");
          PRINT #UBRpt, TAB(86); FUsing$(STR$(CustTRCnt), "####");
          IF CustTRCnt > 0 THEN
            AvgCon# = Round#(MeterConsp# / CustTRCnt)
          ELSE
            AvgCon# = 0
          END IF
          PRINT #UBRpt, TAB(93); FUsing$(STR$(AvgCon#), "##########")
          CustCnt = CustCnt + 1
          LineCnt = LineCnt + 1
          GTotalCon# = GTotalCon# + MeterConsp#
          GTotalTRCnt& = GTotalTRCnt& + CustTRCnt
          SELECT CASE UBCustRec(1).Status
          CASE "A"
            Active = Active + 1
          CASE "F"
            Final = Final + 1
          CASE "I"
            InActive = InActive + 1
          CASE "B"
            Balance = Balance + 1
          CASE ELSE
            UnKnown = UnKnown + 1
          END SELECT
        END IF
      END IF
      IF LineCnt > MaxLines THEN
        PRINT #UBRpt, CHR$(12)
        GOSUB CustByRateHeader
      END IF
    END IF
ActiveOnlySkip:
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
'    IF CustCnt > 100 THEN EXIT FOR
    ShowPctComp Cnt, NumOfCust&
  NEXT

  GOSUB CustByRateTotals

  ERASE UBCustRec    'free up memory
  
  CLOSE

  IF NOT AbortFlag THEN
    PrintRptFile "Consumption by Rate Code & Date.", "UBCSBYRT.RPT", 1, RetCode, EntryPoint
  END IF

oExitCustByRate:
  EXIT SUB

CustByRateHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Consumption by Rate Code & Date    "; "Date: "; DATE$; TAB(70); "Page: "; PageNo
  IF FromCnt = 1 THEN
    PRINT #UBRpt, "                                                                         Water        Trans"
  ELSE
    PRINT #UBRpt, "                                                                      Irrigation      Trans"
  END IF
  PRINT #UBRpt, "Location   Acct.  Customer Name             Service Address           Consumption     Count     Average"
  PRINT #UBRpt, Dash80$
  LineCnt = 4
RETURN

CustByRateTotals:
  'PageNo = PageNo + 1
  PageNo = PageNo + 1
  PRINT #UBRpt, CHR$(12)
  PRINT #UBRpt, "Consumption by Rate Code & Date    "; "Date: "; DATE$; TAB(70); "Page: "; PageNo
  PRINT #UBRpt, "Report Summary"
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt,
  PRINT #UBRpt, USING "     Active: ####"; Active
  PRINT #UBRpt, USING "      Final: ####"; Final
  PRINT #UBRpt, USING "   Inactive: ####"; InActive
  PRINT #UBRpt, USING "    Balance: ####"; Balance
  PRINT #UBRpt, USING "    Unknown: ####"; UnKnown
  PRINT #UBRpt, USING "      TOTAL: ####"; CustCnt
  PRINT #UBRpt,
  PRINT #UBRpt, USING "TOTAL CONSUMPTION: ############"; GTotalCon#
  PRINT #UBRpt, USING "TOTAL TRANS COUNT: ############"; GTotalTRCnt&
  IF GTotalTRCnt& > 0 THEN
    GTotAvg# = GTotalCon# / GTotalTRCnt&
  END IF
  PRINT #UBRpt, USING " COMBINED AVERAGE: #########.##"; GTotAvg#

  PRINT #UBRpt,
  PRINT #UBRpt, "     From Date: "; Num2Date(FromDate)
  PRINT #UBRpt, "     Thru Date: "; Num2Date(ThruDate)
  PRINT #UBRpt, "     RATE CODE: "; WhatRate$
  IF ActiveFlag THEN
    PRINT #UBRpt, "   Active Only: YES"
  ELSE
    PRINT #UBRpt, "   Active Only: NO"
  END IF
  IF MinUnits# = 0 THEN
    PRINT #UBRpt, " Minumum Units: NONE"
  ELSE
    PRINT #UBRpt, USING " Minimum Units: ######"; MinUnits#
  END IF
  IF MaxUnits# > 999999999998# THEN
    PRINT #UBRpt, " Maximum Units: NONE"
  ELSE
    PRINT #UBRpt, USING " Maximum Units: ######"; MaxUnits#
  END IF

  PRINT #UBRpt, "   Report Type: "; RptType$
  PRINT #UBRpt,
  PRINT #UBRpt, CHR$(12)
RETURN

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
  Fld(4).Protected = True
  
  NumOfRates = GetNumRateRecs%
  REDIM UBRateTbls(1 TO NumOfRates) AS UBRateTblRecType
  REDIM Choice$(0 TO NumOfRates + 10, 1)
  
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRates
    GET RateFile, Cnt, UBRateTbls(Cnt)
    Choice$(Cnt, 0) = UCASE$(UBRateTbls(Cnt).RATECODE + " " + UBRateTbls(Cnt).RATEDESC) + " " + STR$(Cnt)
  NEXT
  CLOSE
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  Choice$(0, 0) = "3"
  
  FirstTime = True
  
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      QPrintRC SPACE$(25), 14, 20, 15
    END IF
    
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
      RATECODE$ = QPTrim$(UBRateTbls(1).RATECODE)
      FOR Tnt = 1 TO 10
        IF UBRateTbls(Lnt).TblBreaks(Tnt).UNITS >= 0 THEN
          IF UBRateTbls(Lnt).TblBreaks(Tnt).UNITS > 0 THEN
            Greater = True
          END IF
          IF (UBRateTbls(Lnt).TblBreaks(Tnt).UNITS = 0) AND (Greater = True) THEN
            MaxStep = Tnt
            TblBreak&(Tnt) = 99999999
            EXIT FOR
          END IF
          TblBreak&(Tnt) = UBRateTbls(Lnt).TblBreaks(Tnt).UNITS
          MaxStep = Tnt
        ELSE
          MaxStep = Tnt
          TblBreak&(Tnt) = 99999999
          EXIT FOR
        END IF
      NEXT Tnt
    END IF
  NEXT Lnt
  CLOSE
  
  WhatRev = VAL(Form$(4, 0))
  
  'if WhatRev
  
  '  MType$ = RTRIM$(Form$(4, 0))
  '  IF MType$ = "Water Only" THEN MType = 1: Mtype1 = 3
  '  IF MType$ = "Sewer Only" THEN MType = 2: Mtype1 = 3
  '  IF MType$ = "Both Water/Sewe" THEN MType = 3: Mtype1 = 1
  '  IF MType$ = "Electric" THEN MType = 4: Mtype1 = 9
  '  IF MType$ = "Demand Elec" THEN MType = 5: Mtype1 = 9
  '  IF MType$ = "Gas" THEN MType = 6: Mtype1 = 9
  '  IF MType$ = "Touch Read" THEN MType = 7: Mtype1 = 7
  '  IF MType = 0 THEN
  '    MType = 7
  '  END IF
  
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
        
        'IF CustomerRecord = 10660 THEN STOP
        
        IF CustomerRecord > 0 THEN
          GET UBCust, CustomerRecord, UBCustRec(1)
          IF WhatRev > 0 THEN
            CCode$ = QPTrim$(UBCustRec(1).Serv(WhatRev).RATECODE)
            IF CCode$ = RATECODE$ THEN
              MtrType$ = UBCustRec(1).Serv(WhatRev).RMtrType
              SELECT CASE MtrType$
              CASE "W"
                MType = 1
              CASE "S"
                MType = 2
              CASE "C"
                MType = 3
              CASE ELSE
                MType = 4
              END SELECT
              ValidCustomer = 1
            END IF
          ELSE
            FOR Snt = 1 TO 15
              IF QPTrim$(UBCustRec(1).Serv(Snt).RATECODE) = RATECODE$ THEN
                MtrType$ = UBCustRec(1).Serv(Snt).RMtrType
                SELECT CASE MtrType$
                CASE "W"
                  MType = 1
                CASE "S"
                  MType = 2
                CASE "C"
                  MType = 3
                CASE ELSE
                  MType = 4
                END SELECT
                ValidCustomer = 1
                EXIT FOR
              END IF
            NEXT Snt
          END IF
        END IF
        'IF Customer Has a Rate Code Match then Analyze Meter Type to Get Usage
        IF ValidCustomer = 1 THEN
          Multi& = 0
          Cubic = False
          FOR MCnt = 1 TO 7
            MType$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MtrType)
            IF LEN(MType$) > 0 THEN
              Multi& = UBCustRec(1).LocMeters(MCnt).MtrMulti
              IF UBCustRec(1).LocMeters(MCnt).MTRUnit = "C" THEN
                Cubic = True
              END IF
              EXIT FOR
            END IF
          NEXT
          IF Multi& <= 0 THEN Multi& = 1
          IF WhatRev > 0 THEN
            ChkMtr = True
          ELSE
            ChkMtr = False
          END IF
          FOR MCnt = 1 TO 7
            IF ChkMtr = True THEN
              IF UBTransRec(1).MtrTypes(MCnt) <> MType THEN
                GOTO SkipThisMtr
              END IF
            END IF
            MeterConsp& = MeterConsp& + UBTransRec(1).CurRead(MCnt) - UBTransRec(1).PrevRead(MCnt)
            IF MeterConsp& < 0 THEN
              MaxMeterAmt& = 10& ^ (LEN(STR$(UBTransRec(1).PrevRead(MCnt))) - 1)
              MeterConsp& = (MaxMeterAmt& - UBTransRec(1).PrevRead(MCnt)) + UBTransRec(1).CurRead(MCnt)
            END IF
            IF Cubic THEN
              MeterConsp& = MeterConsp& * 7.481
            END IF
            MeterConsp& = MeterConsp& * Multi&
            'IF MeterConsp& = 1 THEN STOP
            TMeterConsp# = TMeterConsp# + MeterConsp&
            GTMeterConsp# = GTMeterConsp# + MeterConsp&
            MeterConsp& = 0
            'END IF
SkipThisMtr:
          NEXT MCnt
          
        END IF
        
        
        IF (TMeterConsp# >= 0) AND (ValidCustomer = 1) THEN
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
  PRINT #UBRpt, "Grand Total Consumption:"; USING "############"; GTMeterConsp#
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
  
  REDIM CustCnt(1 TO NumOfRateRecs, 1 TO 2) AS INTEGER
  
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
    IF UBCustRec(1).Status = "A" THEN
      '*************************************
      '   Main body of Printing goes here
      FOR RCnt = 1 TO 15
        IF RCnt = 5 THEN GOTO SkipThisRev
        ThisCode$ = QPTrim$(UBCustRec(1).Serv(RCnt).RATECODE)
        IF LEN(ThisCode$) > 0 THEN
          FOR ZCnt = 1 TO NumOfRateRecs
            IF QPTrim$(UBRateTblRec(ZCnt).RATECODE) = ThisCode$ THEN
              CustCnt(ZCnt, 1) = CustCnt(ZCnt, 1) + 1
              EXIT FOR
            END IF
          NEXT
        END IF
SkipThisRev:
      NEXT
      
      ThisCode$ = QPTrim$(UBCustRec(1).Serv(5).RATECODE)
      IF LEN(ThisCode$) > 0 THEN
        FOR ZCnt = 1 TO NumOfRateRecs
          IF QPTrim$(UBRateTblRec(ZCnt).RATECODE) = ThisCode$ THEN
            CustCnt(ZCnt, 2) = CustCnt(ZCnt, 2) + 1
            EXIT FOR
          END IF
        NEXT
      END IF
      
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
    PRINT #UBRpt, UBRateTblRec(Cnt).RATECODE; TAB(12); UBRateTblRec(Cnt).RATEDESC; TAB(51); USING "####"; CustCnt(Cnt, 1); TAB(68); CustCnt(Cnt, 2)
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
  PRINT #UBRpt, "Rate        Description                    Water/Sewer       Irrigation"
  PRINT #UBRpt, Dash80$
  LineCnt = 6
  RETURN
  
DOCustRateFooter:
  PRINT #UBRpt, ""
  PRINT #UBRpt, "Total Customers Count: "; USING "####,#"; CustomerCnt
  RETURN
  
  
END SUB

SUB IrrConSteps
  
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
  
  Fld(4).Protected = True
  
  NumOfRates = GetNumRateRecs%
  REDIM UBRateTbls(1 TO NumOfRates) AS UBRateTblRecType
  REDIM Choice$(0 TO NumOfRates + 10, 1)
  
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRates
    GET RateFile, Cnt, UBRateTbls(Cnt)
    Choice$(Cnt, 0) = UCASE$(UBRateTbls(Cnt).RATECODE + " " + UBRateTbls(Cnt).RATEDESC) + " " + STR$(Cnt)
  NEXT
  CLOSE
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  Choice$(0, 0) = "3"
  Choice$(0, 1) = "4"
  
  FirstTime = True
  
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayUBScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      QPrintRC "Irrigation Consumption Report   ", 5, 23, 14
      QPrintRC SPACE$(25), 14, 20, 15
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
      
    CASE F10Key
      GOSUB IConChkFields:
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
  IF AbortFlag THEN GOTO IExitConRpt
  
  UBRateTblRecLen = LEN(UBRateTbls(1))
  
  RateFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS RateFile LEN = UBRateTblRecLen
  FOR Lnt = 1 TO NumOfRates
    WhatRate$ = QPTrim$(LEFT$(Form$(3, 0), 4))
    ThisRate$ = QPTrim$(LEFT$(Choice$(Lnt, 0), 4))
    IF ThisRate$ = WhatRate$ THEN
      GET RateFile, Lnt, UBRateTbls(1)
      RATECODE$ = QPTrim$(UBRateTbls(1).RATECODE)
      FOR Tnt = 1 TO 10
        IF UBRateTbls(Lnt).TblBreaks(Tnt).UNITS >= 0 THEN
          IF UBRateTbls(Lnt).TblBreaks(Tnt).UNITS > 0 THEN
            Greater = True
          END IF
          IF (UBRateTbls(Lnt).TblBreaks(Tnt).UNITS = 0) AND (Greater = True) THEN
            MaxStep = Tnt
            TblBreak&(Tnt) = 99999999
            EXIT FOR
          END IF
          TblBreak&(Tnt) = UBRateTbls(Lnt).TblBreaks(Tnt).UNITS
          MaxStep = Tnt
        ELSE
          MaxStep = Tnt
          TblBreak&(Tnt) = 99999999
          EXIT FOR
        END IF
      NEXT Tnt
    END IF
  NEXT Lnt
  CLOSE
  
  WhatRev = 5   'VAL(Form$(4, 0))
  
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
  ShowProcessingScrn "Irrigation Consumption Report"
  
  GOSUB IDoUnitStepHeader
  
  FOR Cnt! = 1 TO NumOfRecords!
    GET UBTrans, Cnt!, UBTransRec(1)
    IF UBTransRec(1).TransDate >= BegDate AND UBTransRec(1).TransDate <= Enddate THEN
      IF UBTransRec(1).TransType = 1 OR UBTransRec(1).TransType = 101 THEN
        'Set Valid Flag to Off to Check Each Customer For Valid Rate Code
        ValidCustomer = 0
        IF LineCnt > MaxLines THEN
          PRINT #UBRpt, CHR$(12)
          GOSUB IDoUnitStepHeader
        END IF
        
        'IF Valid Customer Then Get the Customer and Check to See if Rate Applies
        CustomerRecord = UBTransRec(1).CustAcctNo
        IF CustomerRecord > 0 THEN
          GET UBCust, CustomerRecord, UBCustRec(1)
          CCode$ = QPTrim$(UBCustRec(1).Serv(WhatRev).RATECODE)
          IF CCode$ = RATECODE$ THEN
            MtrType$ = UBCustRec(1).Serv(WhatRev).RMtrType
            SELECT CASE MtrType$
            CASE "W"
              MType = 1
            CASE "S"
              MType = 2
            CASE "C"
              MType = 3
            CASE ELSE
              MType = 4
            END SELECT
            ValidCustomer = 1
          END IF
        END IF
        'IF Customer Has a Rate Code Match then Analyze Meter Type to Get Usage
        IF ValidCustomer = 1 THEN
          Multi& = 0
          Cubic = False
          FOR MCnt = 1 TO 7
            MType$ = QPTrim$(UBCustRec(1).LocMeters(MCnt).MtrType)
            IF LEN(MType$) > 0 THEN
              Multi& = UBCustRec(1).LocMeters(MCnt).MtrMulti
              IF UBCustRec(1).LocMeters(MCnt).MTRUnit = "C" THEN
                Cubic = True
              END IF
              EXIT FOR
            END IF
          NEXT
          IF Multi& <= 0 THEN Multi& = 1
          FOR MCnt = 1 TO 7
            IF UBTransRec(1).MtrTypes(MCnt) <> MType THEN
              GOTO ISkipThisMtr
            END IF
            MeterConsp& = MeterConsp& + UBTransRec(1).CurRead(MCnt) - UBTransRec(1).PrevRead(MCnt)
            IF MeterConsp& < 0 THEN
              MaxMeterAmt& = 10& ^ (LEN(STR$(UBTransRec(1).PrevRead(MCnt))) - 1)
              MeterConsp& = (MaxMeterAmt& - UBTransRec(1).PrevRead(MCnt)) + UBTransRec(1).CurRead(MCnt)
            END IF
            IF Cubic THEN
              MeterConsp& = MeterConsp& * 7.481
            END IF
            MeterConsp& = MeterConsp& * Multi&
            TMeterConsp# = TMeterConsp# + MeterConsp&
            GTMeterConsp# = GTMeterConsp# + MeterConsp&
            MeterConsp& = 0
ISkipThisMtr:
          NEXT MCnt
        END IF
        
        IF (TMeterConsp# >= 0) AND (ValidCustomer = 1) THEN
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
  
  GOSUB IDoUnitStepFooter:
  
  CLOSE
  
  ERASE TblBreak&, TotalConsp#, TotalCust
  
  IF NOT AbortFlag THEN
    PrintRptFile "Consumption by Step", "UBBKCNSP.RPT", 1, RetCode, EntryPoint
  END IF
  
  'KillFile "UBBKCNSP.RPT"
  EXIT SUB
  
IDoUnitStepHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, TAB(TabStop); TownName$
  PRINT #UBRpt, TAB(24); "Irrigation Consumption by RateCode"; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, ""
  PRINT #UBRpt, "    For Rate Code: "; Form$(3, 0)
  PRINT #UBRpt, "Period Beginning : "; Form$(1, 0)
  PRINT #UBRpt, "    Period Ending: "; Form$(2, 0)
  PRINT #UBRpt, "      Report Date: "; DATE$
  PRINT #UBRpt, ""
  PRINT #UBRpt, Dash80$
  LineCnt = 6
  RETURN
  
IDoUnitStepFooter:
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
  PRINT #UBRpt, "Grand Total Consumption:"; USING "############"; GTMeterConsp#
  PRINT #UBRpt, CHR$(12);
  RETURN
  GOTO IExitConRpt
  
IConChkFields:
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
  
IExitConRpt:
  CLOSE
  
END SUB

SUB TotalUsageByDate
  
  SHARED Choice$()
  REDIM Book(99), TotalConsp#(99), TotalCust(99)
  
  REDIM TempScrn(0)
  MaxLines = 56
  Dash80$ = STRING$(80, "-")
  
  LibName$ = "UB"
  ScrnName$ = "BKCONRTP"
  
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
  Fld(3).Protected = True
  Fld(4).Protected = True
  'REDIM Choice$(0 TO NumOfRates + 10, 1)

  FirstTime = True
  '--Set screen number to one and display screen
  BlockClear
  DisplayUBScrn ScrnName$

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      QPrintRC " Total Consumption Report  ", 5, 23, 14
      QPrintRC SPACE$(25), 12, 20, 15
      QPrintRC SPACE$(25), 14, 20, 15
    END IF

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
  ShowProcessingScrn "Total Consumption Report"
  
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
            MTRType1$ = QPTrim$(UBCustRec(1).Serv(1).RMtrType)
            MTRType2$ = QPTrim$(UBCustRec(1).Serv(2).RMtrType)
            MTRType5$ = QPTrim$(UBCustRec(1).Serv(5).RMtrType)

            IF LEN(MTRType1$) > 0 THEN
              SELECT CASE MTRType1$
              CASE "W"
                WMtrChk = 1
              CASE "C"
                WMtrChk = 3
              END SELECT
            END IF
            
            IF LEN(MTRType2$) > 0 THEN
              SELECT CASE MTRType2$
              CASE "S"
                SMtrChk = 2
              CASE "C"
                SMtrChk = 3
              END SELECT
            END IF
            
            IF QPTrim$(UBCustRec(1).Serv(Snt).RATECODE) = RATECODE$ THEN
              MtrType$ = UBCustRec(1).Serv(Snt).RMtrType
              ValidCustomer = 1
              EXIT FOR
            END IF
          'NEXT Snt
        END IF
        
        'IF Customer Has a Rate Code Match then Analyze Meter Type to Get Usage
        IF ValidCustomer = 1 THEN
          FOR MCnt = 1 TO 7
            IF (UBTransRec(1).MtrTypes(MCnt) = MType OR UBTransRec(1).MtrTypes(MCnt) = Mtype1) OR MType = 0 THEN
              ' Get Usage Here
              MeterConsp& = MeterConsp& + UBTransRec(1).CurRead(MCnt) - UBTransRec(1).PrevRead(MCnt)
              IF MeterConsp& < 0 THEN
                MaxMeterAmt& = 10& ^ (LEN(STR$(UBTransRec(1).PrevRead(MCnt))) - 1)
                MeterConsp& = (MaxMeterAmt& - UBTransRec(1).PrevRead(MCnt)) + UBTransRec(1).CurRead(MCnt)
              END IF
              
              'MeterConsp& = MeterConsp& * 100
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
  PRINT #UBRpt, TAB(29); "Total Consumption Report."; TAB(70); "Page #"; PageNo
  PRINT #UBRpt, ""
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

