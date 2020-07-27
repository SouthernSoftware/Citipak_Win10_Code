DEFINT A-Z
DECLARE SUB HideCursor ()
DECLARE SUB UBQuickCustRptMenu ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB FGetA (Handle%, SEG Dest AS ANY, NumBytes%)
DECLARE FUNCTION Round# (DblNum#)
'DECLARE SUB SmallPause ()
DECLARE SUB SelectNoteText (NoteText$)
DECLARE FUNCTION CustHasMsg% (RecNo&)
DECLARE SUB PrintTransHist (DetailFlag)
DECLARE SUB PrintConsumpHist ()
DECLARE SUB EditOwnerInfo (RecNo&, OTempName$)
DECLARE SUB ReIndexSystem (PromptFlag%)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Chk4DupeBookSeqNum (Book$, SeqNum$)
DECLARE FUNCTION Compare3% (BYVAL Seg1, BYVAL Adr1, BYVAL Seg2, BYVAL Adr2, NumBytes)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION FmtBook$ (Book$)
DECLARE FUNCTION FmtSeqN$ (SeqN$)
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (Number$)

DECLARE SUB AddEditCustomer (RecNo&)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB CustMessageSystem (RecNo&)
DECLARE SUB DeleteCust ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB EditCustomer ()
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB InsertT (SEG StartElement AS ANY, ElSize%, NumEls%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB PrintCustByRate ()
DECLARE SUB PrintCustList (LocFlag%)
DECLARE SUB PrintLocaList (LocFlag)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveNewCustRec (Form$(), RecNo&)
DECLARE SUB SaveOldCustRec (Form$(), RecNo&)
DECLARE SUB SaveScrn (Array%())
DECLARE SUB SetCustFinal ()
DECLARE SUB ShowCursor ()
DECLARE SUB ShowCustConsHist (CustRec&)
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowWrkOrdHistory (CustRec&)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB StuffBuf (Ky$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB UBLog (Text$)
DECLARE SUB WaitForAction ()
DECLARE SUB MPaintBox (TRow%, LCol%, BRow%, RCol%, PColor%)
DECLARE FUNCTION ChkBillFile% ()
DECLARE FUNCTION GetCustMeterType% (UBCustRec() AS ANY, ThisMeter%)
DECLARE SUB MakeSequenceIndex (Text$)


  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'PageInfo.BI'                  'Form Page info
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'ubTrans.BI'
  ''$INCLUDE: 'ubowner.BI'

  CONST False = 0, True = NOT False

  CrLf$ = CHR$(13) + CHR$(10)

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 8)

  MChoice$(1) = " Customer Consumption History"
  MChoice$(2) = " Customer Transaction History Summary"
  MChoice$(3) = " Customer Transaction History Detailed"
  MChoice$(4) = " Quick Customer Listing by Name"
  MChoice$(5) = " Quick Customer Listing by Location"
  MChoice$(6) = " Quick Customer Listing by Sequence No."
  MChoice$(7) = " Quick Customer Listing by Rate Code"
  MChoice$(8) = " Exit to OS"

  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 20
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

    TitleBox 2, Col, MaxLen + 3, "Quick Customer Reports Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    CASE 1
      PrintConsumpHist
    CASE 2
      PrintTransHist False
    CASE 3
      PrintTransHist True
    CASE 4
      PrintLocaList False
    CASE 5
      PrintLocaList True
    CASE 6
      PrintLocaList 1
    CASE 7
      PrintCustByRate
    CASE 8
      HideCursor
      END
    END SELECT
  LOOP

'  UBQuickCustRptMenu
  RUN "UBMENU"

SUB PrintConsumpHist

  BlockClear
  LookUp RecNo&, "Customer Consumption", 2, True, False

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  IF INSTR(UBSetUpRec(1).UTILNAME, "TROY") > 0 THEN
    TroyFlag = True
  END IF

  IF RecNo& = 0 THEN
    GOTO ExitConsumpHist
  END IF

  Dash80$ = STRING$(80, "-")

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET #UBCust, RecNo&, UBCustRec(1)
  CLOSE UBCust

'  Cubic = False
'  FOR MTCnt = 1 TO 7
'    IF UBCustRec(1).LocMeters(MTCnt).MtrUnit = "C" THEN
'      Cubic = True
'      EXIT FOR
'    END IF
'  NEXT

  UBRpt = FREEFILE
  OPEN "UBCONSMP.RPT" FOR OUTPUT AS UBRpt

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen

  BlockClear

  DisplayUBScrn "PRORPT"
  QPrintRC "Processing Consumption History.", 11, 26, -1

  GOSUB DoConsRptHeader

  ThisTrans& = UBCustRec(1).LastTrans

  DO WHILE ThisTrans& > 0
    GET #UBTran, ThisTrans&, UBTranRec(1)
    IF UBTranRec(1).TransType = TranUtilityBill THEN
      GOSUB PrintConsDetail
      DidCnt = DidCnt + 1
      IF DidCnt = 12 THEN
        EXIT DO
      END IF
    END IF
    ThisTrans& = UBTranRec(1).PrevTrans
  LOOP
  GOSUB DoConsFooter

  CLOSE

  IF NOT AbortFlag THEN
    PrintRptFile "Customer Consumption Report.", "UBCONSMP.RPT", 1, RetCode, EntryPoint
  END IF


ExitConsumpHist:
EXIT SUB

PrintConsDetail:
  
  DidAMeter = False
  EstFlag = False
  FOR EstCnt = 1 TO 7
    IF UBTranRec(1).EstRead(EstCnt) = "Y" THEN
      EstFlag = True
      EXIT FOR
    END IF
  NEXT
  FOR MtrCnt = 1 TO 7
    MTRMulti# = 0
    FOR MCnt = 1 TO 7
      IF UBTranRec(1).MtrTypes(MtrCnt) = GetCustMeterType%(UBCustRec(), MCnt) THEN
        MTRMulti# = UBCustRec(1).LocMeters(MCnt).MTRMulti
        EXIT FOR
      END IF
    NEXT
    IF MTRMulti# = 0 THEN
      IF TroyFlag THEN
        MTRMulti# = 100
      ELSE
        MTRMulti# = 1
      END IF
    END IF

    IF UBTranRec(1).MtrTypes(MtrCnt) <> 0 THEN
      DidAMeter = True
      SELECT CASE UBTranRec(1).MtrTypes(MtrCnt)
      CASE MtrWaterOnly
        MeterType$ = "      Water"
        F$ = "W"
      CASE MtrSewerOnly
        MeterType$ = "      Sewer"
        F$ = "S"
      CASE MtrCombined
        MeterType$ = "Water/Sewer"
        F$ = "C"
      CASE MtrElectric
        MeterType$ = "   Electric"
        F$ = "E"
      CASE MtrDemand
        MeterType$ = " D Electric"
        F$ = "D"
      CASE MtrGas
        MeterType$ = "  Gas Meter"
        F$ = "G"
      CASE MtrTouchRead
        MeterType$ = " Touch Read"
        F$ = "T"
      CASE MtrLightsService
        MeterType$ = "  L Service"
      CASE ELSE
        MeterType$ = "  ?????????"
      END SELECT
      FOR CCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(CCnt).MTRType = F$ THEN
          IF UBCustRec(1).LocMeters(CCnt).MTRUnit = "C" THEN
            CubMeter = True
          ELSE
            CubMeter = False
          END IF
          EXIT FOR
        END IF
      NEXT
      GOSUB PrintThisMeter
    END IF
  NEXT
  IF NOT DidAMeter THEN
    MeterType$ = "        "
    MtrCnt = 1
    GOSUB PrintThisMeter
  END IF

RETURN

PrintThisMeter:

  PRINT #UBRpt, Num2Date(UBTranRec(1).TransDate);
  IF EstFlag THEN
    PRINT #UBRpt, "*E";
  END IF
  PRINT #UBRpt, TAB(19); MeterType$;
  PRINT #UBRpt, TAB(34); USING "##########"; UBTranRec(1).CurRead(MtrCnt);
  PRINT #UBRpt, TAB(46); USING "##########"; UBTranRec(1).PrevRead(MtrCnt);
  MeterConsp# = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
  IF MeterConsp# < 0 THEN
    MaxMeterAmt& = 10& ^ (LEN(STR$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
    MeterConsp# = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
  END IF
  MeterConsp# = MeterConsp# * MTRMulti#
  IF CubMeter THEN
    MeterConsp# = MeterConsp# * 7.481
  END IF
  PRINT #UBRpt, TAB(56); USING "##########"; MeterConsp#;
  IF UBTranRec(1).ReadDate <= 0 THEN
    PRINT #UBRpt, "     ??-??-????"
  ELSE
    PRINT #UBRpt, "     "; Num2Date$(UBTranRec(1).ReadDate) '; "!"; UBTranRec(1).EstRead(MtrCnt); "!"
  END IF

  TotalConsp# = TotalConsp# + MeterConsp#

RETURN

DoConsRptHeader:
  PRINT #UBRpt, TAB(28); "Consumption History Report. "
  PRINT #UBRpt,
  PRINT #UBRpt, "Customer: "; UBCustRec(1).CustName; TAB(57); "Report Date: "; DATE$
  PRINT #UBRpt,
  PRINT #UBRpt, "Transaction                         Current   Previous"
  PRINT #UBRpt, "   Date            Meter Type       Reading    Reading       Usage    ReadDate"
  PRINT #UBRpt, Dash80$
RETURN

DoConsFooter:
  IF DidCnt > 0 THEN
    PRINT #UBRpt, Dash80$
    PRINT #UBRpt, USING "Average Consumption: #########"; TotalConsp# / DidCnt
  ELSE
    PRINT #UBRpt, "NO TRANSACTIONS!!!"
    PRINT #UBRpt, Dash80$
  END IF
RETURN
END SUB

SUB PrintCustByRate

  SHARED Choice$()

  MaxLines = 59
  PageNo = 0
  Dash80$ = STRING$(80, "-")

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen
  IF INSTR(UCASE$(UBSetUpRec(1).UTILNAME), "CAROLINA BEACH") > 0 THEN
    CBeach = True
  END IF

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  IdxRecLen = 4 'we are using a long integer
  IdxFileSize& = FileSize("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen

  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it

'*****************
  LibName$ = "UB"
  ScrnName$ = "CUSBYRAT"

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
    GOTO ExitCustByRate
  END IF

  REDIM Choice$(0 TO NumOfRateRecs, 0 TO 1)
  Choice$(0, 0) = "1"
  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  FOR Cnt = 1 TO NumOfRateRecs
    GET UBFile, Cnt, UBRateTblRec(1)
    Choice$(Cnt, 0) = SPACE$(30)
    LSET Choice$(Cnt, 0) = UBRateTblRec(1).RATECODE
    MID$(Choice$(Cnt, 0), 6, 24) = UBRateTblRec(1).RATEDESC
  NEXT
  CLOSE
  'SortStr2 VARPTR(Choice$(1)), Cnt - 1, 0

  RateRec = 0
  PrnRate = False

  Action = 1
  FirstTime = True

  BlockClear

  DisplayUBScrn "CUSBYRAT"

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      LSET Form$(2, 0) = "N"
      LSET Form$(3, 0) = "Y"
      Action = 2
    END IF
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key 'F10
      WhatRate$ = QPTrim$(LEFT$(Form$(1, 0), 4))
      IF LEN(WhatRate$) > 0 THEN
        IF QPTrim$(Form$(2, 0)) = "Y" THEN
          PrnRate = True
          FOR Cnt = 1 TO NumOfRateRecs
            ThisRate$ = QPTrim$(LEFT$(Choice$(Cnt, 0), 4))
            IF ThisRate$ = WhatRate$ THEN
              RateRec = Cnt
              EXIT FOR
            END IF
          NEXT
        END IF
        SELECT CASE Form$(3, 0)
        CASE "Y"
          ActOnly = True
        END SELECT
        OKFlag = True

      END IF
    CASE EscKey
      ExitFlag = True
    END SELECT
    
    '--Check screen page
  LOOP UNTIL ExitFlag OR OKFlag

  IF ExitFlag THEN
    GOTO ExitCustByRate
  END IF
  
'****************

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBRpt = FREEFILE
  OPEN "UBCSBYRT.RPT" FOR OUTPUT AS UBRpt

  BlockClear
  ShowProcessingScrn "Customer Listing By Rate Code."

  GOSUB CustByRateHeader

  FOR Cnt = 1 TO IdxNumOfRecs
    GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
    IF UBCustRec(1).DelFlag = 0 THEN
      IF ActOnly THEN
        IF UBCustRec(1).Status <> "A" THEN
          GOTO SkipCustRate
        END IF
      END IF
      FOR SCnt = 1 TO 15
        ThisRate$ = QPTrim$(UBCustRec(1).Serv(SCnt).RATECODE)
        IF WhatRate$ = ThisRate$ THEN
          'Mtr$ = "   MTR NO: "
          Mtr$ = ""
          PRINT #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; " "; FUsing$(STR$(IdxBuff(Cnt).RecNum), "#####"); "  "; LEFT$(UBCustRec(1).CustName, 25); "  "; LEFT$(UBCustRec(1).SERVADDR, 30); " "; UBCustRec(1).Status
          CustCnt = CustCnt + 1
          IF CBeach THEN
            'Mtr$=
            FOR MtrCnt = 1 TO 7
              MTRNUM$ = QPTrim$(UBCustRec(1).LocMeters(MtrCnt).MTRNUM)
              IF LEN(MTRNUM$) > 0 THEN
                Mtr$ = Mtr$ + MTRNUM$ + " "
              END IF
            NEXT

            PRINT #UBRpt, Mtr$
            LineCnt = LineCnt + 2
          ELSE
            LineCnt = LineCnt + 1
          END IF


'IF SCnt = 4 AND UBCustRec(1).TAXEXPT = "Y" THEN
'  LPRINT IdxBuff(Cnt).RecNum
'END IF
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
          EXIT FOR
        END IF
      NEXT

      IF LineCnt > MaxLines THEN
        PRINT #UBRpt, CHR$(12)
        GOSUB CustByRateHeader
      END IF
    END IF
SkipCustRate:
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp Cnt, IdxNumOfRecs
  NEXT

  GOSUB CustByRateTotals

  ERASE IdxBuff, UBCustRec   'free up memory

  IF PrnRate THEN
    GOSUB PrintRateCode
  END IF

  CLOSE UBCust, UBRpt

  IF NOT AbortFlag THEN
    PrintRptFile "Customer Listing by Rate Code.", "UBCSBYRT.RPT", 1, RetCode, EntryPoint
  END IF

ExitCustByRate:
  EXIT SUB

CustByRateHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Customer Listing By Rate Code      "; "Date: "; DATE$; TAB(70); "Page: "; PageNo
  PRINT #UBRpt, "RATE CODE: "; WhatRate$
  PRINT #UBRpt, "Location   Acct.  Customer Name             Service Address             Status"
  PRINT #UBRpt, "Meter No:"
  PRINT #UBRpt, Dash80$
  LineCnt = 5
RETURN

CustByRateTotals:
  'PageNo = PageNo + 1
  PRINT #UBRpt,
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Customer Summary"
  PRINT #UBRpt,
  PRINT #UBRpt, USING "  Active: ####"; Active
  PRINT #UBRpt, USING "   Final: ####"; Final
  PRINT #UBRpt, USING "Inactive: ####"; InActive
  PRINT #UBRpt, USING " Balance: ####"; Balance
  PRINT #UBRpt, USING " Unknown: ####"; UnKnown
  PRINT #UBRpt, USING " Deleted: ####"; DeletedCnt
  PRINT #UBRpt,
  PRINT #UBRpt, USING "   TOTAL: ####"; CustCnt
  PRINT #UBRpt, CHR$(12)
RETURN

PrintRateCode:

  REDIM UBRateTblRec(1) AS UBRateTblRecType
  UBRateTblRecLen = LEN(UBRateTblRec(1))

  NumRateRecs = FileSize("UBRATE.DAT") \ UBRateTblRecLen

  IF NumRateRecs = 0 THEN
    GOTO PrintRateExit
  END IF

  UBFile = FREEFILE
  OPEN "UBRATE.DAT" FOR RANDOM SHARED AS UBFile LEN = UBRateTblRecLen
  GET UBFile, RateRec, UBRateTblRec(1)
  CLOSE UBFile

  REDIM StepText(1 TO 10) AS STRING * 40

  GOSUB PrintRateHeader
    PRINT #UBRpt, "       Rate Code:  "; UBRateTblRec(1).RATECODE
    PRINT #UBRpt, "     Description:  "; UBRateTblRec(1).RATEDESC
    PRINT #UBRpt, "  Minimum Charge:"; FUsing$(STR$(UBRateTblRec(1).MINAMT), "#######.##")
    PRINT #UBRpt, "   Minimum Units:"; FUsing$(STR$(UBRateTblRec(1).MINUNITS), "##########")
    PRINT #UBRpt, "      Max Amount:"; FUsing$(STR$(UBRateTblRec(1).MaxAmt), "######.##")
    PRINT #UBRpt, "      [ Step ]        [ Beg Unit ]     [ Amount/Unit ]"
    FOR RCnt = 1 TO 10
      LSET StepText$(RCnt) = ""
      IF UBRateTblRec(1).TblBreaks(RCnt).UNITS >= 0 THEN
        MID$(StepText$(RCnt), 8) = FUsing$(STR$(UBRateTblRec(1).TblBreaks(RCnt).UNITS), "########")
      END IF
      IF UBRateTblRec(1).TblBreaks(RCnt).UNITAMT >= 0 THEN
        MID$(StepText$(RCnt), 25) = FUsing$(STR$(UBRateTblRec(1).TblBreaks(RCnt).UNITAMT), "####.######")
      END IF
    NEXT
    PRINT #UBRpt, "     First Break:"; StepText$(1)
    PRINT #UBRpt, "    Second Break:"; StepText$(2)
    PRINT #UBRpt, "     Third Break:"; StepText$(3)
    PRINT #UBRpt, "    Fourth Break:"; StepText$(4)
    PRINT #UBRpt, "     Fifth Break:"; StepText$(5)
    PRINT #UBRpt, "     Sixth Break:"; StepText$(6)
    PRINT #UBRpt, "   Seventh Break:"; StepText$(7)
    PRINT #UBRpt, "    Eighth Break:"; StepText$(8)
    PRINT #UBRpt, "     Ninth Break:"; StepText$(9)
    PRINT #UBRpt, "        All Over:"; StepText$(10)
    PRINT #UBRpt,
    PRINT #UBRpt, Dash80$
    PRINT #UBRpt, CHR$(12)

  ERASE UBRateTblRec, StepText

  GOTO PrintRateExit

PrintRateHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Utility Rate Table Listing."
  PRINT #UBRpt, "RATE CODE: "; WhatRate$; TAB(70); "Page:"; PageNo
  PRINT #UBRpt, "Report Date: "; DATE$
  PRINT #UBRpt, Dash80$
  NumPrinted = 0
RETURN

PrintRateExit:

RETURN

END SUB

SUB PrintLocaList (LocFlag)
  
  MaxLines = 59
  PageNo = 0
  Dash80$ = STRING$(80, "-")
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  REDIM UBCustBlank(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  IdxRecLen = 4 'we are using a long integer

  SELECT CASE LocFlag
  CASE 1
    MakeSequenceIndex "Sequence Number"
    IdxName$ = TempIndexName
    Title$ = "Quick Customer Listing by Sequence No."
  CASE True
    IdxName$ = "UBCUSTBK.IDX"
    Title$ = "Quick Customer Listing by Location."
  CASE False
    IdxName$ = "UBCUSTNM.IDX"
    Title$ = "Quick Customer Listing by Name."
  END SELECT

  IdxFileSize& = FileSize(IdxName$)
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH IdxName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBLOLIST.RPT" FOR OUTPUT AS UBRpt
  
  BlockClear
  ShowProcessingScrn Title$
  GOSUB DoLocaRptHeader
  
  FOR Cnt = 1 TO IdxNumOfRecs
    GET UBCust, IdxBuff(Cnt).RecNum, UBCustRec(1)
    'IF (LEN(QPTrim$(UBCustRec(1).PumpCode)) > 0) THEN 'OR UBCustRec(1).book <> "02" THEN
    '  GOTO skipEmbubba
    'END IF
    'GET UBCust, Cnt, UBCustRec(1)
    IF LineCnt > MaxLines THEN
      PRINT #UBRpt, CHR$(12)
      GOSUB DoLocaRptHeader
    END IF
    IF UBCustRec(1).DelFlag = 0 THEN
    'CTy$ = QPTrim$(UBCustRec(1).CustType)
    'IF CTy$ <> "R" THEN
      PRINT #UBRpt, UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB; " ";
      'PRINT #UBRpt, FUsing$(STR$(Cnt), "#####"); "  ";
      PRINT #UBRpt, FUsing$(STR$(IdxBuff(Cnt).RecNum), "#####"); "  ";
      PRINT #UBRpt, LEFT$(UBCustRec(1).CustName, 25); "  ";
      'PRINT #UBRpt, UBCustRec(1).Status; "     "; UBCustRec(1).CustType
      IF LocFlag = 1 THEN
        PRINT #UBRpt, LEFT$(UBCustRec(1).SERVADDR, 24); " ";
        PRINT #UBRpt, UBCustRec(1).Status; USING "  ########"; UBCustRec(1).SEQ
        'PRINT #UBRpt, UBCustRec(1).Status; USING "  ########"; UBCustRec(1).BillCycl
      ELSE
        PRINT #UBRpt, LEFT$(UBCustRec(1).SERVADDR, 25); " ";
        PRINT #UBRpt, UBCustRec(1).Status; "     "; UBCustRec(1).ZONE
      END IF
      CustCnt = CustCnt + 1
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
      LineCnt = LineCnt + 1
    ELSE
      DeletedCnt = DeletedCnt + 1
    END IF
    
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF

    ShowPctComp Cnt, IdxNumOfRecs
skipEmbubba:
  NEXT
  
  GOSUB DoLocaRptTotals
  
  CLOSE UBCust, UBLoca, UBRpt
  
  ERASE IdxBuff, UBCustRec
  
  IF NOT AbortFlag THEN
    PrintRptFile Title$, "UBLOLIST.RPT", 1, RetCode, EntryPoint
  END IF
  
  'KillFile "UBLOLIST.RPT"
  
EXIT SUB
  
DoLocaRptHeader:
  PageNo = PageNo + 1
  PRINT #UBRpt, "Customer Listing Report      "; "Date: "; DATE$; TAB(70); "Page: "; PageNo
  IF LocFlag = 1 THEN
    PRINT #UBRpt, "           Acct                                                    Acct    SEQ."
    PRINT #UBRpt, "Location    No.  Customer Name              Service Address        Status  Numb"
  ELSE
    PRINT #UBRpt, "           Acct                                                    Acct    Post"
    PRINT #UBRpt, "Location    No.  Customer Name              Service Address        Status  Route"
  END IF
  PRINT #UBRpt, Dash80$
  LineCnt = 4
RETURN
  
DoLocaRptTotals:
  PageNo = PageNo + 1
  PRINT #UBRpt,
  PRINT #UBRpt, Dash80$
  PRINT #UBRpt, "Customer Summary"
  PRINT #UBRpt,
  PRINT #UBRpt, USING "  Active: #####"; Active
  PRINT #UBRpt, USING "   Final: #####"; Final
  PRINT #UBRpt, USING "Inactive: #####"; InActive
  PRINT #UBRpt, USING " Balance: #####"; Balance
  PRINT #UBRpt, USING " Unknown: #####"; UnKnown
  PRINT #UBRpt, USING " Deleted: #####"; DeletedCnt
  PRINT #UBRpt,
  PRINT #UBRpt, USING "   TOTAL: #####"; CustCnt
  PRINT #UBRpt, CHR$(12)

RETURN
END SUB

SUB PrintTransHist (DetailFlag)

'  DetailFlag = True

  BlockClear
  LookUp RecNo&, "Customer Transaction", 2, True, False

  IF RecNo& = 0 THEN
    GOTO ExitTransHist
  END IF
  
  T$ = SPACE$(10)
  MaxLines = 50

  FF$ = CHR$(12)
  Dash80$ = STRING$(80, "-")

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM TotalConsump(1 TO 7) AS LONG
  REDIM DidCnt(1 TO 7) AS INTEGER

  REDIM UBTranRec(1) AS UBTransRecType
  UBTranRecLen = LEN(UBTranRec(1))

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen

  FOR RevCnt = 1 TO MaxRevsCnt
    TempRev$ = QPTrim$(UBSetUpRec(1).Revenues(RevCnt).REVNAME)
    IF LEN(TempRev$) = 0 THEN
      NumOfRevs = RevCnt - 1
      EXIT FOR
    ELSE
      RSET T$ = QPTrim$(LEFT$(TempRev$, 8))
      IF RevCnt <= 8 THEN
        RevText1$ = RevText1$ + T$
      ELSE
        RevText2$ = RevText2$ + T$
      END IF
    END IF
  NEXT

  IF LEN(QPTrim$(RevText2$)) > 0 THEN
    Rev2Flag = True
  END IF
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  GET #UBCust, RecNo&, UBCustRec(1)
  CLOSE UBCust

  FOR MTCnt = 1 TO 7
    IF UBCustRec(1).LocMeters(MTCnt).MTRUnit = "C" THEN
      Cubic = True
      EXIT FOR
    END IF
  NEXT

  UBRpt = FREEFILE
  OPEN "UBTRAHIS.RPT" FOR OUTPUT AS UBRpt

  UBTran = FREEFILE
  OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS UBTran LEN = UBTranRecLen

  BlockClear

  DisplayUBScrn "PRORPT"
  QPrintRC "Processing Transaction History.", 11, 26, -1

  GOSUB DOTranHistHeader

  ThisTrans& = UBCustRec(1).LastTrans
  
  FirstTrans = True

  DO WHILE ThisTrans& > 0
    GET #UBTran, ThisTrans&, UBTranRec(1)
      IF FirstTrans THEN
        LastDate$ = Num2Date$(UBTranRec(1).TransDate)
        Year = VAL(RIGHT$(LastDate$, 4))
        PDate = Date2Num(LEFT$(LastDate$, 3) + "01-" + QPTrim$(STR$(Year - 1)))
        'PDate = 0
        FirstTrans = False
      END IF
      'IF UBTranRec(1).TransDate < PDate THEN
      '  EXIT DO
      'END IF
      GOSUB DOTransDetail
      PRINT #UBRpt, Dash80$
      LineCnt = LineCnt + 2
      IF LineCnt > MaxLines THEN
        PRINT #UBRpt, FF$
        GOSUB DOTranHistHeader
      END IF
    ThisTrans& = UBTranRec(1).PrevTrans
  LOOP
 ' LPRINT Sew#
  GOSUB DOTranHistFooter

  CLOSE

  IF NOT AbortFlag THEN
    PrintRptFile "Customer Transaction Report.", "UBTRAHIS.RPT", 1, RetCode, EntryPoint
  END IF

ExitTransHist:
EXIT SUB

DOTransDetail:
  PRINT #UBRpt, Num2Date(UBTranRec(1).TransDate);
  'PRINT #UBRpt, ThisTrans&;

'  IF UBTranRec(1).TransType = TranUtilityBill THEN
'    Sew# = Round#(Sew# + UBTranRec(1).RevAmt(2))
'    GOTO skipit
'  END IF


  SELECT CASE UBTranRec(1).TransType
    CASE TranUtilityBill, TranUtilityBill + 100
      DidEst = False
      FOR EstCnt = 1 TO 7
        IF UBTranRec(1).EstRead(EstCnt) = "Y" THEN
          DidEst = True
          EXIT FOR
        END IF
      NEXT

      PRINT #UBRpt, TAB(16); "Utility Bill";
      IF DidEst THEN
        PRINT #UBRpt, "*e";
      END IF
      IF DetailFlag THEN
        PRINT #UBRpt, TAB(31); Num2Date$(UBTranRec(1).ReadDate); TAB(43); Num2Date$(UBTranRec(1).PrevDate);
      END IF
      PRINT #UBRpt, TAB(57); USING "$$#####.##"; UBTranRec(1).TransAmt; TAB(71); UBTranRec(1).RunBalance;
      IF DetailFlag THEN
        GOSUB DoMtrDetail
        GOSUB PrintRevDetail
      ELSE
        PRINT #UBRpt,
      END IF
    CASE TranLateCharge, TranLateCharge + 100
      PRINT #UBRpt, TAB(16); "Late Charge"; TAB(57); USING "$$#####.##"; UBTranRec(1).TransAmt; TAB(71); UBTranRec(1).RunBalance
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranReconnectFee, TranReconnectFee + 100
      PRINT #UBRpt, TAB(16); "Reconnect Fee";
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranBillPayment, TranBillPayment + 100
      PRINT #UBRpt, TAB(16); "Bill Payment"; TAB(57); USING "$$#####.##"; UBTranRec(1).TransAmt; TAB(71); UBTranRec(1).RunBalance
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranAppliedDeposit, TranAppliedDeposit + 100
      PRINT #UBRpt, TAB(16); "Applied Deposit"; TAB(57); USING "$$#####.##"; UBTranRec(1).TransAmt; TAB(67); UBTranRec(1).RunBalance
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranPenaltyCharge, TranPenaltyCharge + 100
      PRINT #UBRpt, TAB(16); "Penalty Charge"; TAB(57); USING "$$#####.##"; UBTranRec(1).TransAmt; TAB(71); UBTranRec(1).RunBalance
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranDepositPayment, TranDepositPayment + 100
      PRINT #UBRpt, TAB(16); "Deposit Payment"; TAB(57); USING "$$#####.##"; UBTranRec(1).TransAmt; TAB(71); UBTranRec(1).RunBalance
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranDraftPayment, TranDraftPayment + 100
      PRINT #UBRpt, TAB(16); "Draft Payment";
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranRefundDeposit, TranRefundDeposit + 100
      PRINT #UBRpt, TAB(16); "Refund Deposit"; TAB(57); USING "$$#####.##"; UBTranRec(1).TransAmt; TAB(71); UBTranRec(1).RunBalance
    CASE TranBeginBalance, TranBeginBalance + 100
      PRINT #UBRpt, TAB(16); "Beginning Balance";
    CASE TranUpwardAdjustment, TranUpwardAdjustment + 100
      PRINT #UBRpt, TAB(16); "UP Adjustment  " + LEFT$(UBTranRec(1).BillMsg, 25); TAB(57); USING "$$#####.##"; UBTranRec(1).TransAmt; TAB(71); UBTranRec(1).RunBalance
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranDownwardAdjustment, TranDownwardAdjustment + 100
      PRINT #UBRpt, TAB(16); "DN Adjustment  " + LEFT$(UBTranRec(1).BillMsg, 25); TAB(57); USING "$$#####.##"; UBTranRec(1).TransAmt; TAB(71); UBTranRec(1).RunBalance
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
    CASE TranMiscPayment, TranMiscPayment + 100
      PRINT #UBRpt, TAB(16); "Misc Payment"
      IF DetailFlag THEN
        GOSUB PrintRevDetail
      END IF
  END SELECT
skipit:
RETURN

DoMtrDetail:
  DidAMeter = False
  FOR MtrCnt = 1 TO 7
    IF UBTranRec(1).MtrTypes(MtrCnt) > 0 THEN
      DidAMeter = True
      SELECT CASE UBTranRec(1).MtrTypes(MtrCnt)
      CASE MtrWaterOnly
        MeterType$ = "      Water"
      CASE MtrSewerOnly
        MeterType$ = "      Sewer"
      CASE MtrCombined
        MeterType$ = "Water/Sewer"
      CASE MtrElectric
        MeterType$ = "   Electric"
      CASE MtrDemand
        MeterType$ = " D Electric"
      CASE MtrGas
        MeterType$ = "  Gas Meter"
      CASE MtrTouchRead
        MeterType$ = " Touch Read"
      CASE MtrLightsService
        MeterType$ = "  L Service"
      END SELECT
      WhatMtrCNT = UBTranRec(1).MtrTypes(MtrCnt)
      IF WhatMtrCNT = 0 THEN
        WhatMtrCNT = 1
      END IF
      GOSUB PrintMtrDetail
    END IF
  NEXT
  IF NOT DidAMeter THEN
    MeterType$ = "        "
    'WhatMtrCNT = 1
    MtrCnt = 1
    GOSUB PrintMtrDetail
  END IF
RETURN

PrintMtrDetail:
  PRINT #UBRpt, TAB(16); MeterType$;
  PRINT #UBRpt, TAB(31); USING "##########"; UBTranRec(1).CurRead(MtrCnt);
  PRINT #UBRpt, TAB(43); USING "##########"; UBTranRec(1).PrevRead(MtrCnt);
  MeterConsp& = UBTranRec(1).CurRead(MtrCnt) - UBTranRec(1).PrevRead(MtrCnt)
  IF MeterConsp& < 0 THEN
    MaxMeterAmt& = 10& ^ (LEN(STR$(UBTranRec(1).PrevRead(MtrCnt))) - 1)
    MeterConsp& = (MaxMeterAmt& - UBTranRec(1).PrevRead(MtrCnt)) + UBTranRec(1).CurRead(MtrCnt)
  END IF
  IF Cubic THEN
    MeterConsp& = MeterConsp& * 7.481
  END IF
  PRINT #UBRpt, TAB(57); USING "##########"; MeterConsp&
  IF DidAMeter THEN
    TotalConsump(WhatMtrCNT) = TotalConsump(WhatMtrCNT) + MeterConsp&
    DidCnt(WhatMtrCNT) = DidCnt(WhatMtrCNT) + 1
  END IF
  LineCnt = LineCnt + 1
RETURN

PrintRevDetail:
    'PRINT #UBRpt, TAB(49); USING "$$#####.##"; UBTranRec(1).TransAmt
    PrintedOne = False
    FOR RevCnt = 0 TO 7
      IF UBTranRec(1).RevAmt(RevCnt + 1) <> 0 THEN
        PrintedOne = True
        TabStop = (RevCnt * 10) + 1
        PRINT #UBRpt, TAB(TabStop); USING "#######.##"; UBTranRec(1).RevAmt(RevCnt + 1);
      END IF
    NEXT
    IF PrintedOne THEN
      PRINT #UBRpt,
      LineCnt = LineCnt + 1
    END IF
    RevOffset = 7
    PrintedOne = False
    FOR RevCnt = 0 TO 6
      IF UBTranRec(1).RevAmt(RevCnt + 1 + RevOffset) <> 0 THEN
        PrintedOne = True
        TabStop = (RevCnt * 10) + 1
        PRINT #UBRpt, TAB(TabStop); USING "#######.##"; UBTranRec(1).RevAmt(RevCnt + 1 + RevOffset);
      END IF
    NEXT
    IF PrintedOne THEN
      PRINT #UBRpt,
      LineCnt = LineCnt + 1
    END IF

RETURN

DOTranHistHeader:
  LineCnt = 7
  PRINT #UBRpt, TAB(28); "Transaction History Report. "
  PRINT #UBRpt, "Customer: "; UBCustRec(1).CustName; TAB(57); "Report Date: "; DATE$
  IF DetailFlag THEN
    PRINT #UBRpt, " Account:"; RecNo&
    PRINT #UBRpt, "Ser Addr: "; UBCustRec(1).SERVADDR
    'PRINT #UBRpt, TAB(11); UBCustRec(1).Addr1
    'PRINT #UBRpt, TAB(11); UBCustRec(1).Addr2
    'PRINT #UBRpt, TAB(11); UBCustRec(1).City; UBCustRec(1).State; "  "; UBCustRec(1).ZipCode
    LineCnt = LineCnt + 2
    FOR Mtr = 1 TO 7
      IF LEN(QPTrim$(UBCustRec(1).LocMeters(Mtr).MTRNUM)) > 0 THEN
        PRINT #UBRpt, TAB(6); "Mtr# "; QPTrim$(UBCustRec(1).LocMeters(Mtr).MTRNUM)
        LineCnt = LineCnt + 1
      END IF
    NEXT
  END IF
  PRINT #UBRpt,
  IF DetailFlag THEN
    PRINT #UBRpt, "Trans Date     Trans Type     Cur.Date     Pre.Date      TR Amount      Balance"
  ELSE
    PRINT #UBRpt, "Trans Date     Trans Type                                TR Amount      Balance"
  END IF
  IF DetailFlag THEN
    PRINT #UBRpt, "               Meter Type     Cur.Read     Pre.Read       Usage"
  END IF
  IF DetailFlag THEN
    PRINT #UBRpt, RevText1$
    IF Rev2Flag THEN
      PRINT #UBRpt, RevText2$
      LineCnt = 8
    END IF
  ELSE
    LineCnt = 5
  END IF
  PRINT #UBRpt, Dash80$
RETURN

DOTranHistFooter:
  IF FirstTrans THEN
    PRINT #UBRpt, "NO TRANSACTIONS!!!"
    PRINT #UBRpt, Dash80$
  END IF
  FOR MCnt = 1 TO 7
    IF DidCnt(MCnt) > 0 THEN
      PRINT #UBRpt, USING "Average Consumption: #########"; TotalConsump(MCnt) / DidCnt(MCnt)
    END IF
  NEXT
  PRINT #UBRpt, CHR$(12)
RETURN

END SUB

SUB UBQuickCustRptMenu



END SUB

