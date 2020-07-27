DECLARE SUB DisplayPrnScrn (ScrnName$)
DECLARE SUB PrintAlignMask ()
DEFINT A-Z
DECLARE SUB FindLastBill (RecNo&, TransNo&)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE SUB SmallPause ()
DECLARE SUB PrintAlign1Mask ()
DECLARE SUB Post1BillTrans (VoidLastFlag%)
DECLARE SUB Print1Bill (RecNo&)
DECLARE SUB Calc1Bill (RecNo&, FromFlag%)
DECLARE SUB VoidReprintMenu ()
DECLARE SUB IndBillPrintMenu ()
DECLARE SUB SendDist2GL ()
DECLARE SUB UBMiscMenu ()
DECLARE SUB UBBillMenu ()
DECLARE SUB UBCustomerMenu ()
DECLARE SUB ClearScrn ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB BlockClear ()
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION GetAdjFactor# ()
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION FindRateTbl% (RATECODE$, NumOfRates%, UBRateTbls() AS ANY)
DECLARE FUNCTION GetCustMeterType% (UBLocatRec() AS ANY, ThisMeter%)
DECLARE FUNCTION GetRevCharge# (RateTbl AS ANY, TMeterConsp&, MeterMulti&)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNum%)

'$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB WazzWind (BYVAL a, BYVAL b, BYVAL c, BYVAL d, BYVAL E, BYVAL f, BYVAL g)
DECLARE SUB RestScrn (Array())
DECLARE SUB SaveScrn (Array())
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB UBLog (Text$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE FUNCTION Date2Num% (DateString$)

  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'newcust.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'ubpinfo.bi'
  '$INCLUDE: 'ubdraft.BI'

  STACK 5000
  
  CONST False = 0, True = NOT False
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 3)

  MChoice$(1) = " Reprint Customers LAST Bill "
  MChoice$(2) = " Exit to OS"
  
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
    
    TitleBox 2, Col, MaxLen + 3, "Posted Bill Reprinting Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      BlockClear
      LookUp RecNo&, "Posted Bill Reprint", 2, True, True
      IF RecNo& > 0 THEN
        FindLastBill RecNo&, TransNo&
        IF TransNo& > 0 THEN
          'Print1Bill RecNo&, TransNo&
        END IF
      END IF

    CASE 2
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
  IF INSTR(COMMAND$, "TEST") = 0 THEN
    RUN "ubprnbil"
  ELSE
    HideCursor
    ClearScrn
  END IF
  
  END

SUB Print1Bill (RecNo&)

  LPIFlag = False
  
  REDIM ScrnArray(0)

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen      'load setup file

  'Section to check for customer modifications

  TownName$ = UBSetUpRec(1).UTILNAME

  IF INSTR(TownName$, "INDIAN TRAIL") THEN
    IndianFlag = True
  END IF

  IF UBSetUpRec(1).BANKDFT = "Y" THEN
    UseDraftFlag = True
  END IF

  LibName$ = "UB"

  ScrnName$ = "PRNIBILL"

  MActionRow = 19

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

  FrstBFld = FldNum%("FRSTBILL", Fld())
  LastBFld = FldNum%("LASTBILL", Fld())

  BillDFld = FldNum%("BILLDATE", Fld())
  PastDFld = FldNum%("PASTDATE", Fld())

  PRDateFld = FldNum%("PRDATE", Fld())
  CRDateFld = FldNum%("CRDATE", Fld())

  DraftDFld = FldNum%("DRFTDATE", Fld())
  MesgFld = FldNum%("MSGLINE1", Fld())

  Desc1Fld = FldNum%("MSGLINE2", Fld())
  Desc2Fld = FldNum%("MSGLINE3", Fld())
  Desc3Fld = FldNum%("MSGLINE4", Fld())

  IF IndianFlag THEN
    Fld(Desc1Fld).Protected = False
    Fld(Desc2Fld).Protected = False
    Fld(Desc3Fld).Protected = False
  END IF

  Today$ = DATE$

  Action = 1
  Frm(1).StayOnField = True

  FirstTime = True
  '--Set screen number to one and display screen

  BlockClear

  DisplayUBScrn ScrnName$

  IF IndianFlag THEN
    FOR Cnt = 1 TO 3
      QPrintRC "Description " + QPTrim$(STR$(Cnt)) + ":", Cnt + 15, 21, -1
    NEXT
  END IF

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(BillDFld, 0) = Today$
      LSET Form$(PastDFld, 0) = Today$
      LSET Form$(DraftDFld, 0) = Today$
      Action = 2
    END IF

    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB CheckReqFields
      IF ReqFldsOK THEN
        ExitFlag = True
      ELSE
        Action = 2
      END IF
    CASE F5KEY
      PrintAlign1Mask
    CASE EscKey
      AbortFlag = True
      ExitFlag = True
    END SELECT

    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE MActionRow           'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 18 TO 30           '--Cancel button
          PressButton 27, MActionRow, 18, 30
        CASE 31 TO 46           '--F5 button
          PressButton F5KEY, MActionRow, 31, 46
        CASE 47 TO 64           '--F10 button
          PressButton F10Key, MActionRow, 47, 64
        END SELECT
      END SELECT                'row
    END IF

    '--Check screen page
  LOOP UNTIL ExitFlag

  IF AbortFlag GOTO ExitPrintBill

  PastDate$ = Form$(PastDFld, 0)

  'do bill printing here
  '**************************************************************************

  REDIM BillInfoRec(1) AS PrintBillInfoType

  BillInfoRec(1).MsgLine2 = Form$(Desc1Fld, 0)
  BillInfoRec(1).MsgLine3 = Form$(Desc2Fld, 0)
  BillInfoRec(1).MsgLine4 = Form$(Desc3Fld, 0)

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBDraftPayRec(1) AS UBDraftPayRecType
  UBDraftPayLen = LEN(UBDraftPayRec(1))

  REDIM UBBillRec(1) AS UBTransRecType
  UBBillRecLen = LEN(UBBillRec(1))

  UBBill = FREEFILE
  OPEN UBIBillFile FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBRpt = FREEFILE
  OPEN "UBIBILL.PRN" FOR OUTPUT AS UBRpt

  IF UseDraftFlag THEN
    UBDraft = FREEFILE
    OPEN DFFileName$ FOR RANDOM SHARED AS UBDraft LEN = UBDraftPayLen
  END IF

  ShowProcessingScrn "Creating Utility Bills."

  '-----------------------------------------
  PrintedCnt = 0

  GET UBBill, 1, UBBillRec(1)
  CustAcctNo& = UBBillRec(1).CustAcctNo

  GET UBCust, UBBillRec(1).CustAcctNo, UBCustRec(1)

  Num2Print = UBCustRec(1).BILLCOPY
  IF Num2Print < 1 THEN Num2Print = 1

  IF UBBillRec(1).ActiveFlag THEN
    IF UBBillRec(1).TransAmt <> 0 OR UBCustRec(1).CurrBalance <> 0 OR UBCustRec(1).PrevBalance THEN
      PrintedCnt = PrintedCnt + 1
      UBBillRec(1).BillNumber = PrintedCnt
      UBBillRec(1).TransDate = BillDate
      UBBillRec(1).TransDesc = "UTILITY BILL"

      FOR MtrCnt = 1 TO 7
        IF UBCustRec(1).LocMeters(MtrCnt).CurDate > 0 THEN
          UBBillRec(1).PrevDate = UBCustRec(1).LocMeters(MtrCnt).PastDate
          UBBillRec(1).ReadDate = UBCustRec(1).LocMeters(MtrCnt).CurDate
          EXIT FOR
        END IF
      NEXT

      IF UBBillRec(1).ReadDate <= 0 THEN
        UBBillRec(1).ReadDate = BillDate - 30
      END IF
      IF UBBillRec(1).PrevDate <= 0 THEN
        UBBillRec(1).PrevDate = UBBillRec(1).ReadDate - 30
      END IF

      IF UseEDateFlag THEN
        UBBillRec(1).PrevDate = PRDate
        UBBillRec(1).ReadDate = CRDate
      END IF

      '*****************
      DaysINRead = UBBillRec(1).ReadDate - UBBillRec(1).PrevDate

      UBBillRec(1).BillDate = BillDate
      UBBillRec(1).PastDueDate = PastDate
      UBBillRec(1).DraftDate = DraftDate
      UBBillRec(1).BillMsg = Message$

      'these are for reprinting bills
      'UBBillRec(1)CustLocation = CustAcctNo&
      UBBillRec(1).CustAcctNo = CustAcctNo&

      PrevDate$ = Num2Date$(UBBillRec(1).PrevDate)
      PastDate$ = Num2Date$(UBBillRec(1).PrevDate)
      DateRead$ = Num2Date$(UBBillRec(1).ReadDate)
      PastDueDate$ = Num2Date$(UBBillRec(1).PastDueDate)

      TotalTax# = 0
      FOR TaxCnt = 1 TO MaxRevsCnt - 1
        TotalTax# = Round(TotalTax# + UBBillRec(1).TaxAmt(TaxCnt))
      NEXT

      PUT UBBill, 1, UBBillRec(1)

      DidADraftFlag = False
      IF UseDraftFlag AND UBCustRec(1).USEDRAFT = "Y" AND UBCustRec(1).PreNoteFlag THEN
        UBDraftPayRec(1).CustAcctNum = CustAcctNo&
        UBDraftPayRec(1).DraftAmt = UBBillRec(1).TransAmt
        PUT UBDraft, , UBDraftPayRec(1)
        DidADraftFlag = True
      END IF

      'Custom Mod Here For Lilesville, NC
      IF Lilesville > 0 THEN
        IF UBCustRec(1).Serv(1).RATECODE = "WIN " OR UBCustRec(1).Serv(1).RATECODE = "WOUT" THEN
          TenPercentAmount# = (UBBillRec(1).TransAmt - UBBillRec(1).RevAmt(1)) + (UBBillRec(1).RevAmt(1) * 1.1111)
        ELSE
          TenPercentAmount# = UBBillRec(1).TransAmt
        END IF
      END IF
      'End Lilesville Custom Mod

      FOR BillCopies = 1 TO Num2Print
        '$INCLUDE: 'UBSTDBIL.BI'        'Std  Lilesville
        ''$INCLUDE: 'UBLANDIS.BI'  'landis
        ''$INCLUDE: 'UBLONGVW.BI'  'Longview
        ''$INCLUDE: 'UBHOLLY.BI'   '
        ''$INCLUDE: 'UBTROUTM.BI'   'Troutman Format
        ''$INCLUDE: 'UBMCMICK.BI'  'Mccormick
        ''$INCLUDE: 'UBOLDSDS.BI'  'Cleveland NC
        ''$INCLUDE: 'UBOLDSTD.BI'  'Old citipak standard bill
        ''$INCLUDE: 'UBCONCRD.BI'  'Concord
        ''$INCLUDE: 'UBELLEN.BI'   'Ellenboro NC
        ''$INCLUDE: 'UBGROVER.BI'  'Grover NC
        ''$INCLUDE: 'UBGILES.BI'   'Giles Cty Va
        ''$INCLUDE: 'UBINDIAN.BI'        'INDIAN TRAIL
        ''$INCLUDE: 'UBHRTFRD.BI'        'Hertford
        ''$INCLUDE: 'UBGORDON.BI'        'Gordonsville VA
        ''$INCLUDE: 'UBTUSK.BI'          'Tuskaseigee NC
        ''$INCLUDE: 'UBBURNET.BI'        'Burnettown SC
        ''$INCLUDE: 'UBHEMIBL.BI'        'Hemmingway
      NEXT
    END IF
  END IF

SkipEm:
  
  IF LPIFlag = -2 THEN
    PRINT #UBRpt, CHR$(27); CHR$(50);           'set printer in 6 lines per inch
  END IF

  CLOSE
  
  ERASE Frm, Form$, Fld, UBCustRec, UBBillRec

  IF NOT AbortFlag THEN
    PrintRptFile "Individual Bill Printing", "UBIBILL.PRN", 1, RetCode, 0
  END IF

'**************************************************************************
'end bill printing
  GOTO ExitPrintBill:

CheckReqFields:

  FBillNO& = VAL(Form$(FrstBFld, 0))
  LBillNO& = VAL(Form$(LastBFld, 0))
  BillDate = Date2Num(Form$(BillDFld, 0))

  BillDate$ = Num2Date$(BillDate)

  PastDate = Date2Num(Form$(PastDFld, 0))

  DraftDate = Date2Num(Form$(DraftDFld, 0))
  DueDate$ = Form$(PastDFld, 0)

  Message$ = Form$(MesgFld, 0)

  PRDate = Date2Num(Form$(PRDateFld, 0))
  CRDate = Date2Num(Form$(CRDateFld, 0))

  IF (CRDate > 0) AND (PRDate > 0) THEN
    UseEDateFlag = True
  ELSE
    UseEDateFlag = False
  END IF

  IF BillDate = -32768 THEN
    SaveScrn ScrnArray()
    DisplayUBScrn "BADBDATE"
    WaitForAction
    RestScrn ScrnArray()
    Frm(1).FldNo = BillDFld
  ELSEIF PastDate < BillDate THEN
    SaveScrn ScrnArray()
    DisplayUBScrn "BADBDATE"
    WaitForAction
    RestScrn ScrnArray()
    Frm(1).FldNo = PastDFld
  ELSEIF (UseDraftFlag AND DraftDate = -32768) OR (UseDraftFlag AND DraftDate < BillDate) THEN
    SaveScrn ScrnArray()
    DisplayUBScrn "BADDDATE"
    WaitForAction
    RestScrn ScrnArray()
    Frm(1).FldNo = 3
  ELSE
    ReqFldsOK = True
  END IF

  IF UseDraftFlag THEN
    DraftDate$ = Form$(DraftDFld, 0)
    DFFileName$ = "DF" + LEFT$(DraftDate$, 2) + MID$(DraftDate$, 4, 2) + RIGHT$(DraftDate$, 2) + ".DAT"
  END IF

  RETURN

ExitPrintBill:

END SUB

SUB PrintAlign1Mask
  CursorOff
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  PrintRptFile "Utility Bill Alignment Mask ", "UBBLMASK.DAT", 1, RetCode, 4
  RestScrn ScrnArray()
  ERASE ScrnArray
END SUB

SUB SmallPause
  ST! = TIMER
  DO: LOOP UNTIL TIMER > ST! + .9
END SUB

