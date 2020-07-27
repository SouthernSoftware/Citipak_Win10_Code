DEFINT A-Z
DECLARE FUNCTION GetPrevBillCnt% (UBCustRec() AS ANY)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (N#)
DECLARE SUB BlockClear ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB HideCursor ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB MakeBillFile (AbortFlag%, FuelAdjAmt#)
DECLARE SUB MakePostalIndex (IndexText$)
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB PrintAlignMask ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB PrintUtilBills ()
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RePrintUtilBills ()
DECLARE SUB RestScrn (Array())
DECLARE SUB SaveScrn (Array())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'Newcust.bi'
  '$INCLUDE: 'ubpinfo.bi'
  
  CONST False = 0, True = NOT False

  STACK 8000
  
  'GOTO Skip2Here
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 4)
  
  MChoice$(1) = "Print ALL Utility Bills "
  MChoice$(2) = "Print/Reprint Selected Bills"
  'MChoice$(3) = "Print Final/Individual Bills"
  
  MChoice$(3) = "Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 17   'limit the box length to go no lower than line 20
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
  Col = ((80 - MaxLen) \ 2) - 2
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 3, Col, MaxLen + 3, "Bill Printing Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN
      Choice = 0
      EXIT DO
    END IF
    
    SELECT CASE Choice
    CASE 1
      PrintUtilBills
    CASE 2
      RePrintUtilBills
    CASE 3
      ExitFlag = True
      EXIT DO
    END SELECT
    
  LOOP
  IF ExitFlag THEN
    HideCursor
    ClearScrn
  ELSE
    RUN "ubbillin"
  END IF
  
  END

FUNCTION GetPrevBillCnt (UBCustRec() AS NewUBCustRecType)

  REDIM UBTransRec(1) AS UBTransRecType
  TranLen = LEN(UBTransRec(1))
  PrevBillCnt = 0

  ThisTrans& = UBCustRec(1).LastTrans

  IF ThisTrans& > 0 THEN
    REDIM UBTransRec(1) AS UBTransRecType
    TranFile = FREEFILE
    OPEN "UBTRANS.DAT" FOR RANDOM SHARED AS TranFile LEN = TranLen
    DO WHILE ThisTrans& > 0
      GET #TranFile, ThisTrans&, UBTransRec(1)
      IF UBTransRec(1).TransType = TranUtilityBill THEN
        'HasPrevBill = True
        PrevBillCnt = PrevBillCnt + 1
        IF PrevBillCnt >= 2 THEN
          EXIT DO
        END IF
      END IF
      ThisTrans& = UBTransRec(1).PrevTrans
    LOOP
    CLOSE TranFile
  END IF
  GetPrevBillCnt = PrevBillCnt

END FUNCTION

SUB PrintAlignMask
  REDIM ScrnArray(0)
  CursorOff
  SaveScrn ScrnArray()
  BlockClear
  PrintRptFile "Utility Bill Alignment Mask ", "UBBLMASK.DAT", 1, RetCode, 4
  RestScrn ScrnArray()
END SUB

SUB PrintUtilBills
  
  SHARED Choice$()
  
  LPIFlag = False
  
  IF NOT Exist("UBBILLS.DAT") THEN
    CursorOff
    BlockClear
    DisplayUBScrn "NON2PRNT"
    WaitForAction
    GOTO ExitPrintBill
  END IF
  
  REDIM BillInfoRec(1) AS PrintBillInfoType
  BillInfoRecLen = LEN(BillInfoRec(1))
  
  REDIM ScrnArray(0)
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen      'load setup file
  
  TownName$ = UBSetUpRec(1).UTILNAME
  
  'Section to check for customer modifications
  'Town of Lilesville Special Discount Situation

  IF INSTR(TownName$, "INDIAN TRAIL") THEN
    IndianFlag = True
  END IF
  
  IF UBSetUpRec(1).BANKDFT = "Y" THEN
    UseDraftFlag = True
  END IF
  
  REDIM Choice$(0 TO 6, 0)
  
  Choice$(1, 0) = "Customer Name Order"
  Choice$(2, 0) = "Account Number Order"
  Choice$(3, 0) = "Location Number Order"
  Choice$(4, 0) = "Postal Carrier Route Order"
  Choice$(5, 0) = "ZipCode Order"
  
  IF UBSetUpRec(1).UseSeq = "Y" THEN
    Choice$(6, 0) = "Sequence Number Order"
  END IF
  
  LibName$ = "UB"
  
  ScrnName$ = "PRNTBILL"

  Choice$(0, 0) = "6"

  MActionRow = 20
  
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
  BillDFld = FldNum%("BILLDATE", Fld())
  PastDFld = FldNum%("PASTDATE", Fld())
  DraftDFld = FldNum%("DRFTDATE", Fld())
  BillOFld = FldNum%("PRNORDER", Fld())
  MesgFld = FldNum%("MSGLINE", Fld())

  PRDateFld = FldNum%("PRDATE", Fld())
  CRDateFld = FldNum%("CRDATE", Fld())

  Desc1Fld = FldNum%("DESC1", Fld())
  Desc2Fld = FldNum%("DESC2", Fld())
  Desc3Fld = FldNum%("DESC3", Fld())

  IF IndianFlag THEN
    Fld(Desc1Fld).Protected = False
    Fld(Desc2Fld).Protected = False
    Fld(Desc3Fld).Protected = False
  END IF
  
  Action = 1
  Frm(1).StayOnField = True
  
  FirstTime = True
  Today = Date2Num(DATE$)
  PastDay = Today + 10
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayUBScrn ScrnName$

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(BillDFld, 0) = DATE$
      LSET Form$(PastDFld, 0) = Num2Date$(PastDay)
      Action = 1
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
      PrintAlignMask
      
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
  
  'PastDate$ = Form$(2, 0)
  
  IF AbortFlag THEN GOTO ExitPrintBill
  
  SELECT CASE BillOrder$
  CASE "C"
    IndexName$ = NameIndexFile
    UsingName = True
    OKFlag = True
  CASE "A"
    IndexName$ = ""
    UsingAcct = True
    OKFlag = True
  CASE "L"
    IndexName$ = BookIndexFile
    UsingBook = True
    OKFlag = True
  CASE "P", "Z"
    IF BillOrder$ = "P" THEN
      IdxTypeText$ = "Postal Route"
    ELSE
      IdxTypeText$ = "Zip-Code"
    END IF
    MakePostalIndex IdxTypeText$
    IndexName$ = TempIndexName
    OKFlag = True
  CASE "S"
    IdxTypeText$ = "Sequence Number"
    MakeSequenceIndex IdxTypeText$
    IndexName$ = TempIndexName
    OKFlag = True
  END SELECT
  
  IF AbortFlag GOTO ExitPrintBill
  
  PastDate$ = Form$(PastDFld, 0)
  
  'do bill printing here
  '**************************************************************************
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBBillRec(1) AS UBTransRecType
  UBBillRecLen = LEN(UBBillRec(1))
  
  REDIM UBDraftPayRec(1) AS UBDraftPayRecType
  UBDraftPayLen = LEN(UBDraftPayRec(1))

  BillInfoRec(1).BillDate = Date2Num(Form$(BillDFld, 0))
  BillInfoRec(1).PastDate = Date2Num(Form$(PastDFld, 0))
  BillInfoRec(1).PRDate = Date2Num(Form$(PRDateFld, 0))
  BillInfoRec(1).CRDate = Date2Num(Form$(CRDateFld, 0))
  BillInfoRec(1).DrftDate = Date2Num(Form$(DraftDFld, 0))
  BillInfoRec(1).PrnOrder = QPTrim$(Form$(BillOFld, 0))
  BillInfoRec(1).MsgLine = Form$(MesgFld, 0)

  BillInfoRec(1).Desc1 = Form$(Desc1Fld, 0)
  BillInfoRec(1).Desc2 = Form$(Desc2Fld, 0)
  BillInfoRec(1).Desc3 = Form$(Desc3Fld, 0)

  'GOSUB UpdateInfoRec

  IF UsingAcct THEN             'load the index
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  ELSE
    NumOfRecs = FileSize(IndexName$) \ 4
    REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
  END IF
  
  UBBill = FREEFILE
  OPEN "UBBILLS.DAT" FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBBILLS.PRN" FOR OUTPUT AS UBRpt
  
  IF UseDraftFlag THEN
    UBDraft = FREEFILE
    OPEN DFFileName$ FOR RANDOM SHARED AS UBDraft LEN = UBDraftPayLen
  END IF
  
  ShowProcessingScrn "Creating Utility Bills."
  
  '-----------------------------------------
  PrintedCnt = 0
  NotDone = True

  FOR Cnt = 1 TO NumOfRecs
    IF UsingAcct THEN
      CustAcctNo& = Cnt
    ELSE
      CustAcctNo& = IndexArray(Cnt).RecNum
    END IF
    
    GET UBCust, CustAcctNo&, UBCustRec(1)
    
    Num2Print = UBCustRec(1).BILLCOPY
    IF Num2Print < 1 THEN Num2Print = 1
    
    IF UBCustRec(1).STATUS <> "F" AND UBCustRec(1).STATUS <> "I" THEN
      GET UBBill, CustAcctNo&, UBBillRec(1)
      IF UBBillRec(1).ActiveFlag THEN
        
        IF UBBillRec(1).TransAmt <> 0 OR UBCustRec(1).PrevBalance THEN
          PrintedCnt = PrintedCnt + 1
          UBBillRec(1).BillNumber = PrintedCnt
          UBBillRec(1).TransDate = BillDate
          UBBillRec(1).TransDesc = "UTILITY BILL"
          '****************
          '02-05-97 added code to get a valid meter read date,  maybe?  ;-]
          '         from one of the meters
          FOR MtrCnt = 1 TO 7
            IF UBBillRec(1).MtrTypes(MtrCnt) > 0 THEN
              UBBillRec(1).PrevDate = UBCustRec(1).LocMeters(MtrCnt).PastDate
              UBBillRec(1).ReadDate = UBCustRec(1).LocMeters(MtrCnt).CurDate
              EXIT FOR
            END IF
          NEXT
          '02-05-97 this is to add a read date to a bill that has no metered
          '         services
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

          UBBillRec(1).BillDate = BillDate
          UBBillRec(1).PastDueDate = PastDate
          UBBillRec(1).DraftDate = DraftDate
          UBBillRec(1).BillMsg = Message$
          
          'these are for reprinting bills
          UBBillRec(1).CustLocation = CustAcctNo&
          UBBillRec(1).CustAcctNo = CustAcctNo&

          BillDate$ = Num2Date$(UBBillRec(1).BillDate)

'if they entered a previous read date
          IF BillInfoRec(1).PRDate > 0 THEN
            PrevDate$ = Num2Date$(BillInfoRec(1).PRDate)
            PastDate$ = Num2Date$(BillInfoRec(1).PRDate)
            DateRead$ = Num2Date$(BillInfoRec(1).CRDate)
            DaysINRead = BillInfoRec(1).CRDate - BillInfoRec(1).PRDate
          ELSE
            PrevDate$ = Num2Date$(UBBillRec(1).PrevDate)
            PastDate$ = Num2Date$(UBBillRec(1).PrevDate)
            DateRead$ = Num2Date$(UBBillRec(1).ReadDate)
            DaysINRead = UBBillRec(1).ReadDate - UBBillRec(1).PrevDate
          END IF
          
          PastDueDate$ = Num2Date$(UBBillRec(1).PastDueDate)
          
          PUT UBBill, CustAcctNo&, UBBillRec(1)
          TotalTax# = 0
          
          FOR TaxCnt = 1 TO MaxRevsCnt
            TotalTax# = Round(TotalTax# + UBBillRec(1).TaxAmt(TaxCnt))
          NEXT
          
          DidADraftFlag = False
          IF UseDraftFlag AND UBCustRec(1).USEDRAFT = "Y" AND UBCustRec(1).PreNoteFlag THEN
            UBDraftPayRec(1).CustAcctNum = CustAcctNo&
            UBDraftPayRec(1).DraftAmt = UBBillRec(1).TransAmt
            PUT UBDraft, , UBDraftPayRec(1)
            DidADraftFlag = True
          END IF

          'PrevCnt = GetPrevBillCnt(UBCustRec())

          FOR BillCopies = 1 TO Num2Print
            '$INCLUDE: 'UBWADES.BI'         '
          NEXT
          
        END IF
      END IF
    END IF
    
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    ShowPctComp Cnt, NumOfRecs
SkipEm:
  NEXT
  
  IF LPIFlag = -2 THEN
    PRINT #UBRpt, CHR$(27); CHR$(50);           'set printer in 6 lines per inch
  END IF
  
  CLOSE
  
  BillInfoRec(1).FrstBill = 1
  BillInfoRec(1).LastBill = PrintedCnt
  BillInfoRec(1).BillDate = Date2Num(Form$(BillDFld, 0))
  BillInfoRec(1).PastDate = Date2Num(Form$(PastDFld, 0))
  BillInfoRec(1).PRDate = Date2Num(Form$(PRDateFld, 0))
  BillInfoRec(1).CRDate = Date2Num(Form$(CRDateFld, 0))
  BillInfoRec(1).DrftDate = Date2Num(Form$(DraftDFld, 0))
  BillInfoRec(1).PrnOrder = QPTrim$(Form$(BillOFld, 0))
  BillInfoRec(1).MsgLine = Form$(MesgFld, 0)

  BillInfoRec(1).Desc1 = Form$(Desc1Fld, 0)
  BillInfoRec(1).Desc2 = Form$(Desc2Fld, 0)
  BillInfoRec(1).Desc3 = Form$(Desc3Fld, 0)

  UBFile = FREEFILE
  OPEN "UBPINFON.DAT" FOR RANDOM AS #UBFile LEN = BillInfoRecLen
  PUT #UBFile, 1, BillInfoRec(1)
  CLOSE
  
  ERASE Frm, Form$, Fld, UBCustRec, UBBillRec, BillInfoRec
  
  IF NOT AbortFlag THEN
    PrintRptFile "Utility Bill Printing ", "UBBILLS.PRN", 1, RetCode, 1
  END IF
  
  '**************************************************************************
  'end bill printing
  GOTO ExitPrintBill:
  '******************
  
  '******************
CheckReqFields:
  
  BillDate = Date2Num(Form$(BillDFld, 0))
  PastDate = Date2Num(Form$(PastDFld, 0))
  DraftDate = Date2Num(Form$(DraftDFld, 0))
  DueDate$ = Form$(PastDFld, 0)
  BillOrder$ = QPTrim$(LEFT$(Form$(BillOFld, 0), 1))
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
  ELSEIF LEN(BillOrder$) = 0 THEN
    SaveScrn ScrnArray()
    DisplayUBScrn "NOBORDER"
    WaitForAction
    RestScrn ScrnArray()
    Frm(1).FldNo = 4
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

SUB RePrintUtilBills
  
  LPIFlag = False
  
  REDIM TempScrn(0)
  
  IF NOT Exist("UBBILLS.PRN") THEN
    CursorOff
    BlockClear
    DisplayUBScrn "NOTPRNTD"
    WaitForAction
    GOTO ExitRePrintBill
  END IF
  
  NotBeenDone = True
  
  REDIM UBBillRec(1) AS UBTransRecType
  UBBillRecLen = LEN(UBBillRec(1))
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen      'load setup file
  
  TownName$ = UBSetUpRec(1).UTILNAME
  
'Section to check for customer modifications
  IF INSTR(TownName$, "INDIAN TRAIL") THEN
    IndianFlag = True
  END IF

  ShowProcessingScrn "Indexing Billing Information."
  QPrintRC "      Reading Bill Records       ", 11, 25, -1
  
  UBBill = FREEFILE
  OPEN "UBBILLS.DAT" FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  NumOfRecs = LOF(UBBill) \ UBBillRecLen
  REDIM RePrintIdx(1 TO NumOfRecs) AS RePrintIndexType
  FOR Cnt = 1 TO NumOfRecs
    GET UBBill, Cnt, UBBillRec(1)
    IF UBBillRec(1).ActiveFlag THEN
      RePrintIdx(Cnt).BillNum = UBBillRec(1).BillNumber
      RePrintIdx(Cnt).BillRec = Cnt
    END IF
    ShowPctComp Cnt, NumOfRecs
  NEXT
  
  CLOSE UBBill
  
  REDIM BillInfoRec(1) AS PrintBillInfoType
  BillInfoRecLen = LEN(BillInfoRec(1))
  
  UBFile = FREEFILE
  OPEN "UBPINFON.DAT" FOR RANDOM AS #UBFile LEN = BillInfoRecLen
  GET #UBFile, 1, BillInfoRec(1)
  CLOSE
  
  QPrintRC "         Sorting Index.        ", 11, 25, -1
  SortT RePrintIdx(1), NumOfRecs, 0, 8, 0, -2
  
  '--Initialize the form name array
  LibName$ = "UB"
  ScrnName$ = "PRNTBIL2"
  MActionRow = 20
  
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
  BillOFld = FldNum%("PRNORDER", Fld())
  MesgFld = FldNum%("MSGLINE", Fld())
  Desc1Fld = FldNum%("DESC1", Fld())
  Desc2Fld = FldNum%("DESC2", Fld())
  Desc3Fld = FldNum%("DESC3", Fld())

  Action = 1
  Frm(1).StayOnField = True
  
  '--Set screen number to one and display screen
  FirstTime = True
  
  BlockClear
  
  DisplayUBScrn ScrnName$

'mod for Indian Trail
  IF IndianFlag THEN
    FOR Cnt = 1 TO 3
      QPrintRC "Description " + QPTrim$(STR$(Cnt)) + ":", Cnt + 15, 21, -1
    NEXT
  END IF
'end Indian Trail mod

  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      GOSUB FillForm
      FirstTime = False
    END IF
    '--Check for Key presses
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      GOSUB CheckFields
      IF ReqFldsOK THEN
        ExitFlag = True
      ELSE
        Action = 2
      END IF
    CASE F5KEY
      PrintAlignMask
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
  
  IF AbortFlag GOTO ExitRePrintBill:
  
  ShowProcessingScrn "Creating Utility Bills."
  
  UBBill = FREEFILE
  OPEN "UBBILLS.DAT" FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBBILLSR.PRN" FOR OUTPUT AS UBRpt
  
  Cnt = 0
  DO
    Cnt = Cnt + 1
    IF Cnt > NumOfRecs THEN EXIT DO
    ThisBillNum = RePrintIdx(Cnt).BillNum
    
    IF ThisBillNum >= FirstBill AND ThisBillNum <= LastBill THEN
      PrintedCnt = ThisBillNum
      GET UBBill, RePrintIdx(Cnt).BillRec, UBBillRec(1)
      GET UBCust, RePrintIdx(Cnt).BillRec, UBCustRec(1)

      CustAcctNo& = RePrintIdx(Cnt).BillRec
      BillDate$ = Num2Date$(UBBillRec(1).BillDate)

      PrevDate$ = Num2Date$(UBBillRec(1).PrevDate)
      PastDate$ = Num2Date$(UBBillRec(1).PrevDate)
      
      'DateRead$ = Num2Date$(UBBillRec(1).ReadDate)
      DaysINRead = UBBillRec(1).ReadDate - UBBillRec(1).PrevDate
      PastDueDate$ = Num2Date$(UBBillRec(1).PastDueDate)
      
'if they entered a previous read date
          IF BillInfoRec(1).PRDate > 0 THEN
            PrevDate$ = Num2Date$(BillInfoRec(1).PRDate)
            PastDate$ = Num2Date$(BillInfoRec(1).PRDate)
            DateRead$ = Num2Date$(BillInfoRec(1).CRDate)
            DaysINRead = BillInfoRec(1).CRDate - BillInfoRec(1).PRDate
          ELSE
            PrevDate$ = Num2Date$(UBBillRec(1).PrevDate)
            PastDate$ = Num2Date$(UBBillRec(1).PrevDate)
            DateRead$ = Num2Date$(UBBillRec(1).ReadDate)
            DaysINRead = UBBillRec(1).ReadDate - UBBillRec(1).PrevDate
          END IF

      ''Custom Mod Here For Lilesville, NC
      'IF Lilesville > 0 THEN
      '  IF UBCustRec(1).Serv(1).RateCode = "WIN " OR UBCustRec(1).Serv(1).RateCode = "WOUT" THEN
      
      '  ELSE
      '    TenPercentAmount# = UBBillRec(1).TransAmt
      '  END IF
      'END IF
      'End Lilesville Custom Mod
      
      Num2Print = UBCustRec(1).BILLCOPY
      IF Num2Print < 1 THEN Num2Print = 1
      FOR BillCopies = 1 TO Num2Print
        '$INCLUDE: 'UBWADES.BI'         '
      NEXT
    END IF
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT DO
    END IF
    ShowPctComp Cnt, NumOfRecs
    
  LOOP
  IF LPIFlag = -2 THEN
    PRINT #UBRpt, CHR$(27); CHR$(50);           'set printer in 6 lines per inch
  END IF
  
  CLOSE
  
  ERASE Frm, Form$, Fld, UBCustRec, UBBillRec, RePrintIdx
  
  IF NOT AbortFlag THEN
    PrintRptFile "Utility Bill Printing ", "UBBILLSR.PRN", 1, RetCode, 1
  END IF
  
ExitRePrintBill:
  EXIT SUB
  
CheckFields:
  FirstBill = Value#(Form$(FrstBFld, 0), ErrCode%)
  LastBill = Value#(Form$(LastBFld, 0), ErrCode%)
  IF FirstBill > LastBill THEN
    SaveScrn TempScrn()
    DisplayUBScrn "BADBILNO"
    WaitForAction
    RestScrn TempScrn()
  ELSE
    ReqFldsOK = True
  END IF
  RETURN
  
FillForm:
  
  LSET Form$(FrstBFld, 0) = QPTrim$(STR$(BillInfoRec(1).FrstBill))
  LSET Form$(LastBFld, 0) = QPTrim$(STR$(BillInfoRec(1).LastBill))
  LSET Form$(BillDFld, 0) = Num2Date$(BillInfoRec(1).BillDate)
  LSET Form$(PastDFld, 0) = Num2Date$(BillInfoRec(1).PastDate)

  LSET Form$(PRDateFld, 0) = Num2Date$(BillInfoRec(1).PRDate)
  LSET Form$(CRDateFld, 0) = Num2Date$(BillInfoRec(1).CRDate)

  LSET Form$(DraftDFld, 0) = Num2Date$(BillInfoRec(1).DrftDate)
  LSET Form$(BillOFld, 0) = BillInfoRec(1).PrnOrder
  LSET Form$(MesgFld, 0) = BillInfoRec(1).MsgLine
  LSET Form$(Desc1Fld, 0) = BillInfoRec(1).Desc1
  LSET Form$(Desc2Fld, 0) = BillInfoRec(1).Desc2
  LSET Form$(Desc3Fld, 0) = BillInfoRec(1).Desc3

  Message$ = QPTrim$(BillInfoRec(1).MsgLine)

  Action = 2
  RETURN
  
END SUB

