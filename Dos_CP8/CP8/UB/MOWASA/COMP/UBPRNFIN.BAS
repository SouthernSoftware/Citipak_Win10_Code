DECLARE SUB FMakeSequenceIndex (IndexText$)
DEFINT A-Z
DECLARE SUB FMakeMowZipCodeIndex (IndexText$)
DECLARE SUB DisplayPrnScrn (ScrnName$)
DECLARE SUB MakeMowZipCodeIndex (IndexText$)
DECLARE SUB MakeZipCodeIndex (IdxText$)
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
DECLARE SUB FCreate (FileName$)
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
DECLARE FUNCTION GetZipEDigit$ (Zip$)
DECLARE FUNCTION MakeMonth$ (TDate$)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubowner.bi'
  '$INCLUDE: 'ubpinfo.bi'
  '$INCLUDE: 'ubscsprn.bi'

  CONST False = 0, True = NOT False

  PrintUtilBills
  
  IF INSTR(COMMAND$, "TEST") = 0 THEN
    RUN "UBFINBIL"
  ELSE
    HideCursor
    ClearScrn
    END
  END IF

SUB DisplayPrnScrn (ScrnName$)
  LibFile2Scrn "UBPRNBIL", ScrnName$, MonoCode%, Attribute%, ErrorCode%
END SUB

SUB FMakeMowZipCodeIndex (IndexText$)

  REDIM ZipIndex(1 TO 1)   AS MOWZipIndexType
  ShowProcessingScrn "Creating " + IndexText$ + " Index "
  QPrintRC "    Reading Customer Records     ", 11, 25, -1

  REDIM UBCustRec(1)      AS NewUBCustRecType
  CustRecLen = LEN(UBCustRec(1))

  NumOfBillRec = FileSize("UBCUST.DAT") \ CustRecLen

  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = CustRecLen

  REDIM ZipIndex(1 TO 1)   AS MOWZipIndexType
  FOR BCnt = 1 TO NumOfBillRec
    GET CHandle, BCnt, UBCustRec(1)
    IF UBCustRec(1).Status = "F" THEN
      FCnt = FCnt + 1
      REDIM PRESERVE ZipIndex(1 TO FCnt)   AS MOWZipIndexType
      ZipIndex(FCnt).ZIPCODE = UBCustRec(1).ZIPCODE
      ZipIndex(FCnt).RecNum = BCnt
    END IF
    ShowPctComp BCnt, NumOfBillRec              'show user percentage complete
  NEXT
  CLOSE

  QPrintRC "         Sorting Index.        ", 11, 25, -1
  SortT ZipIndex(1), FCnt, 0, 16, 0, 10
  QPrintRC "      Writing Index Records      ", 11, 25, -1

  IHandle = FREEFILE
  OPEN TempIndexName FOR OUTPUT AS IHandle
  CLOSE IHandle

  IHandle = FREEFILE
  OPEN TempIndexName FOR RANDOM SHARED AS IHandle LEN = 4
  FOR Cnt = 1 TO FCnt
    Prec& = ZipIndex(Cnt).RecNum
    PUT IHandle, Cnt, Prec&
    ShowPctComp Cnt, NumOfBillRec               'show user percentage complete
  NEXT
  CLOSE IHandle

  ERASE UBCustRec, ZipIndex

END SUB

SUB FMakeSequenceIndex (IndexText$)

  REDIM SequenceIndex(1 TO 1)  AS UBMowSequenceIndexType
  ShowProcessingScrn "Creating " + IndexText$ + " Index"
  QPrintRC "    Reading Location Records     ", 11, 25, -1
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  CustRecLen = LEN(UBCustRec(1))
  
  NumCustRecs& = FileSize("UBCUST.DAT") \ CustRecLen
  
  IndexRecLen = LEN(SequenceIndex(1))
  
  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = CustRecLen
  FOR Cnt = 1 TO NumCustRecs&
    GET CHandle, Cnt, UBCustRec(1)
    IF UBCustRec(1).Status = "F" THEN
      FCnt = FCnt + 1
      REDIM PRESERVE SequenceIndex(1 TO FCnt)  AS UBMowSequenceIndexType
      SequenceIndex(FCnt).SeqNumber = UBCustRec(1).SEQ
      SequenceIndex(FCnt).RecNum = Cnt
    END IF
    ShowPctComp Cnt, NumCustRecs&               'show user percentage complete
  NEXT
  
  CLOSE CHandle
  
  QPrintRC "         Sorting Index.        ", 11, 25, -1
  
  SortT SequenceIndex(1), FCnt, 0, 8, 0, -2
  
  QPrintRC "      Writing Index Records      ", 11, 25, -1
  
  FCreate TempIndexName
  
  IHandle = FREEFILE
  OPEN TempIndexName FOR RANDOM SHARED AS IHandle LEN = 4
  
  FOR Cnt = 1 TO FCnt
    Prec& = SequenceIndex(Cnt).RecNum
    PUT IHandle, Cnt, Prec&
    ShowPctComp Cnt, FCnt               'show user percentage complete
  NEXT
  CLOSE IHandle
  
  ERASE UBCustRec, SequenceIndex

END SUB

SUB PrintUtilBills

  CrLf$ = CHR$(13) + CHR$(10)

  Fmt10$ = "##########"
  Fmt10a$ = "#######.##"
  Fmt15$ = "############.##"

  OPEN "UBDEPFLG.DAT" FOR RANDOM SHARED AS #1 LEN = 2
  GET #1, , UseDepositFlag
  CLOSE #1
  
  SHARED Choice$()
  
  LPIFlag = False
  
  IF NOT Exist(UBFinBillsFile) THEN
    CursorOff
    BlockClear
    DisplayUBScrn "NON2PRNT"
    WaitForAction
    GOTO ExitPrintBill
  END IF
  
  REDIM BillInfoRec(1) AS PrintBillInfoType
  BillInfoRecLen = LEN(BillInfoRec(1))
  
  REDIM IndexArray(1) AS UBCustIndexRecType
  
  REDIM ScrnArray(0)
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen      'load setup file
  
  FOR ThisRevCnt = 1 TO 15
    IF INSTR(UBSetUpRec(1).Revenues(ThisRevCnt).RevName, "ELECTRIC") THEN
      ElecRev = ThisRevCnt
      EXIT FOR
    END IF
  NEXT
  
  'Section to check for customer modifications
  
  TownName$ = UBSetUpRec(1).UTILNAME
  
  IF INSTR(TownName$, "MOORE") THEN
    MowFlag = True
  END IF
  
  IF UBSetUpRec(1).BANKDFT = "Y" THEN
    UseDraftFlag = True
  END IF
  
  REDIM Choice$(0 TO 6, 0)
  
  Choice$(1, 0) = "Customer Name Order"
  Choice$(2, 0) = "Account Number Order"
  'Choice$(3, 0) = "Location Number Order"
  Choice$(3, 0) = "Postal Carrier Route Order"
  Choice$(4, 0) = "ZipCode Order"
  
  IF UBSetUpRec(1).UseSeq = "Y" THEN
    Choice$(5, 0) = "Sequence Number Order"
  END IF
  
  LibName$ = "UBPRNBIL"
  
  ScrnName$ = "PRNBILF1"
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
  
  FrstBFld = FldNum%("FRSTBILL", Fld())
  LastBFld = FldNum%("LASTBILL", Fld())
  
  BillDFld = FldNum%("BILLDATE", Fld())
  PastDFld = FldNum%("PASTDATE", Fld())
  
  PRDateFld = FldNum%("PRDATE", Fld())
  CRDateFld = FldNum%("CRDATE", Fld())
  
  DraftDFld = FldNum%("DRFTDATE", Fld())
  BillOFld = FldNum%("PRNORDER", Fld())
  DepositFld = FldNum%("APPLYDEP", Fld())
  
  MsgF1 = FldNum%("MSGLINE1", Fld())
  MsgF2 = FldNum%("MSGLINE2", Fld())
  MsgF3 = FldNum%("MSGLINE3", Fld())
  MsgF4 = FldNum%("MSGLINE4", Fld())
  
  Fld(MsgF1).Protected = True
  
  Today$ = DATE$
  
  Action = 1
  Frm(1).StayOnField = True
  
  FirstTime = True
  '--Set screen number to one and display screen
  
  BlockClear
  
  DisplayPrnScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(BillDFld, 0) = Today$
      LSET Form$(PastDFld, 0) = Today$
      LSET Form$(DraftDFld, 0) = Today$
      SELECT CASE UseDepositFlag
      CASE True
        LSET Form$(DepositFld, 0) = "Y"
      CASE False
        LSET Form$(DepositFld, 0) = "N"
      END SELECT
      
      LSET Form$(MsgF1, 0) = "Final Billing"
      
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
      
    CASE EscKey
      AbortFlag = True
      ExitFlag = True
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE MActionRow           'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 29 TO 40           '--Cancel button
          PressButton 27, MActionRow, 29, 40
        CASE 41 TO 53           '--F10 button
          PressButton F10Key, MActionRow, 41, 53
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
  CASE "P", "Z"
    IF BillOrder$ = "P" THEN
      IdxTypeText$ = "Postal Route"
      MakePostalIndex IdxTypeText$
    ELSE
      IdxTypeText$ = "Zip-Code"
      FMakeMowZipCodeIndex IdxTypeText$
    END IF
    IndexName$ = TempIndexName
    OKFlag = True
  CASE "S"
    IdxTypeText$ = "Sequence Number"
    FMakeSequenceIndex IdxTypeText$
    IndexName$ = TempIndexName
    OKFlag = True
  END SELECT
  
  IF AbortFlag GOTO ExitPrintBill
  
  PastDate$ = Form$(PastDFld, 0)
  
  'do bill printing here
  '**************************************************************************
  FinalFlag = True

  REDIM PrintRec(1) AS BillOutRecType
  BillOutRecLen = LEN(PrintRec(1))
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  REDIM UBOwnerRec(1) AS UBOwnerRecType
  UBOwnerRecLen = LEN(UBOwnerRec(1))
  
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
  
  BillInfoRec(1).MsgLine1 = Form$(MsgF1, 0)
  BillInfoRec(1).MsgLine2 = Form$(MsgF2, 0)
  BillInfoRec(1).MsgLine3 = Form$(MsgF3, 0)
  BillInfoRec(1).MsgLine4 = Form$(MsgF4, 0)
  
  IF UsingAcct THEN             'load the index
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  ELSE
    NumOfRecs = FileSize(IndexName$) \ 4
    REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
  END IF
  
  UBBill = FREEFILE
  OPEN UBFinBillsFile FOR RANDOM SHARED AS UBBill LEN = UBBillRecLen
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBOwn = FREEFILE
  OPEN "UBOWNER.DAT" FOR RANDOM SHARED AS UBOwn LEN = UBOwnerRecLen

'add mowasa
  IF MowFlag THEN
    Ext$ = ".MOW"
  ELSE
    Ext$ = ".BTF"
  END IF
  
  SCSFileName$ = LEFT$(Today$, 2) + MID$(Today$, 4, 2) + RIGHT$(Today$, 2) + "F"
  FOR Cnt = 1 TO 9
    ChkName$ = SCSFileName$ + QPTrim$(STR$(Cnt)) + Ext$
    IF Exist(ChkName$) = False THEN
      SCSFileName$ = ChkName$
      EXIT FOR
    END IF
  NEXT

  UBRpt = FREEFILE
  OPEN SCSFileName$ FOR RANDOM SHARED AS UBRpt LEN = BillOutRecLen
  
  IF UseDraftFlag THEN
    UBDraft = FREEFILE
    OPEN DFFileName$ FOR RANDOM SHARED AS UBDraft LEN = UBDraftPayLen
  END IF
  
  ShowProcessingScrn "Creating Transfer Bill File."
  
  '-----------------------------------------
  PrintedCnt = 0
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingAcct THEN
      CustAcctNo& = Cnt
    ELSE
      CustAcctNo& = IndexArray(Cnt).RecNum
    END IF
    
    GET UBCust, CustAcctNo&, UBCustRec(1)
    
    Num2Print = UBCustRec(1).BILLCOPY
    IF Num2Print < 1 THEN Num2Print = 1

    GET UBBill, CustAcctNo&, UBBillRec(1)
    IF UBBillRec(1).ActiveFlag THEN
'121598 Finished adding bill to owner
      IF UBCustRec(1).BillTo = "O" THEN   'if they want to send bill to owner
        'temp copy the owners info to the customers rec
        GET UBOwn, CustAcctNo&, UBOwnerRec(1)
        OName$ = QPTrim$(QPTrim$(UBOwnerRec(1).OwnFName) + " " + QPTrim$(UBOwnerRec(1).OwnLName))
        UBCustRec(1).CUSTNAME = OName$
        UBCustRec(1).ADDR1 = UBOwnerRec(1).ADDR1
        UBCustRec(1).ADDR2 = UBOwnerRec(1).ADDR2
        UBCustRec(1).CITY = UBOwnerRec(1).CITY
        UBCustRec(1).STATE = UBOwnerRec(1).STATE
        UBCustRec(1).ZIPCODE = UBOwnerRec(1).ZIPCODE
      END IF
      
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
        
        IF ApplyDepFlag$ = "Y" THEN
          CDeposit# = UBCustRec(1).DepositAmt
        ELSE
          CDeposit# = 0
        END IF
        
        UBBillRec(1).ApplyDepFlag = ApplyDepFlag$
        
        PUT UBBill, CustAcctNo&, UBBillRec(1)
        
        DidADraftFlag = False
        IF UseDraftFlag AND UBCustRec(1).USEDRAFT = "Y" AND UBCustRec(1).PreNoteFlag THEN
          UBDraftPayRec(1).CustAcctNum = CustAcctNo&
          UBDraftPayRec(1).DraftAmt = UBBillRec(1).TransAmt
          PUT UBDraft, , UBDraftPayRec(1)
          DidADraftFlag = True
        END IF

          WFoundMtr = False
          IFoundMtr = False
          CubicMtr = False
          WCurrRead& = 0
          WPrevRead& = 0
          WUsageAmt& = 0
          ICurrRead& = 0
          IPrevRead& = 0
          IUsageAmt& = 0

          MtrType$ = ""
          FOR mChk = 1 TO 7
            IF UBBillRec(1).MtrTypes(mChk) > 0 THEN
              MtrNumb$ = QPTrim$(UBCustRec(1).LocMeters(mChk).MTRNUM)
              IF LEN(MtrNumb$) = 0 THEN
                MtrNumb$ = "???"
              END IF
              MtrTyp = UBBillRec(1).MtrTypes(mChk)
              SELECT CASE MtrTyp
              CASE 1
                IFoundMtr = True
                ICurrRead& = UBBillRec(1).CurRead(mChk)
                IPrevRead& = UBBillRec(1).PrevRead(mChk)
                IUsageAmt& = ICurrRead& - IPrevRead&
                IF IUsageAmt& < 0 THEN
                  MaxMeterAmt& = 10& ^ (LEN(STR$(IPrevRead&)) - 1)
                  IUsageAmt& = (MaxMeterAmt& - IPrevRead&) + ICurrRead&
                END IF
                IF UBCustRec(1).LocMeters(mChk).MtrUnit = "C" THEN
                  CubicMtr = True
                  IUsageAmt& = IUsageAmt& * 7.481
                END IF
                IF MtrTyp = 1 THEN
                  MtrType$ = "W"
                ELSEIF MtrTyp = 2 THEN
                  MtrType$ = "S"
                ELSEIF MtrTyp = 3 THEN
                  MtrType$ = "C"
                ELSEIF MtrType = 7 THEN
                  MtrType$ = "T"
                END IF
              '  EXIT FOR
              CASE 2, 3, 7
                WFoundMtr = True
                WCurrRead& = UBBillRec(1).CurRead(mChk)
                WPrevRead& = UBBillRec(1).PrevRead(mChk)
                WUsageAmt& = WCurrRead& - WPrevRead&
                IF WUsageAmt& < 0 THEN
                  MaxMeterAmt& = 10& ^ (LEN(STR$(WPrevRead&)) - 1)
                  WUsageAmt& = (MaxMeterAmt& - WPrevRead&) + WCurrRead&
                END IF
                IF UBCustRec(1).LocMeters(mChk).MtrUnit = "C" THEN
                  CubicMtr = True
                  WUsageAmt& = WUsageAmt& * 7.481
                END IF
                IF MtrTyp = 1 THEN
                  MtrType$ = "W"
                ELSEIF MtrTyp = 2 THEN
                  MtrType$ = "S"
                ELSEIF MtrTyp = 3 THEN
                  MtrType$ = "C"
                ELSEIF MtrType = 7 THEN
                  MtrType$ = "T"
                END IF
               ' EXIT FOR
              END SELECT
            END IF
          NEXT

          IF WFoundMtr = False AND IFoundMtr = False THEN
            'if no metered services then adjust read dates to billdate
            'and billdate - 30
            DateRead$ = Num2Date$(UBBillRec(1).BillDate)
            PrevDate$ = Num2Date$(UBBillRec(1).BillDate - 30)
            DaysINRead = 30
          END IF

'**********************************************************
          Previous# = Round#(UBCustRec(1).PrevBalance + UBCustRec(1).CurrBalance)
          TotalAmt# = Round#(Previous# + UBBillRec(1).TransAmt)

          IF FinalFlag AND CDeposit# THEN
            TotalAmt# = Round#(TotalAmt# - UBCustRec(1).DepositAmt)
          END IF

          IF TotalAmt# > 0 THEN
            TenPct# = 0
            IF DaysINRead < 1 THEN DaysINRead = 1
            AvgCst# = Round#(TotalAmt# / DaysINRead)
          ELSE
            TenPct# = 0
            AvgCst# = 0
          END IF

          REDIM PrintRec(1) AS BillOutRecType

          PrintRec(1).AcctNo = FUsing(STR$(CustAcctNo&), "########")
          PrintRec(1).LocationNum = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
          RSET PrintRec(1).CUSTNAME = QPTrim$(UBCustRec(1).CUSTNAME)
          RSET PrintRec(1).ADDR1 = QPTrim$(UBCustRec(1).ADDR1)
          RSET PrintRec(1).ADDR2 = QPTrim$(UBCustRec(1).ADDR2)
          RSET PrintRec(1).SERVADDR = QPTrim$(UBCustRec(1).SERVADDR)
          RSET PrintRec(1).CITY = QPTrim$(UBCustRec(1).CITY)
          RSET PrintRec(1).STATE = QPTrim$(UBCustRec(1).STATE)
          RSET PrintRec(1).ZIPCODE = QPTrim$(UBCustRec(1).ZIPCODE)
          PrintRec(1).BillType = "F"

          PrintRec(1).DepAppAmt = FUsing(STR$(CDeposit#), Fmt10a$)
          PrintRec(1).MtrType = MtrType$
          IF CubicMtr THEN
            PrintRec(1).MtrUnit = "C"
          ELSE
            PrintRec(1).MtrUnit = "G"
          END IF
          PrintRec(1).PrevDue = FUsing(STR$(Previous#), Fmt15$)
          PrintRec(1).CurrDue = FUsing(STR$(UBBillRec(1).TransAmt), Fmt15$)
          PrintRec(1).TotalDue = FUsing(STR$(TotalAmt#), Fmt15$)
          PrintRec(1).CurrDate = DateRead$
          PrintRec(1).PrevDate = PrevDate$


          IF WFoundMtr = False THEN
            PrintRec(1).CurrRead = ""
            PrintRec(1).PrevRead = ""
            PrintRec(1).Consump = ""
            PrintRec(1).ServDays = ""
          ELSE
            PrintRec(1).CurrRead = FUsing(STR$(WCurrRead&), Fmt10$)
            PrintRec(1).PrevRead = FUsing(STR$(WPrevRead&), Fmt10$)
            PrintRec(1).Consump = FUsing(STR$(WUsageAmt&), Fmt10$)
            PrintRec(1).ServDays = FUsing(STR$(DaysINRead), "####")
          END IF
          IF IFoundMtr = False THEN
            PrintRec(1).ICurrRead = ""
            PrintRec(1).IPrevRead = ""
            PrintRec(1).IConsump = ""
            PrintRec(1).IServDays = ""
          ELSE
            PrintRec(1).ICurrRead = FUsing(STR$(ICurrRead&), Fmt10$)
            PrintRec(1).IPrevRead = FUsing(STR$(IPrevRead&), Fmt10$)
            PrintRec(1).IConsump = FUsing(STR$(IUsageAmt&), Fmt10$)
            PrintRec(1).IServDays = FUsing(STR$(DaysINRead), "####")
          END IF


'******************************************************
          FOR Serv = 1 TO 15
            PrintRec(1).ServInfo(Serv).ServText = QPTrim$(UBSetUpRec(1).Revenues(Serv).RevName)
            PrintRec(1).ServInfo(Serv).ServAmt = FUsing(STR$(UBBillRec(1).RevAmt(Serv)), Fmt10a$)
          NEXT
          PrintRec(1).BillDate = BillDate$
          PrintRec(1).PastDueDate = PastDueDate$
          IF DidADraftFlag THEN
            PrintRec(1).DraftDate = DraftDate$
          ELSE
            PrintRec(1).DraftDate = ""
          END IF
          RSET PrintRec(1).MsgLine1 = Message$
          RSET PrintRec(1).MsgLine2 = Msg2$
          RSET PrintRec(1).MsgLine3 = Msg3$
          RSET PrintRec(1).MsgLine4 = Msg4$
'unrem
'          LSET PrintRec(1).MtrNumb = MtrNumb$

          PrintRec(1).CrLf = CrLf$
          PUT #UBRpt, , PrintRec(1)

'***************************************************
        'FOR BillCopies = 1 TO Num2Print
        '  ''$INCLUDE: 'UBMOWASA.BI'        'Mowasa
        'NEXT
      END IF
    END IF
    
SkipEm:
    
    'IF AskAbandonPrint% THEN
    '  AbortFlag = True
    '  EXIT FOR
    'END IF
    ShowPctComp Cnt, NumOfRecs
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
  
  BillInfoRec(1).MsgLine1 = Form$(MsgF1, 0)
  BillInfoRec(1).MsgLine2 = Form$(MsgF2, 0)
  BillInfoRec(1).MsgLine3 = Form$(MsgF3, 0)
  BillInfoRec(1).MsgLine4 = Form$(MsgF4, 0)
'trType
  UBFile = FREEFILE
  OPEN "UBPINFOF.DAT" FOR RANDOM AS #UBFile LEN = BillInfoRecLen
  PUT #UBFile, 1, BillInfoRec(1)
  CLOSE
  
  ERASE Frm, Form$, Fld, UBCustRec, UBBillRec, BillInfoRec, IndexArray
  
  IF NOT AbortFlag THEN
'    PrintRptFile "Final Utility Bill Printing ", SCSFileName$, 1, RetCode, 1
    BlockClear
    DisplayUBScrn "UPDATEOK"
    QPrintRC "File Name:  " + SCSFileName$, 10, 29, 112
    WaitForAction
  END IF
  
  '**************************************************************************
  'end bill printing
  GOTO ExitPrintBill:
  '******************
  
  '******************
CheckReqFields:
  
  FBillNO& = VAL(Form$(FrstBFld, 0))
  LBillNO& = VAL(Form$(LastBFld, 0))
  BillDate = Date2Num(Form$(BillDFld, 0))
  
  BillDate$ = Num2Date$(BillDate)
  
  '  BilledDate$ = Form$(BillDFld, 0)
  PastDate = Date2Num(Form$(PastDFld, 0))
  
  DraftDate = Date2Num(Form$(DraftDFld, 0))
  DueDate$ = Form$(PastDFld, 0)
  BillOrder$ = QPTrim$(LEFT$(Form$(BillOFld, 0), 1))
  ApplyDepFlag$ = Form$(DepositFld, 0)
  
  Message$ = QPTrim$(Form$(MsgF1, 0))
  Msg2$ = QPTrim$(Form$(MsgF2, 0))
  Msg3$ = QPTrim$(Form$(MsgF3, 0))
  Msg4$ = QPTrim$(Form$(MsgF4, 0))
  
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
    Frm(1).FldNo = 6
    '  ELSEIF ((FBillNO& <= 0) OR (LBillNO& < FBillNO&)) THEN
    '    SaveScrn ScrnArray()
    '    DisplayUBScrn "NOSTRTNO"
    '    WaitForAction
    '    RestScrn ScrnArray()
    '    Frm(1).FldNo = FrstBFld
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

