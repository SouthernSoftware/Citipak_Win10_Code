DEFINT A-Z

DECLARE SUB CustCutOffListing ()
DECLARE SUB CustomerFlatRates ()
DECLARE SUB WestCaroConsumpList ()
DECLARE FUNCTION Monitor% ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
DECLARE SUB MakeZipCodeIndex (IndexText$)
DECLARE SUB SortServAddress (IndexText$)
DECLARE SUB CustEditMeterEST (CustRec&)
DECLARE SUB CustMessageSystem (RecNo&)
DECLARE SUB ShowWrkOrdHistory (RecNo&)
DECLARE SUB ShowCustConsHist (CustRec&)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB PaymentSumReport ()
DECLARE SUB MailingLabel ()
DECLARE SUB AgedBalanceListing ()
DECLARE SUB CustomerInquiry ()
DECLARE SUB DetailedTransJournal ()
DECLARE SUB MasterBalanceListing ()
DECLARE SUB MasterCustomerListing ()
DECLARE SUB MasterDepositListing ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE FUNCTION GetNumOfCust% ()
DECLARE FUNCTION QPValI% (Numb$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Round# (N#)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE SUB BlockClear ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB FCreate (FileName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB HideCursor ()
DECLARE SUB BlockClear ()
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB ShowCursor ()
DECLARE SUB SaveScrn (array())
DECLARE SUB RestScrn (array())
DECLARE SUB TextCursor (FG%, BG%)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE SUB WaitForAction ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB KillFile (FileName$)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, LocationFlag%)
DECLARE FUNCTION QPStrL$ (LongNumber&)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB MakeSequenceIndex (Text$)

  CONST False = 0, True = NOT False

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'newcust.bi'


   ShowCursor
   MailingLabel
   IF INSTR(COMMAND$, "TEST") THEN
     ClearScrn
     END
   ELSE
     RUN "UBCUSTRE"
   END IF

SUB MailingLabel
  
  SHARED Choice$()

  REDIM Choice$(0 TO 15, 0 TO 2)
  REDIM ToPrint(1 TO 5) AS STRING * 132
  FOR Cnt = 1 TO 5
    LSET ToPrint(Cnt) = ""
  NEXT

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetupLen%
  
  Choice$(0, 0) = "3"
  Choice$(1, 0) = " Customer Name Order   "
  Choice$(2, 0) = " Account Number Order  "
  Choice$(3, 0) = " Location Number Order "
  Choice$(4, 0) = " Service Address Order "
  Choice$(5, 0) = " Sequence Number Order "
  Choice$(5, 0) = " ZipCode Number Order  "

  Choice$(0, 1) = "5"
  didcnt = 0
  FOR RevCnt = 1 TO 15
    Rev$ = QPTrim$(UBSetUpRec(1).Revenues(RevCnt).RevName)
    IF LEN(Rev$) > 0 THEN
      didcnt = didcnt + 1
      Choice$(didcnt, 1) = QPTrim$(STR$(didcnt)) + ") " + Rev$
    ELSE
      EXIT FOR
    END IF
  NEXT

  Choice$(0, 2) = "7"
  Choice$(1, 2) = "Laser 3 Wide"
  
  REDIM OSet(1 TO 4) AS INTEGER

  OSet(1) = 1
  OSet(2) = 35
  OSet(3) = 67

  UsingBook = False
  UsingAcct = False
  UsingName = False
  UsingAddr = False
  
  AbortFlag = False
  PageNo = 0
  
  LibName$ = "UBWRKORD"
  ScrnName$ = "UBPRNLAB"
  
  ' Define Fields
  
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  '--Set screen number to one and display screen
  Scr = 1
  
  BlockClear
  
  LibFile2Scrn LibName$, ScrnName$, MonoCode, Attribute%, ErrCode
  
  ShowCursor
  
  FirstTime = True
  
  Action = 1
  
  FirstTime = True
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = "1"
      LSET Form$(2, 0) = "99"
      LSET Form$(6, 0) = "N"
      LSET Form$(7, 0) = Choice$(1, 2)
      Fld(7).Protected = True
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      Order$ = LEFT$(QPTrim$(Form$(3, 0)), 3)
      SELECT CASE Order$
      CASE "Cus"
        IndexName$ = NameIndexFile
        UsingName = True
        OKFlag = True
      CASE "Acc"
        IndexName$ = ""
        UsingAcct = True
        OKFlag = True
      CASE "Loc"
        IndexName$ = BookIndexFile
        UsingBook = True
        OKFlag = True
      CASE "Ser"
        IndexName$ = TempIndexName
        UsingAddr = True
        OKFlag = True
      CASE "Seq"
        MakeSequenceIndex "Sequence Numbers"
        IndexName$ = TempIndexName
        UsingSeq = True
        OKFlag = True
      CASE "Zip"
        MakeZipCodeIndex "ZipCode"
        IndexName$ = TempIndexName
        UsingZip = True
        OKFlag = True
      CASE ELSE
        GOSUB ShowErrScrn1
        GOTO BadParm
      END SELECT

    CASE ESC
      GOTO ExitMailLabListing
    CASE ELSE
      Done = False
    END SELECT
    
BadParm:
  LOOP UNTIL OKFlag
  
'***************
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  REDIM UBSetup(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetup(1))
  
  CustomerType$ = QPTrim$(Form$(4, 0))

  IF CustomerType$ = "DRFT" THEN
    DraftOnlyFlag = True
  END IF

  FirstBook = Value#(Form$(1, 0), ECode)
  LastBook = Value#(Form$(2, 0), ECode)
  
  IF Form$(6, 0) = "Y" THEN
    IncInactive = True
  END IF
  
  RType = VAL(LEFT$(Form$(5, 0), 1))

  ERASE Frm, Form$, Fld, Choice$
  
  IF UsingAddr THEN
    SortServAddress "Service Address"
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize&(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
  ELSEIF UsingName OR UsingBook OR UsingSeq OR UsingZip THEN
    IdxRecLen = 4               'we are using a long integer
    IdxFileSize& = FileSize(IndexName$)
    IdxNumOfRecs = IdxFileSize& \ IdxRecLen
    REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
    NumOfRecs = IdxNumOfRecs
  ELSE
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  END IF
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen
  
  UBRpt = FREEFILE
  OPEN "UBLABEL.RPT" FOR OUTPUT AS UBRpt
  
  BlockClear
  ShowProcessingScrn "Mailing Labels"
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook OR UsingSeq OR UsingZip THEN
      AcctNumber& = IdxBuff(Cnt).RecNum
    ELSE
      AcctNumber& = Cnt
    END IF
    
    GET UBCust, AcctNumber&, UBCustRec(1)

    IF UBCustRec(1).DelFlag = True THEN
      GOTO NextLabel
    END IF
    
    CustBook = Value#(UBCustRec(1).BOOK, ECode)
    
    IF CustBook < FirstBook OR CustBook > LastBook THEN
      GOTO NextLabel
    END IF
    
    IF (UBCustRec(1).Status = "I") AND (IncInactive = 0) THEN
      GOTO NextLabel
    END IF
    
    IF DraftOnlyFlag THEN
      IF UBCustRec(1).USEDRAFT <> "Y" THEN
        GOTO NextLabel
      END IF
    ELSEIF LEN(CustomerType$) > 0 THEN
      IF UCASE$(QPTrim$(UBCustRec(1).CUSTTYPE)) <> UCASE$(CustomerType$) THEN
        GOTO NextLabel
      END IF
    END IF
    
    IF QPTrim$(UBCustRec(1).CUSTNAME) = "VACANT" THEN
      GOTO NextLabel
    END IF

    IF RType > 0 THEN
      Rev$ = QPTrim$(UBCustRec(1).Serv(RType).RateCode)
      IF LEN(Rev$) > 0 THEN
        GOTO GoodCust
      END IF
      FOR FlatCnt = 1 TO 4
        IF UBCustRec(1).FlatRates(FlatCnt).REVSRC = RType THEN
          GOTO GoodCust
        END IF
      NEXT
      GOTO NextLabel
    END IF

GoodCust:
    CustPCnt = CustPCnt + 1

    Zip$ = LEFT$(UBCustRec(1).ZipCode, 5) + "-" + MID$(UBCustRec(1).ZipCode, 6)
    Zip$ = QPTrim$(Zip$)
    Zip1$ = LEFT$(UBCustRec(1).ZipCode, 5)

GotZip:
      IF CustPCnt = 1 THEN
        'PRINT #UBRpt, "&l6d3e0l99F&k4S"
        PRINT #UBRpt, "&l3E";  'top margin
        PRINT #UBRpt, "&l70P"; 'page length
        PRINT #UBRpt, "&l0L";  'skip perf off
        PRINT #UBRpt, "(s12H";  'char per inch
      END IF
      LabelCnt = LabelCnt + 1
      MID$(ToPrint(1), OSet(LabelCnt)) = "Cust #" + STR$(AcctNumber&)
      MID$(ToPrint(2), OSet(LabelCnt)) = LEFT$(QPTrim$(UBCustRec(1).CUSTNAME), 23)
      MID$(ToPrint(3), OSet(LabelCnt)) = LEFT$(QPTrim$(UBCustRec(1).ADDR1), 23)
      IF LEN(QPTrim$(UBCustRec(1).ADDR2)) > 0 THEN
        MID$(ToPrint(4), OSet(LabelCnt)) = LEFT$(QPTrim$(UBCustRec(1).ADDR2), 23)
        MID$(ToPrint(5), OSet(LabelCnt)) = LEFT$(QPTrim$(UBCustRec(1).CITY), 18) + ", " + UBCustRec(1).STATE + " " + Zip$
      ELSE
        MID$(ToPrint(4), OSet(LabelCnt)) = LEFT$(QPTrim$(UBCustRec(1).CITY), 18) + ", " + UBCustRec(1).STATE + " " + Zip$
      END IF
      IF LabelCnt = 3 THEN
        FOR PCnt = 1 TO 5
          PRINT #UBRpt, ToPrint(PCnt)
          LSET ToPrint(PCnt) = ""
        NEXT
        PRINT #UBRpt,
        LabelCnt = 0
      END IF
    
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    
NextLabel:
    ShowPctComp Cnt, NumOfRecs
    IF CustPCnt >= 30 THEN
      PRINT #UBRpt, CHR$(12);
      CustPCnt = 0
      DidIt = DidIt + 1
      IF DidIt = 2 THEN
        EXIT FOR
      END IF
    END IF
  NEXT
  
  CLOSE
  
  ERASE IdxBuff, UBCustRec, ToPrint
  ERASE UBSetup, Frm, Form$, Fld, OSet
  
  IF NOT AbortFlag THEN
    PrintRptFile "Mailing Labels", "UBLABEL.RPT", 1, RetCode, EntryPoint
  END IF
  
  'KillFile "UBLABEL.RPT"
  
ExitMailLabListing:
  
  EXIT SUB
  
DoAlignLabelMask:

  REDIM TempScrn(0)
  SaveScrn TempScrn()
  Align$ = STRING$(34, "X")
  UBRpt = FREEFILE
  OPEN "UBLABEL.RPT" FOR OUTPUT AS UBRpt
  SELECT CASE LType
  CASE 1
    FOR Cnt = 1 TO 5
      PRINT #UBRpt, Align$
    NEXT
    PRINT #UBRpt,
  CASE 2
    FOR Cnt = 1 TO 5
      PRINT #UBRpt, Align$; TAB(OSet(2)); Align$; TAB(OSet(3)); Align$
    NEXT
    PRINT #UBRpt,

  CASE 3
    FOR Cnt = 1 TO 5
      PRINT #UBRpt, Align$; TAB(OSet(2)); Align$; TAB(OSet(3)); Align$; TAB(OSet(4)); Align$
    NEXT
    PRINT #UBRpt,
  END SELECT
  
  CLOSE UBRpt

  PrintRptFile "Mailing Labels Alignment", "UBLABEL.RPT", 2, RetCode, EntryPoint
  RestScrn TempScrn()
  ERASE TempScrn
  Action = 1
RETURN
  
ShowErrScrn1:
  SELECT CASE ErrCode
  CASE 2
    EText$ = "  Invalid Label Type!"
    Frm(1).FldNo = 6
  CASE ELSE
    EText$ = "Invalid Printing Order!"
    Frm(1).FldNo = 3
  END SELECT


  REDIM TempScrn(0)
  SaveScrn TempScrn()
  CursorOff
  WazzWind 13, 27, 19, 63, 78, 1, True
  QPrintRC EText$, 15, 34, 79
  QPrintRC "Press Any Key To Continue.", 17, 33, 79
  WaitForAction
  RestScrn TempScrn()
  ERASE TempScrn
  Action = 1
  ERASE TempScrn

  RETURN
  
END SUB

SUB SortServAddress (IndexText$)
  ShowProcessingScrn "Creating " + IndexText$ + " Index"
  QPrintRC "    Reading Customer Records     ", 11, 25, -1
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  CustRecLen = LEN(UBCustRec(1))
  
  NumCustRecs = GetNumOfCust%
  
  REDIM ServIndex(1 TO NumCustRecs)  AS UBServiceAddressIndexType
  IndexRecLen = LEN(ServIndex(1))
  
  CHandle = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CHandle LEN = CustRecLen
  FOR Cnt = 1 TO NumCustRecs
    GET CHandle, Cnt, UBCustRec(1)
    ServIndex(Cnt).ServiceAddress = UBCustRec(1).ServAddr
    ServIndex(Cnt).RecNum = Cnt
    ShowPctComp Cnt, NumCustRecs                'show user percentage complete
  NEXT
  
  CLOSE CHandle
  
  QPrintRC "         Sorting Index.        ", 11, 25, -1
  
  SortT ServIndex(1), NumCustRecs, 0, 16, 0, 14
  ' SortT (Elemen, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
  QPrintRC "      Writing Index Records      ", 11, 25, -1
  IHandle = FREEFILE
  IndexName$ = TempIndexName
  FCreate IndexName$
  OPEN IndexName$ FOR RANDOM SHARED AS IHandle LEN = 4
  FOR Cnt = 1 TO NumCustRecs
    CRec& = ServIndex(Cnt).RecNum
    PUT IHandle, Cnt, CRec&
    ShowPctComp Cnt, NumCustRecs                'show user percentage complete
  NEXT
  CLOSE IHandle
  
  ERASE UBCustRec, ServIndex
  
END SUB

