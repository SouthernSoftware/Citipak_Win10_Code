DEFINT A-Z
DECLARE SUB MakeMowZipCodeIndex (IndexText$)
DECLARE SUB MakePostalIndex (IndexText$)
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
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)
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
    HideCursor
    'ClearScrn
    END
  ELSE
    RUN "UBMTREAD"
  END IF

SUB MailingLabel
  
  SHARED Choice$()
  REDIM Choice$(0 TO 7, 0 TO 1)
  
  REDIM ToPrint(1 TO 5) AS STRING * 132
  
  FOR Cnt = 1 TO 5
    LSET ToPrint(Cnt) = ""
  NEXT
  
  REDIM UBSetup(1) AS UBSetupRecType
  UBSetUpRecLen = LEN(UBSetup(1))
  LoadUBSetUpFile UBSetup(), UBSetUpRecLen
  HighVar# = UBSetup(1).HighRead
  
  Choice$(0, 0) = "4"
  Choice$(1, 0) = " Customer Name Order   "
  Choice$(2, 0) = " Account Number Order  "
  Choice$(3, 0) = " Location Number Order "
  Choice$(4, 0) = " Service Address Order "
  Choice$(5, 0) = " Sequence Number Order "
  Choice$(6, 0) = " ZipCode Number Order  "
  Choice$(7, 0) = " Postal Carrier Order  "

  
  Choice$(0, 1) = "5"
  Choice$(1, 1) = "1)  1 x 3" + CHR$(171) + "  1 Wide"
  Choice$(2, 1) = "2)  1 x 3" + CHR$(171) + "  3 Wide"
  Choice$(3, 1) = "3)  1 x 3" + CHR$(171) + "  4 Wide"
  

  REDIM OSet(1 TO 4) AS INTEGER
  
  OSet(1) = 1
  OSet(2) = 37
  OSet(3) = 74
  OSet(4) = 110
  
  UsingBook = False
  UsingAcct = False
  UsingName = False
  UsingAddr = False
  
  AbortFlag = False
  PageNo = 0
  
  LibName$ = "UBHH"
  ScrnName$ = "DBLCONLB"
  
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
      LSET Form$(3, 0) = QPTrim$(STR$(HighVar#))
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F5KEY
      LType = VAL(LEFT$(Form$(5, 0), 1))
      IF LType = 0 THEN
        ErrCode = 2
        GOSUB ShowErrScrn1
      ELSE
        GOSUB DoAlignLabelMask
      END IF
      
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
        MakeMowZipCodeIndex "ZipCode"
        IndexName$ = TempIndexName
        UsingZip = True
        OKFlag = True
      CASE "Pos"
        MakePostalIndex "Postal Route"
        IndexName$ = TempIndexName
        UsingPos = True
        OKFlag = True
      CASE ELSE
        GOSUB ShowErrScrn1
        GOTO BadParm
      END SELECT
      
      HighVar# = Value#(Form$(3, 0), ECode)
      IF HighVar# <= 0 THEN
        ErrCode = 1
        GOSUB ShowErrScrn1
        OKFlag = False
        GOTO BadParm
      END IF
      
      LType = VAL(LEFT$(Form$(5, 0), 1))
      IF LType = 0 THEN
        ErrCode = 2
        GOSUB ShowErrScrn1
        OKFlag = False
        GOTO BadParm
      END IF
      
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
  
  FirstBook = Value#(Form$(1, 0), ECode)
  LastBook = Value#(Form$(2, 0), ECode)
  
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
  ShowProcessingScrn "High Reading Mailing Labels"
  
  FOR Cnt = 1 TO NumOfRecs
    IF UsingName OR UsingBook OR UsingSeq OR UsingZip THEN
      AcctNumber& = IdxBuff(Cnt).RecNum
    ELSE
      AcctNumber& = Cnt
    END IF
    
    GET UBCust, AcctNumber&, UBCustRec(1)
    
    IF UBCustRec(1).DelFlag = True THEN
      GOTO SkipThisCust
    END IF
    
    CustBook = Value#(UBCustRec(1).BOOK, ECode)
    
    IF CustBook < FirstBook OR CustBook > LastBook THEN
      GOTO SkipThisCust
    END IF
    
    IF UBCustRec(1).Status = "I" THEN
      GOTO SkipThisCust
    END IF
    
    IF QPTrim$(UBCustRec(1).CUSTNAME) = "VACANT" THEN
      GOTO SkipThisCust
    END IF
    '********************************
    PrintEm = False
    FOR MCnt = 1 TO 7
      IF LEN(UBCustRec(1).LocMeters(MCnt).MtrType) = 0 THEN
        GOTO SkipThisMeter:
      END IF
      IF (UBCustRec(1).LocMeters(MCnt).CurRead < 0) OR (UBCustRec(1).LocMeters(MCnt).PrevRead < 0) THEN
        GOTO SkipThisMeter:
      END IF
      MeterConsp& = UBCustRec(1).LocMeters(MCnt).CurRead - UBCustRec(1).LocMeters(MCnt).PrevRead
      IF MeterConsp& < 0 THEN
        MaxMeterAmt& = 10& ^ (LEN(STR$(UBCustRec(1).LocMeters(MCnt).PrevRead)) - 1)
        MeterConsp& = (MaxMeterAmt& - UBCustRec(1).LocMeters(MCnt).PrevRead) + UBCustRec(1).LocMeters(MCnt).CurRead
      END IF
      'convert units from cubic feet to gallons here
      '  MeterConsp& = MeterConsp& * 7.481
      AvgUse# = UBCustRec(1).LocMeters(MCnt).AvgUse
      
      IF AvgUse# < 0 THEN
        AvgUse# = 1
      END IF
      HiAvg# = AvgUse# * (HighVar# / 100)
      
      IF MeterConsp& > HiAvg# THEN
        PrintEm = True
        EXIT FOR
      END IF
SkipThisMeter:
    NEXT
    '********************************
    
    IF PrintEm THEN
      
      Zip$ = LEFT$(UBCustRec(1).ZIPCODE, 5) + "-" + MID$(UBCustRec(1).ZIPCODE, 6)
      Zip$ = QPTrim$(Zip$)
      
      SELECT CASE LType
      CASE 1
        PRINT #UBRpt, "Cust #" + STR$(AcctNumber&)
        PRINT #UBRpt, LEFT$(QPTrim$(UBCustRec(1).CUSTNAME), 23)
        PRINT #UBRpt, LEFT$(QPTrim$(UBCustRec(1).ADDR1), 23)
        IF LEN(QPTrim$(UBCustRec(1).ADDR2)) > 0 THEN
          PRINT #UBRpt, LEFT$(QPTrim$(UBCustRec(1).ADDR2), 23)
          PRINT #UBRpt, LEFT$(QPTrim$(UBCustRec(1).CITY), 18) + ", " + UBCustRec(1).STATE + " " + Zip$
        ELSE
          PRINT #UBRpt, LEFT$(QPTrim$(UBCustRec(1).CITY), 18) + ", " + UBCustRec(1).STATE + " " + Zip$
          PRINT #UBRpt,
        END IF
        PRINT #UBRpt,
        didcnt = didcnt + 1
      CASE 2

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
        IF LabelCnt = 4 THEN
          FOR PCnt = 1 TO 5
            'LPRINT QPTrim$(ToPrint(PCnt))
            PRINT #UBRpt, ToPrint(PCnt)
            LSET ToPrint(PCnt) = ""
          NEXT
          PRINT #UBRpt,
          LabelCnt = 0
        END IF

      END SELECT
      
      IF AskAbandonPrint% THEN
        AbortFlag = True
        EXIT FOR
      END IF
      
SkipThisCust:
      ShowPctComp Cnt, NumOfRecs
      'IF didcnt > 4 THEN EXIT FOR
    END IF
  NEXT
  
  IF LType = 2 THEN
    IF LabelCnt > 0 THEN
      FOR PCnt = 1 TO 5
        PRINT #UBRpt, QPTrim$(ToPrint(PCnt))
      NEXT
      PRINT #UBRpt,
    END IF
  END IF
  
  PRINT #UBRpt, CHR$(12);
  CLOSE
  
  ERASE IdxBuff, UBCustRec, ToPrint
  ERASE UBSetup, Frm, Form$, Fld, OSet
  
  IF NOT AbortFlag THEN
    PrintRptFile "High Reading Mailing Labels", "UBLABEL.RPT", 1, RetCode, EntryPoint
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
      PRINT #UBRpt, Align$; TAB(OSet(2)); Align$; TAB(OSet(3)); Align$; TAB(OSet(4)); Align$
    NEXT
    PRINT #UBRpt,
  END SELECT
  
  CLOSE UBRpt
  
  PrintRptFile "Mailing Labels Alignment", "UBLABEL.RPT", 1, RetCode, EntryPoint
  RestScrn TempScrn()
  ERASE TempScrn
  Action = 1
  RETURN
  
ShowErrScrn1:
  SELECT CASE ErrCode
  CASE 1
    EText$ = "  Invalid Percentage!"
    Frm(1).FldNo = 3
    
  CASE 2
    EText$ = "  Invalid Label Type!"
    Frm(1).FldNo = 5
  CASE ELSE
    EText$ = "Invalid Printing Order!"
    Frm(1).FldNo = 4
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

SUB orgMailingLabel
  
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
  Choice$(6, 0) = " ZipCode Number Order  "
  Choice$(7, 0) = " Postal Carrier Order  "

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
  Choice$(1, 2) = "1)  1 x 3" + CHR$(171) + "  1 Wide"
  Choice$(2, 2) = "2)  1 x 3" + CHR$(171) + "  3 Wide"
  Choice$(3, 2) = "3)  1 x 3" + CHR$(171) + "  4 Wide"
  
  REDIM OSet(1 TO 4) AS INTEGER

  OSet(1) = 1
  OSet(2) = 37
  OSet(3) = 74
  OSet(4) = 110

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
      Action = 1
    END IF
    
    SELECT CASE Frm(1).KeyCode
    CASE F5KEY
      LType = VAL(LEFT$(Form$(7, 0), 1))
      IF LType = 0 THEN
        ErrCode = 2
        GOSUB oErrScrn1
      ELSE
        GOSUB oDoAlignMask
      END IF

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
        MakeMowZipCodeIndex "ZipCode"
        IndexName$ = TempIndexName
        UsingZip = True
        OKFlag = True
      CASE "Pos"
        MakePostalIndex "Postal Route"
        IndexName$ = TempIndexName
        UsingPos = True
        OKFlag = True
      CASE ELSE
        GOSUB oErrScrn1
        GOTO oBadParm
      END SELECT

      LType = VAL(LEFT$(Form$(7, 0), 1))
      IF LType = 0 THEN
        ErrCode = 2
        GOSUB oErrScrn1
        OKFlag = False
        GOTO oBadParm
      END IF
    CASE ESC
      GOTO oExitMailLabListing
    CASE ELSE
      Done = False
    END SELECT
    
oBadParm:
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
  ELSEIF UsingName OR UsingBook OR UsingSeq OR UsingZip OR UsingPos THEN
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
    IF UsingName OR UsingBook OR UsingSeq OR UsingZip OR UsingPos THEN
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

    Zip$ = LEFT$(UBCustRec(1).ZIPCODE, 5) + "-" + MID$(UBCustRec(1).ZIPCODE, 6)
    Zip$ = QPTrim$(Zip$)

    Zip1$ = LEFT$(UBCustRec(1).ZIPCODE, 5)
    IF CustPCnt = 1 THEN
      ZipCnt = ZipCnt + 1
      REDIM PRESERVE PrnArray(1 TO ZipCnt) AS UBPostalIndexType
      PrnArray(ZipCnt).ZIPCODE = Zip1$
      PrnArray(ZipCnt).RecNum = 1
    ELSE
      FOR ZZ = 1 TO ZipCnt
        IF INSTR(PrnArray(ZZ).ZIPCODE, Zip1$) > 0 THEN
          PrnArray(ZZ).RecNum = PrnArray(ZZ).RecNum + 1
          GOTO GotZip
        END IF
      NEXT
      ZipCnt = ZipCnt + 1
      REDIM PRESERVE PrnArray(1 TO ZipCnt) AS UBPostalIndexType
      PrnArray(ZipCnt).ZIPCODE = Zip1$
      PrnArray(ZipCnt).RecNum = 1
    END IF

GotZip:
    SELECT CASE LType
    CASE 1
      PRINT #UBRpt, "Cust #" + STR$(AcctNumber&)
      PRINT #UBRpt, LEFT$(QPTrim$(UBCustRec(1).CUSTNAME), 23)
      PRINT #UBRpt, LEFT$(QPTrim$(UBCustRec(1).ADDR1), 23)
      IF LEN(QPTrim$(UBCustRec(1).ADDR2)) > 0 THEN
        PRINT #UBRpt, LEFT$(QPTrim$(UBCustRec(1).ADDR2), 23)
        PRINT #UBRpt, LEFT$(QPTrim$(UBCustRec(1).CITY), 18) + ", " + UBCustRec(1).STATE + " " + Zip$
      ELSE
        PRINT #UBRpt, LEFT$(QPTrim$(UBCustRec(1).CITY), 18) + ", " + UBCustRec(1).STATE + " " + Zip$
        PRINT #UBRpt,
      END IF
      PRINT #UBRpt,
      didcnt = didcnt + 1
    CASE 2
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
          'LPRINT QPTrim$(ToPrint(PCnt))
          PRINT #UBRpt, ToPrint(PCnt)
          LSET ToPrint(PCnt) = ""
        NEXT
        PRINT #UBRpt,
        LabelCnt = 0
      END IF
    
    CASE 3
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
      IF LabelCnt = 4 THEN
        FOR PCnt = 1 TO 5
          'LPRINT QPTrim$(ToPrint(PCnt))
          PRINT #UBRpt, ToPrint(PCnt)
          LSET ToPrint(PCnt) = ""
        NEXT
        PRINT #UBRpt,
        LabelCnt = 0
      END IF
    
    END SELECT

    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    
NextLabel:
    ShowPctComp Cnt, NumOfRecs
    'IF CustPCnt > 60 THEN EXIT FOR
  NEXT
  
  IF LType = 2 OR LType = 3 THEN
    IF LabelCnt > 0 THEN
      FOR PCnt = 1 TO 5
        PRINT #UBRpt, QPTrim$(ToPrint(PCnt))
      NEXT
      PRINT #UBRpt,
    END IF
  END IF
  PCnt = 0

  SortT PrnArray(1), ZipCnt, 0, 16, 0, 14
  FOR Cnt = 1 TO ZipCnt
    'PCnt = PCnt + 1
    PRINT #UBRpt, PrnArray(Cnt).ZIPCODE; TAB(40); PrnArray(Cnt).RecNum
    'IF PCnt = 5 THEN
    '  PRINT #UBRpt,
    '  PCnt = 0
    'END IF
  NEXT
  LSET ToPrint(1) = "Total:" + STR$(CustPCnt)
  PRINT #UBRpt, ToPrint(1)
  PRINT #UBRpt, CHR$(12);
  CLOSE
  
  ERASE IdxBuff, UBCustRec, ToPrint
  ERASE UBSetup, Frm, Form$, Fld, OSet
  
  IF NOT AbortFlag THEN
    PrintRptFile "Mailing Labels", "UBLABEL.RPT", 1, RetCode, EntryPoint
  END IF
  
  'KillFile "UBLABEL.RPT"
  
oExitMailLabListing:
  
  EXIT SUB
  
oDoAlignMask:

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

  PrintRptFile "Mailing Labels Alignment", "UBLABEL.RPT", 1, RetCode, EntryPoint
  RestScrn TempScrn()
  ERASE TempScrn
  Action = 1
RETURN
  
oErrScrn1:
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

