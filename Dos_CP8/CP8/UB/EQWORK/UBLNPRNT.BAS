DEFINT A-Z
DECLARE FUNCTION MakeMonth$ (TDate$)
DECLARE FUNCTION MakeDay$ (DayNum%)
DECLARE FUNCTION MakeLongDate$ (PDate$)
DECLARE SUB PrintLateNotices (NoticeInfo AS ANY)
DECLARE FUNCTION GetNoticeInfo% (NoticeInfo AS ANY)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION Num2Day% (DayNum%)
DECLARE FUNCTION Date2Num% (TDate$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (DateNum%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (LongNum$)
DECLARE FUNCTION Round# (n#)
DECLARE SUB BlockClear ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB WaitForAction ()
DECLARE FUNCTION WEnvTest ()
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB NoticeMailLabel (NoticeInfo AS ANY)
DECLARE SUB MakeZipCodeIndex (IndexText$)
DECLARE SUB MakeMowZipCodeIndex (IndexText$)

TYPE NoticeInfoType
  FromBC        AS INTEGER
  ThruBC        AS INTEGER
  NoticeDate    AS INTEGER         '1
  PayByDate     AS INTEGER         '2
  'FromDate      AS INTEGER         '3
  'TODate        AS INTEGER         '4
  MinBalance    AS DOUBLE          '5
  BalanceType   AS INTEGER         '6
  PrnOrder      AS INTEGER         '7
  UseAFlag      AS INTEGER         '8
  MsgLine1      AS STRING * 25
  MsgLine2      AS STRING * 25
  MsgLine3      AS STRING * 25
  MsgLine4      AS STRING * 25
  PrnCnt        AS INTEGER
END TYPE

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'newcust.bi'

  DIM NoticeInfo AS NoticeInfoType

  CONST False = 0, True = NOT False
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 3)

  MChoice$(1) = "Print Utility Billing Late Notices"
  MChoice$(2) = "Print Late Notice Mailing Labels"
  MChoice$(3) = "Exit to OS"

  MaxLen = 0    'Set menu width to zero
  BoxBot = 15   'limit the box length to go no lower than line
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

    TitleBox 2, Col, MaxLen + 3, "Late Notice Processing", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN
      Choice = 0
      ExitFlag = True
      EXIT DO
    END IF

    SELECT CASE Choice
    CASE 1
      IF GetNoticeInfo(NoticeInfo) THEN
        PrintLateNotices NoticeInfo
      END IF
    CASE 2
      NoticeMailLabel NoticeInfo
    CASE 3
      HideCursor
      ClearScrn
      END
    END SELECT

  LOOP UNTIL ExitFlag

ExitPrint:
  
  IF INSTR(COMMAND$, "TEST") THEN
    'ClearScrn
    HideCursor
    END
  ELSE
    RUN "UBBILLIN"
  END IF

FUNCTION GetNoticeInfo% (NoticeInfo AS NoticeInfoType)

  Fill$ = SPACE$(25)

  REDIM ScrnArray(0)

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen  'load setup file
  TownName$ = UBSetUpRec(1).UTILNAME

  IF UBSetUpRec(1).BILLCYCL = "Y" THEN
    UseCycle = True
  END IF
  
  LibName$ = "UB"
  ScrnName$ = "UBPNOTE"
  NumScrns = 1
  
  '--define the multi-choice fields
  SHARED Choice$()
  REDIM Choice$(4, 1)
  
  Choice$(0, 0) = "6"
  Choice$(1, 0) = "Current Balance"
  Choice$(2, 0) = "Previous Balance"
  Choice$(3, 0) = "Total Balance"

  Choice$(0, 1) = "7"
  Choice$(1, 1) = "Customer Name Order   "
  Choice$(2, 1) = "Account Number Order  "
  Choice$(3, 1) = "Location Number Order "
  Choice$(4, 1) = "Zipcode Order "

  '--Initialize the form name array
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  '--Set screen number to one and display screen
  Frm(1).StayOnField = True

  Scr = 1
  BlockClear
  DisplayUBScrn ScrnName$

  FirstTime = True
  ShowCursor
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      LSET Form$(1, 0) = "0"
      LSET Form$(2, 0) = "99"
      LSET Form$(3, 0) = DATE$
      LSET Form$(4, 0) = Form$(3, 0)
      LSET Form$(5, 0) = "0"
      LSET Form$(8, 0) = "Y"
      IF UseCycle THEN
        QPrintRC "From Cycle", 6, 25, -1
        QPrintRC "Thru Cycle", 7, 25, -1
      ELSE
        QPrintRC "From Book", 6, 26, -1
        QPrintRC "Thru Book", 7, 26, -1
      END IF
      IF WEnvTest THEN
        LSET Form$(6, 0) = Choice$(3, 0)
        LSET Form$(7, 0) = Choice$(2, 1)
      END IF
      Action = 1
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F5KEY
      SaveScrn ScrnArray()
      'Section to check for customer modifications
      CursorOff
      IndianTrail = INSTR(TownName$, "INDIAN T")
      IF IndianTrail THEN
        PrintRptFile " ", "UBLALIGN.DAT", 2, RetCode, 4
      ELSE
        PrintRptFile " ", "UBLALIGN.DAT", 1, RetCode, 4
      END IF
      RestScrn ScrnArray()
    CASE -68    'F10Key
      SaveScrn ScrnArray()
      GOSUB CheckDates
      IF NOT OK2Print THEN
        Action = 1
      ELSE
        GetNoticeInfo = True
        ExitFlag = True
      END IF
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 20   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 21 TO 33           '--cancel button
          PressButton 27, 20, 21, 33
        CASE 34 TO 48           '--Alignment
          PressButton F5KEY, 20, 34, 48
        CASE 49 TO 60           '--Save Button
          PressButton -68, 20, 49, 60
        END SELECT
      END SELECT                'row
    END IF
  '--Check screen page
  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag
  
  IF Frm(1).KeyCode = 27 THEN
    GetNoticeInfo% = False
    GOTO ExitEdit
  END IF
  
ExitEdit:
  HideCursor
EXIT FUNCTION
  
CheckDates:
  OK2Print = False
  BadDate = False
  BadBC = False

'011599 Added book/cycle for
  FromBC = Value(Form$(1, 0), ECode)
  ThruBC = Value(Form$(2, 0), ECode)
  NDate = Date2Num(Form$(3, 0))
  PDate = Date2Num(Form$(4, 0))

  IF NDate < 0 THEN
    BadDate = True
    Frm(1).FldNo = 3
  ELSEIF PDate < 0 THEN
    BadDate = True
    Frm(1).FldNo = 4
  END IF

  IF BadDate THEN
    OK = MsgBox%("UB", "UBPBDATE")
    RestScrn ScrnArray()
    GOTO CheckReturn
  END IF

  IF FromBC > ThruBC THEN
    DisplayUBScrn "ERRSCRN1"
    QPrintRC "Invalid Field Information!", 10, 27, -1
    QPrintRC "Please Correct and try Again.", 12, 26, -1
    Frm(1).FldNo = 1
    WaitForAction
    RestScrn ScrnArray()
    GOTO CheckReturn
  END IF

  IF LEN(QPTrim$(Form$(5, 0))) = 0 THEN
    GOSUB NoBlankErr
    RestScrn ScrnArray()
    Frm(1).FldNo = 5
    GOTO CheckReturn
  ELSE
    NoticeInfo.MinBalance# = Value#(Form$(5, 0), ECode)
    MinBalance# = Value#(Form$(5, 0), ECode)
  END IF
  
  SELECT CASE LEFT$(Form$(6, 0), 1)
  CASE "C"
    NoticeInfo.BalanceType = 1
  CASE "P"
    NoticeInfo.BalanceType = 2
  CASE "T"
    NoticeInfo.BalanceType = 3
  CASE ELSE
    OK = MsgBox%("UB", "UBPBTYPE")
    RestScrn ScrnArray()
    Frm(1).FldNo = 6
    GOTO CheckReturn
  END SELECT

  SELECT CASE LEFT$(Form$(7, 0), 1)
  CASE "C"
    NoticeInfo.PrnOrder = 1
  CASE "A"
    NoticeInfo.PrnOrder = 2
  CASE "L"
    NoticeInfo.PrnOrder = 3
  CASE "Z"
    NoticeInfo.PrnOrder = 4
  CASE ELSE
    OK = MsgBox%("UB", "NOBORDER")
    RestScrn ScrnArray()
    Frm(1).FldNo = 7
    GOTO CheckReturn
  END SELECT

  IF LEN(QPTrim$(Form$(8, 0))) = 0 THEN
    GOSUB NoBlankErr
    RestScrn ScrnArray()
    Frm(1).FldNo = 8
    GOTO CheckReturn
  ELSEIF Form$(8, 0) = "Y" THEN
    NoticeInfo.UseAFlag = True
  ELSE
    NoticeInfo.UseAFlag = False
  END IF

  NoticeInfo.FromBC = FromBC
  NoticeInfo.ThruBC = ThruBC
  NoticeInfo.NoticeDate = NDate
  NoticeInfo.PayByDate = PDate
  NoticeInfo.MsgLine1 = QPTrim$(Form$(9, 0))
  NoticeInfo.MsgLine2 = QPTrim$(Form$(10, 0))
  NoticeInfo.MsgLine3 = QPTrim$(Form$(11, 0))
  NoticeInfo.MsgLine4 = QPTrim$(Form$(12, 0))

  OK2Print = True

CheckReturn:
  
RETURN

NoBlankErr:
  DisplayUBScrn "ERRSCRN1"
  QPrintRC "Invalid, or Blank Field Information!", 10, 23, -1
  QPrintRC "Please Correct and try Again.", 12, 26, -1
  WaitForAction
RETURN

END FUNCTION

FUNCTION MakeDay$ (DayNum)
  SELECT CASE DayNum
  CASE 1
    MakeDay$ = "Sunday"
  CASE 2
    MakeDay$ = "Monday"
  CASE 3
    MakeDay$ = "Tuesday"
  CASE 4
    MakeDay$ = "Wednesday"
  CASE 5
    MakeDay$ = "Thursday"
  CASE 6
    MakeDay$ = "Friday"
  CASE 7
    MakeDay$ = "Saturday"
  END SELECT
  
END FUNCTION

FUNCTION MakeLongDate$ (PDate$)
  DayNum = Num2Day(Date2Num(PDate$))
  DayName$ = MakeDay$(DayNum)
  MonthName$ = MakeMonth$(PDate$)
  MakeLongDate$ = DayName$ + ", " + MonthName$ + " " + MID$(PDate$, 4, 2) + ", " + RIGHT$(PDate$, 4)
END FUNCTION

SUB NoticeMailLabel (NoticeInfo AS NoticeInfoType)

  SHARED Choice$()
  REDIM Choice$(0 TO 3, 0 TO 1)

  REDIM ToPrint(1 TO 5) AS STRING * 132

  FOR Cnt = 1 TO 5
    LSET ToPrint(Cnt) = ""
  NEXT

  DIM POrder$(1 TO 6)
  POrder$(1) = "Customer Name Order"
  POrder$(2) = "Account Number Order"
  POrder$(3) = "Location Number Order"
  POrder$(4) = "ZipCode Number Order"

  Choice$(0, 1) = "1"
  Choice$(1, 1) = "1)  1 x 3" + CHR$(171) + "  1 Wide"
  Choice$(2, 1) = "2)  1 x 3" + CHR$(171) + "  4 Wide"
  
  REDIM OSet(1 TO 4) AS INTEGER

  OSet(1) = 1
  OSet(2) = 37
  OSet(3) = 74
  OSet(4) = 110

  IF FileSize&("UBLNINFO.DAT") = 0 THEN
    GOTO ExitMailLabListing:
  END IF

  FGetAH "UBLNINFO.DAT", NoticeInfo, LEN(NoticeInfo), 1

  AbortFlag = False

  LibName$ = "UBWRKORD"
  ScrnName$ = "UBLNLABL"

  ' Define Fields

  NumFlds = LibNumberOfFields(LibName$, ScrnName$)

  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

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
      GOSUB ShowNoticeInfo
      Action = 1
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F5KEY
      LType = VAL(LEFT$(Form$(1, 0), 1))
      IF LType = 0 THEN
        ErrCode = 2
        GOSUB ShowErrScrn1
      ELSE
        GOSUB DoAlignLabelMask
      END IF
    CASE F10Key
      LType = VAL(LEFT$(Form$(1, 0), 1))
      IF LType = 0 THEN
        ErrCode = 2
        GOSUB ShowErrScrn1
        OKFlag = False
      ELSE
        OKFlag = True
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

  ERASE Frm, Form$, Fld, Choice$
  
  IdxRecLen = 4               'we are using a long integer
  IdxFileSize& = FileSize&("UBLNIDX.DAT")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBLNIDX.DAT", IdxBuff(1), IdxRecLen, IdxNumOfRecs      'load it
  NumOfRecs = IdxNumOfRecs

  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBRpt = FREEFILE
  OPEN "UBLNMAIL.RPT" FOR OUTPUT AS UBRpt

  BlockClear
  ShowProcessingScrn "Mailing Labels"

  FOR Cnt = 1 TO NumOfRecs
    AcctNumber& = IdxBuff(Cnt).RecNum
    GET UBCust, AcctNumber&, UBCustRec(1)
    
    SELECT CASE LType
    CASE 1
      PRINT #UBRpt, "Cust #" + STR$(AcctNumber&)
      PRINT #UBRpt, LEFT$(QPTrim$(UBCustRec(1).CustName), 23)
      PRINT #UBRpt, LEFT$(QPTrim$(UBCustRec(1).Addr1), 23)
      IF LEN(QPTrim$(UBCustRec(1).Addr2)) > 0 THEN
        PRINT #UBRpt, LEFT$(QPTrim$(UBCustRec(1).Addr2), 23)
        PRINT #UBRpt, LEFT$(QPTrim$(UBCustRec(1).City), 13) + ", " + UBCustRec(1).State + " " + LEFT$(UBCustRec(1).ZipCode, 5)
      ELSE
        PRINT #UBRpt, LEFT$(QPTrim$(UBCustRec(1).City), 13) + ", " + UBCustRec(1).State + " " + LEFT$(UBCustRec(1).ZipCode, 5)
        PRINT #UBRpt,
      END IF
      PRINT #UBRpt,
      didcnt = didcnt + 1
    CASE 2
      LabelCnt = LabelCnt + 1
      MID$(ToPrint(1), OSet(LabelCnt)) = "Cust #" + STR$(AcctNumber&)
      MID$(ToPrint(2), OSet(LabelCnt)) = LEFT$(QPTrim$(UBCustRec(1).CustName), 23)
      MID$(ToPrint(3), OSet(LabelCnt)) = LEFT$(QPTrim$(UBCustRec(1).Addr1), 23)
      IF LEN(QPTrim$(UBCustRec(1).Addr2)) > 0 THEN
        MID$(ToPrint(4), OSet(LabelCnt)) = LEFT$(QPTrim$(UBCustRec(1).Addr2), 23)
        MID$(ToPrint(5), OSet(LabelCnt)) = LEFT$(QPTrim$(UBCustRec(1).City), 13) + ", " + UBCustRec(1).State + " " + LEFT$(UBCustRec(1).ZipCode, 5)
      ELSE
        MID$(ToPrint(4), OSet(LabelCnt)) = LEFT$(QPTrim$(UBCustRec(1).City), 13) + ", " + UBCustRec(1).State + " " + LEFT$(UBCustRec(1).ZipCode, 5)
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
  'IF didcnt > 4 THEN EXIT FOR
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
  ERASE Frm, Form$, Fld, OSet

  IF NOT AbortFlag THEN
    PrintRptFile "Mailing Labels", "UBLNMAIL.RPT", 1, RetCode, EntryPoint
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
  EText$ = "  Invalid Label Type!"
  Frm(1).FldNo = 1
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

ShowNoticeInfo:

   QPrintRC "Notice Count:", 8, 23, 3
   QPrintRC "Notice Date:", 10, 24, 3
   QPrintRC "Print Order:", 12, 24, 3
   QPrintRC "Min Balance:", 14, 24, 3
   QPrintRC QPTrim$(STR$(NoticeInfo.PrnCnt)), 8, 37, 14
   QPrintRC Num2Date$(NoticeInfo.NoticeDate), 10, 37, 14
   QPrintRC QPTrim$(POrder$(NoticeInfo.PrnOrder)), 12, 37, 14
   QPrintRC QPTrim$(FUsing$(STR$(NoticeInfo.MinBalance), "$$####.##")), 14, 37, 14

'  PayByDate     AS INTEGER         '2
'  FromDate      AS INTEGER         '3
'  TODate        AS INTEGER         '4
'  MinBalance    AS DOUBLE          '5
'  BalanceType   AS INTEGER         '6
'  PrnOrder      AS INTEGER         '7
'  UseAFlag      AS INTEGER         '8

RETURN
  
END SUB

SUB PrintLateNotices (NoticeInfo AS NoticeInfoType)

  PDate$ = Num2Date(NoticeInfo.PayByDate)
  NDate$ = Num2Date(NoticeInfo.NoticeDate)

  NMonth$ = LEFT$(MakeMonth$(NDate$), 3) + "."

  LongPDate$ = MakeLongDate$(PDate$)
  LongNDate$ = MakeLongDate$(NDate$)
  
  'load setup file
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen

  IF UBSetUpRec(1).BILLCYCL = "Y" THEN
    UseCycle = True
  END IF
  FromBC = NoticeInfo.FromBC
  ThruBC = NoticeInfo.ThruBC

  MinBalance# = NoticeInfo.MinBalance

  PageNo = 0
  
  Choice = 1
  'Section to check for customer modifications
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen      'load setup file
  TownName$ = UBSetUpRec(1).UTILNAME
  
  SELECT CASE NoticeInfo.PrnOrder
  CASE 1
    IndexName$ = NameIndexFile
  CASE 2
    NoIndex = True
  CASE 3
    IndexName$ = BookIndexFile
  CASE 4
    IndexName$ = "UBTEMP.IDX"
    MakeMowZipCodeIndex "ZipCode"
  END SELECT

  OKFlag = True
  
  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  TBooks = 0
  IF NoIndex = False THEN
    NumOfRecs = FileSize(IndexName$) \ 4
    REDIM IndexArray(1 TO NumOfRecs) AS UBCustIndexRecType
    FGetAH IndexName$, IndexArray(1), 4, NumOfRecs
  ELSE
    NumOfRecs = FileSize("UBCUST.DAT") \ UBCustRecLen
  END IF
  
  UBCust = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS UBCust LEN = UBCustRecLen

  UBRpt = FREEFILE
  OPEN "UBLATNOT.RPT" FOR OUTPUT AS UBRpt

'01-07-99 Added record list of late notices printed (for mailing labels)
  KillFile "UBLNIDX.DAT"
  UBLRec = FREEFILE
  OPEN "UBLNIDX.DAT" FOR RANDOM SHARED AS UBLRec LEN = 4
  
  BlockClear
  ShowProcessingScrn "Processing Late Notices"
  
  Next2Print = 1
  
  FOR Cnt = 1 TO NumOfRecs
    IF NoIndex THEN
      AcctNo& = Cnt
    ELSE
      AcctNo& = IndexArray(Cnt).RecNum
    END IF
    GET UBCust, AcctNo&, UBCustRec(1)
    GotWater = False
    IF UBCustRec(1).DelFlag = 0 AND UBCustRec(1).CUTOFFYN = "Y" THEN
      IF UseCycle THEN
        CustBC = UBCustRec(1).BILLCYCL
      ELSE
        CustBC = Value(UBCustRec(1).Book, ECode)
      END IF
      IF CustBC < FromBC OR CustBC > ThruBC THEN
        GOTO SkipEm
      END IF

      IF UBCustRec(1).CurrRevAmts(1) > 0 THEN
        GotWater = True
      END IF
      Location$ = UBCustRec(1).Book + "-" + UBCustRec(1).SeqNumb
      Acct$ = QPTrim$(STR$(AcctNo&))
      Zip$ = QPTrim$(UBCustRec(1).ZipCode)
      ZipLen = LEN(Zip$)
      SELECT CASE ZipLen
      CASE 9, 10
        Zip$ = LEFT$(Zip$, 5) + "-" + RIGHT$(Zip$, 4)
      CASE ELSE
        Zip$ = LEFT$(Zip$, 5)
      END SELECT
      
      TotalBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
      SELECT CASE NoticeInfo.BalanceType
      CASE 1
        CustBal# = Round#(UBCustRec(1).CurrBalance)
        IF CustBal# > MinBalance# THEN
          IF TotalBal# > 0 THEN
            IF NoticeInfo.UseAFlag THEN
              IF UBCustRec(1).Status = "A" THEN
                Print1 = Print1 + 1
                GOSUB PrintThemOne
              END IF
            ELSE
              Print1 = Print1 + 1
              GOSUB PrintThemOne
            END IF
          END IF
        END IF
      CASE 2
        CustBal# = Round#(UBCustRec(1).PrevBalance)
        IF CustBal# > MinBalance# THEN
          IF TotalBal# > 0 THEN
            IF NoticeInfo.UseAFlag THEN
              IF UBCustRec(1).Status = "A" THEN
                Print1 = Print1 + 1
                GOSUB PrintThemOne
              END IF
            ELSE
              Print1 = Print1 + 1
              GOSUB PrintThemOne
            END IF
          END IF
        END IF
      CASE 3
        CustBal# = Round#(UBCustRec(1).CurrBalance + UBCustRec(1).PrevBalance)
        IF CustBal# > MinBalance# THEN
          IF TotalBal# > 0 THEN
            IF NoticeInfo.UseAFlag THEN
              IF UBCustRec(1).Status = "A" THEN
                Print1 = Print1 + 1
                GOSUB PrintThemOne
              END IF
            ELSE
              Print1 = Print1 + 1
              GOSUB PrintThemOne
            END IF
          END IF
        END IF
      END SELECT
    END IF
    
    IF Next2Print = 4 THEN
      Next2Print = 1
      PRINT #UBRpt, CHR$(12)
    END IF
    
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
SkipEm:
    ShowPctComp Cnt, NumOfRecs
    IF INSTR(COMMAND$, "TEST") > 0 THEN
      IF PrnCnt > 0 THEN EXIT FOR
    END IF
  NEXT
  
  CLOSE
  NoticeInfo.PrnCnt = PrnCnt

  FPutAH "UBLNINFO.DAT", NoticeInfo, LEN(NoticeInfo), 1

  IF NOT AbortFlag THEN
    PrintRptFile "Late Notices Printing", "UBLATNOT.RPT", 1, RetCode, 0
  END IF
  
  GOTO ExitLatePrint
  
PrintThemOne:
   PrnCnt = PrnCnt + 1

  '$INCLUDE: 'MCORLNOT.bi'           'mccormick

  PUT UBLRec, , AcctNo&
RETURN

ExitLatePrint:

END SUB

