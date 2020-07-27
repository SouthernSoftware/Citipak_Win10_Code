DEFINT A-Z
DECLARE FUNCTION MakeMonth$ (TDate$)
DECLARE FUNCTION MakeDay$ (DayNum%)
DECLARE FUNCTION MakeLongDate$ (PDate$)
DECLARE SUB PrintLateNotices (NoticeInfo AS ANY)
DECLARE FUNCTION GetNoticeDates% (NoticeInfo AS ANY)
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

TYPE NoticeInfoType
  NoticeDate    AS INTEGER         '1
  paybydate     AS INTEGER         '2
  FromDate      AS INTEGER         '3
  TODate        AS INTEGER         '4
  MinBalance    AS DOUBLE          '5
  BalanceType   AS INTEGER         '6
  PrnOrder      AS INTEGER         '7
  UseAFlag      AS INTEGER         '8
  MsgLine1      AS STRING * 25
  MsgLine2      AS STRING * 25
  MsgLine3      AS STRING * 25
  MsgLine4      AS STRING * 25
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
  
  DO
    IF NOT GetNoticeDates(NoticeInfo) THEN GOTO ExitPrint
    PrintLateNotices NoticeInfo
  LOOP
  
ExitPrint:
  
  IF INSTR(COMMAND$, "TEST") THEN
    'ClearScrn
    HideCursor
    END
  ELSE
    RUN "UBBILLIN"
  END IF

FUNCTION GetNoticeDates% (NoticeInfo AS NoticeInfoType)

  Fill$ = SPACE$(25)

  REDIM ScrnArray(0)

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen  'load setup file
  TownName$ = UBSetUpRec(1).UTILNAME

  IF INSTR(TownName$, "APPALACH") THEN
    AppFlag = True
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

  IF NOT AppFlag THEN
    Fld(3).Protected = True
    Fld(4).Protected = True
    QPrintRC Fill$, 8, 26, 0
    QPrintRC Fill$, 9, 26, 0
    Frm(1).StayOnField = False
  END IF

  FirstTime = True
  ShowCursor
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      IF NOT AppFlag THEN
        LSET Form$(1, 0) = DATE$
        LSET Form$(2, 0) = Form$(1, 0)
        LSET Form$(5, 0) = "0"
        LSET Form$(8, 0) = "Y"
        IF WEnvTest THEN
          LSET Form$(6, 0) = Choice$(3, 0)
          LSET Form$(7, 0) = Choice$(2, 1)
        END IF
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
        GetNoticeDates = True
        ExitFlag = True
      END IF
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 15   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 21 TO 33           '--cancel button
          PressButton 27, 15, 21, 33
        CASE 34 TO 48           '--Alignment
          PressButton F5KEY, 15, 34, 48
        CASE 49 TO 60           '--Save Button
          PressButton -68, 15, 49, 60
        END SELECT
      END SELECT                'row
    END IF
    
    '--Check screen page
  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag
  
  IF Frm(1).KeyCode = 27 THEN
    GetNoticeDates% = False
    GOTO ExitEdit
  END IF
  
ExitEdit:
  HideCursor
EXIT FUNCTION
  
CheckDates:
  OK2Print = False
  BadDate = False

  NDate = Date2Num(Form$(1, 0))
  PDate = Date2Num(Form$(2, 0))
  FDate = Date2Num(Form$(3, 0))
  TDate = Date2Num(Form$(4, 0))

  IF NDate < 0 THEN
    BadDate = True
    Frm(1).FldNo = 1
  ELSEIF PDate < 0 THEN
    BadDate = True
    Frm(1).FldNo = 2
  END IF

  IF BadDate THEN
    OK = MsgBox%("UB", "UBPBDATE")
    RestScrn ScrnArray()
    GOTO CheckReturn
  END IF
  
  IF AppFlag THEN
    IF FDate < 0 THEN
      BadDate = True
      Frm(1).FldNo = 3
    ELSEIF TDate < 0 THEN
      BadDate = True
      Frm(1).FldNo = 4
    END IF
  END IF

  IF BadDate THEN
    OK = MsgBox%("UB", "UBPBDATE")
    RestScrn ScrnArray()
    GOTO CheckReturn
  ELSE
    NoticeInfo.NoticeDate = NDate
    NoticeInfo.paybydate = PDate
    NoticeInfo.FromDate = FDate
    NoticeInfo.TODate = TDate
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

FUNCTION MakeMonth$ (TDate$)
  Month = VAL(LEFT$(TDate$, 2))
  SELECT CASE Month
  CASE 1
    MakeMonth$ = "January"
  CASE 2
    MakeMonth$ = "February"
  CASE 3
    MakeMonth$ = "March"
  CASE 4
    MakeMonth$ = "April"
  CASE 5
    MakeMonth$ = "May"
  CASE 6
    MakeMonth$ = "June"
  CASE 7
    MakeMonth$ = "July"
  CASE 8
    MakeMonth$ = "August"
  CASE 9
    MakeMonth$ = "September"
  CASE 10
    MakeMonth$ = "October"
  CASE 11
    MakeMonth$ = "November"
  CASE 12
    MakeMonth$ = "December"
  END SELECT
END FUNCTION

SUB PrintLateNotices (NoticeInfo AS NoticeInfoType)

  PDate$ = Num2Date(NoticeInfo.paybydate)
  NDate$ = Num2Date(NoticeInfo.NoticeDate)
  
  LongPDate$ = MakeLongDate$(PDate$)
  LongNDate$ = MakeLongDate$(NDate$)
  
  'load setup file
  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
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
    IF NOT UBCustRec(1).DelFlag AND UBCustRec(1).CUTOFFYN = "Y" THEN
      IF UBCustRec(1).CurrRevAmts(1) > 0 THEN
        GotWater = True
      END IF
      Location$ = UBCustRec(1).BOOK + "-" + UBCustRec(1).SEQNUMB
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
    IF PCnt2 > 1 THEN EXIT FOR
  NEXT
  
  'PRINT #UBRpt, CHR$(12)
  
  CLOSE
  
  PrintRptFile "Late Notices Printing", "UBLATNOT.RPT", 1, RetCode, 0
  
  GOTO ExitLatePrint
  
PrintThemOne:
  PCnt2 = PCnt2 + 1
  ''$INCLUDE: 'HEMILNOT.bi'            'Hemingway
  ''$INCLUDE: 'ClevLNOT.bi'            'cleveland
  ''$INCLUDE: 'SaluLNOT.bi'            'Saluda
  ''$INCLUDE: 'LONGLNOT.bi'             'Longview
  ''$INCLUDE: 'TROULNOT.bi'            'troutman
  ''$INCLUDE: 'INDIANTR.bi'            'Indian Trail Late Notice
  ''$INCLUDE: 'LANDLNOT.bi'           'landis
  ''$INCLUDE: 'BRNCHNOT.bi'           'Branchville
  ''$INCLUDE: 'PSA1LNOT.bi'           'psa giles
  ''$INCLUDE: 'GORDLNOT.bi'           'gordonsville
  ''$INCLUDE: 'MCORLNOT.bi'           'mccormick
  ''$INCLUDE: 'LAWRLNOT.bi'           'lawrenceville
  ''$INCLUDE: 'GILBLNOT.bi'           'gilbert
  ''$INCLUDE: 'MONTLNOT.bi'           'montross
  '$INCLUDE: 'wadeLNOT.bi'           'wadesboro
  ''$INCLUDE: 'TennLNOT.bi'           'Tenn Ridge
  ''$INCLUDE: 'jvilLNOT.bi'          'Johnsonville
  ''$INCLUDE: 'lanelnot.bi'          'lane
  ''$INCLUDE: 'POLKLNOT.bi'          'Polkton NC
  ''$INCLUDE: 'HILLLNOT.BI'          'hillsville
  ''$INCLUDE: 'RUSLLNOT.bi'          'russel county
  ''$INCLUDE: 'MIDTLNOT.bi'          'Middletown VA
  ''$INCLUDE: 'WHITLNOT.bi'          'Whitakers
  ''$INCLUDE: 'BROKLNOT.bi'          'Brookneal
  ''$INCLUDE: 'JNCOLNOT.bi'          'Johnston CO
  ''$INCLUDE: 'ALBRLNOT.bi'          'Alberta
  ''$INCLUDE: 'APPLLNOT.bi'          'Appalach
  ''$INCLUDE: 'GATELNOT.bi'          'Gate City
  ''$INCLUDE: 'CHILHOWE.BI'          'chillhowe
  ''$INCLUDE: 'BLADLNOT.BI'          '
  ''$INCLUDE: 'TROYLNOT.BI'          '
  ''$INCLUDE: 'CEDRLNOT.BI'          'Cedar Bluff
  ''$INCLUDE: 'DSONLNOT.BI'          'Dobson
  ''$INCLUDE: 'RYANLNOT.BI'          'Ryan OK
  ''$INCLUDE: 'crewLNOT.BI'          'Crewe VA
  ''$INCLUDE: 'BURKLNOT.BI'          'Burkeville
  ''$INCLUDE: 'SPENLNOT.BI'           'Spencer TN
  ''$INCLUDE: 'OSTDLNOT.BI'           'Brunswick
  ''$INCLUDE: 'MCADLNOT.BI'           'Mcadenville
RETURN

ExitLatePrint:

END SUB

