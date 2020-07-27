DEFINT A-Z
DECLARE SUB DisplayABScrn (ScrnName$)
DECLARE FUNCTION GetAppDate$ (NOGOFlag%)
DECLARE SUB AppPrint (RecNo&, AppDate$)
DECLARE FUNCTION PreBillYouSure% ()
DECLARE FUNCTION Date2Num (TheDate$)
DECLARE FUNCTION AskAbandonPrint% ()
DECLARE FUNCTION ErrorScrn (WhatError%, Acct&)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION GetAdjFactor# ()
DECLARE FUNCTION GetBillBook% ()
DECLARE FUNCTION GetBillCycle% ()
DECLARE FUNCTION GetCustMeterType% (UBCustRec() AS ANY, ThisMeter%)
DECLARE FUNCTION GetNumOfRevs% ()
DECLARE FUNCTION GetNumRateRecs% ()
DECLARE FUNCTION GetRevCharge# (RateTbl AS ANY, TMeterConsp&, MeterMulti&)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (LongNum$)
DECLARE FUNCTION Round# (N#)
DECLARE SUB BlockClear ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB Get.Moose.OR.Key (Ky$, MooseButton%, MRow%, MCol%)
DECLARE SUB GetPreBillOrder (Choice%, ExitFlag%, SeqFlag$)
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetUpLen%)
DECLARE SUB MakeBillFile (AbortFlag%, FuelAdjAmt#, ThisCycle%, ThisBook%)
DECLARE SUB MakeMowZipCodeIndex (IndexText$)
DECLARE SUB MakePostalIndex (IndexText$)
DECLARE SUB MakeSequenceIndex (IndexText$)
DECLARE SUB MakeZipCodeIndex (IndexText$)
DECLARE SUB PostBillTrans ()
DECLARE SUB PreBillReport ()
DECLARE SUB PrintHighLowReport ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RateCodeErrScrn (RATECODE$)
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB ShowCursor ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB UBLog (Text$)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB QPSound (Freq, Dur)
DECLARE SUB LookUp (RecNo&, Text$, DefaultLook%, CLSFlag%, ActiveOnly%)

  '$INCLUDE: 'ubsetup.bi'
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  '$INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.BI'

  STACK 8000

  CONST False = 0, True = NOT False

  CrLf$ = CHR$(13) + CHR$(10)
  FF$ = CHR$(12)

  REDIM UBSetUpRec(1) AS UBSetupRecType
  LoadUBSetUpFile UBSetUpRec(), UBSetUpLen
  
  '033198 Added Conway special handling

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 3)
  MChoice$(1) = " Print ALL Permits "
  MChoice$(2) = " Print a Selected Permits "
  MChoice$(3) = " Exit to OS"

  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line
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

    TitleBox 2, Col, MaxLen + 3, "Permits Menu ", Cnf
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
      AppDate$ = GetAppDate$(NOGOFlag)
      IF NOGOFlag = 0 THEN
        AppPrint 0&, AppDate$
      END IF
    CASE 2

      DO
        BlockClear
        LookUp RecNo&, "Customer", 2, True, False
        IF RecNo& > 0 THEN
          AppDate$ = GetAppDate$(NOGOFlag)
          IF NOGOFlag = 0 THEN
            AppPrint RecNo&, AppDate$
          END IF
        END IF
      LOOP WHILE RecNo& > 0
    CASE 3
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP UNTIL ExitFlag

  RUN "UBBillin"

SUB AppPrint (RecNo&, AppDate$)

  Year = VAL(RIGHT$(AppDate$, 4))
  RYear$ = QPTrim$(STR$(Year + 1))

  REDIM UBCustRec(1) AS NewUBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))

  NumCustRec& = FileSize&("UBCUST.DAT") \ UBCustRecLen

  CustFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM SHARED AS CustFile LEN = UBCustRecLen

  UBRpt = FREEFILE
  OPEN "UBAPPPRN.RPT" FOR OUTPUT AS UBRpt

  IF RecNo& > 0 THEN
    GET #CustFile, RecNo&, UBCustRec(1)
    GOSUB PrintApp
  ELSE
    ShowProcessingScrn "Creating Applications"
    FOR LCnt& = 1 TO NumCustRec&
      RecNo& = LCnt&
      GET #CustFile, LCnt&, UBCustRec(1)
      IF UBCustRec(1).DelFlag <> 0 THEN
        GOTO NotThisOneMan
      END IF
      GOSUB PrintApp
      ShowPctComp LCnt&, NumCustRec&
NotThisOneMan:
    NEXT
  END IF
  CLOSE
  IF NOT AbortFlag THEN
    PrintRptFile "Print Application", "UBAPPPrn.RPT", LPTPort, RetCode, EntryPoint
  END IF
  
  EXIT SUB

PrintApp:

  TName$ = QPTrim$(UBCustRec(1).CustName)
  CNLen = LEN(TName$)
  SPos = INSTR(TName$, ",")
  'STOP
  IF SPos > 0 THEN 'AND SPos < CNLen THEN
    MID$(TName$, SPos, 1) = " "
    'CName$ = LEFT$(TName$, SPos - 1) + "," + MID$(TName$, SPos)
  'ELSE
  '  CName$ = TName$
  END IF
  CName$ = TName$

'  TName$ = QPTrim$(UBCustRec(1).CustName)
'  CNLen = LEN(TName$)
'
'  SPos = INSTR(TName$, " ")
'  IF SPos > 0 AND SPos < CNLen THEN
'    CName$ = LEFT$(TName$, SPos - 1) + "," + MID$(TName$, SPos)
'  ELSE
'    CName$ = TName$
'  END IF

  CityStaZip$ = QPTrim$(UBCustRec(1).City) + " " + QPTrim$(UBCustRec(1).State) + " " + QPTrim$(UBCustRec(1).ZipCode)
  PRINT #UBRpt, "~"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(50); AppDate$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(8); CName$
  PRINT #UBRpt, TAB(8); QPTrim$(UBCustRec(1).Addr1)
  PRINT #UBRpt, TAB(8); CityStaZip$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(8); "Dear Sir/Madam:"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(8); "THANK YOU FOR COMPLETING AND RETURNING THE ALARM PERMIT APPLICATION"
  PRINT #UBRpt, TAB(8); "FORM FOR PROPERTY LOCATED AT: "; QPTrim$(UBCustRec(1).ServAddr)
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(8); " YOUR ALARM PERMIT NUMBER IS: "; UBCustRec(1).Book; "-"; UBCustRec(1).Seqnumb
  PRINT #UBRpt, TAB(8); "YOUR ALARM ACCOUNT NUMBER IS:"; RecNo&
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(8); "ALARM PERMITS MUST BE RENEWED ANNUALLY.  YOUR PERMIT WILL NEED TO"
  PRINT #UBRpt, TAB(8); "BE RENEWED IN JANUARY "; RYear$
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(8); "SINCERELY,"

  PRINT #UBRpt, TAB(8); "LIBBY WEST"
  PRINT #UBRpt, TAB(8); "COMMUNICATIONS DIRECTOR"
  PRINT #UBRpt, TAB(8); "(843) 398-4920"
  PRINT #UBRpt, CHR$(12)


RETURN

END SUB

SUB DisplayABScrn (ScrnName$)
  LibFile2Scrn "UBALARM", ScrnName$, MonoCode, Attribute%, ErrCode
END SUB

FUNCTION GetAppDate$ (NOGOFlag)

  NOGOFlag = False

  LibName$ = "UBALARM"
  ScrnName$ = "APPDATE"
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

  'FirstTime = True
  '--Set screen number to one and display screen

  BlockClear

  DisplayABScrn ScrnName$
  QPrintRC "     Permit", 10, 25, -1

  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      OKFlag = True

    CASE EscKey
      ExitFlag = True
      NOGOFlag = True
    END SELECT
  LOOP UNTIL OKFlag OR ExitFlag

  IF NOGOFlag THEN
    GetAppDate$ = ""
  ELSE
    ThisDate$ = QPTrim$(Form$(1, 0))
    IF LEN(ThisDate$) < 10 THEN
      GetAppDate$ = ""
      NOGOFlag = True
    ELSE
      GetAppDate$ = ThisDate$
    END IF
  END IF

END FUNCTION

