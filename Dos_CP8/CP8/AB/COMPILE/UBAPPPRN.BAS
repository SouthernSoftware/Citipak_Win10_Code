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
  MChoice$(1) = " Print ALL Applications "
  MChoice$(2) = " Print a Selected Application "
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

    TitleBox 2, Col, MaxLen + 3, "Application Menu ", Cnf
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
      GET #CustFile, LCnt&, UBCustRec(1)
      IF UBCustRec(1).Delflag = True THEN
        GOTO NoAppThisCust
      END IF
      RecNo& = LCnt&
      GOSUB PrintApp
NoAppThisCust:
      ShowPctComp LCnt&, NumCustRec&
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

  CityStaZip$ = QPTrim$(UBCustRec(1).City) + " " + QPTrim$(UBCustRec(1).State) + " " + QPTrim$(UBCustRec(1).ZipCode)
  PRINT #UBRpt, "~"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(23); "    USER'S ALARM PERMIT APPLICATION"
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(4); CName$; TAB(52); "Account Number: "; RecNo&''' UBCustRec(1).Book; "-"; UBCustRec(1).Seqnumb
  PRINT #UBRpt, TAB(4); QPTrim$(UBCustRec(1).Addr1); TAB(50); "Application Date: "; AppDate$
  PRINT #UBRpt, TAB(4); CityStaZip$
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(20); "Application Information Required For an Alarm:"
  PRINT #UBRpt, TAB(19); "PLEASE VERIFY/CORRECT ANY PRE-PRINTED INFORMATION"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(4); "1. Applicant Name:_________________________"; TAB(50); "Phone No. "; QPTrim$(UBCustRec(1).HPhone)
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(4); "2. Business/Residents Name:    "; QPTrim$(UBCustRec(1).CustName)
  PRINT #UBRpt, TAB(4); "                               컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴"
  PRINT #UBRpt, TAB(4); "3. Mailing Address:            "; QPTrim$(UBCustRec(1).Addr1)
  PRINT #UBRpt, TAB(4); "                               컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴"
  PRINT #UBRpt, TAB(4); "                               "; CityStaZip$
  PRINT #UBRpt, TAB(4); "                               컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴"
  PRINT #UBRpt, TAB(4); "4. Address of Alarm:           "; QPTrim$(UBCustRec(1).ServAddr)
  PRINT #UBRpt, TAB(4); "                               컴컴컴컴컴컴컴컴컴컴컴컴컴컴컴"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(4); "List three (3) individuals who may be contacted in the event of an alarm"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(4); "(1)_______________________________________________________________________"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(4); "(2)_______________________________________________________________________"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(4); "(3)_______________________________________________________________________"
  PRINT #UBRpt, "                      (Name)                (Address)            (Phone No.)"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(4); "Alarm Company:__________________________  Type of Alarm:__________________"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(4); "Representative:______________________________  Phone No:__________________"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(4); "I agree to accept responsibility for any malfunctions of my alarm system"
  PRINT #UBRpt, TAB(4); "and accept responsibility for maintenance/service of my alarm equipment."
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(8); "Applicant Signature:___________________________   Date:_______________"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(8); "    Print Your Name:___________________________   Date:_______________"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(4); "Please return completed application form to:"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(4); "Darlington County Communications Center"
  PRINT #UBRpt, TAB(4); "Attn: Libby West"
  PRINT #UBRpt, TAB(4); "1625 Harry Byrd Highway"
  PRINT #UBRpt, TAB(4); "Darlington, SC 29532"
  PRINT #UBRpt,
  PRINT #UBRpt, TAB(4); "Upon receipt of application you will be sent an Alarm Permit Number."
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

  '--Set screen number to one and display screen

  BlockClear

  DisplayABScrn ScrnName$

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

