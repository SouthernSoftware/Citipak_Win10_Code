DEFINT A-Z
DECLARE SUB SortStr2 (BYVAL Address%, Size%, Direction%)
DECLARE SUB SelectHHCode (RecNo%)
DECLARE SUB AddEditHHCode (EDFlag)
DECLARE SUB PrintHHCode ()
DECLARE SUB SendHHCode2Husky ()
DECLARE SUB HHCodeMaint ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION nGetPassword$ (Row%, Col%, Max%, Abort%)
DECLARE SUB UBLog (Text$)
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB WaitForAction ()
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB KillFile (File2Kill$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB SaveScrn (ScrnArray())
DECLARE SUB RestScrn (ScrnArray())
DECLARE SUB QPrintRC (Text$, Row%, Col%, Kolor)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION Date2Num% (DateString$)
DECLARE FUNCTION Num2Date$ (DateNum)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION PromptSaveData% ()
DECLARE FUNCTION ASCII (Character$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValI% (Number$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION QPStrL$ (Number&)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB HideCursor ()
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB ClearScrn ()
DECLARE SUB LoadUBSetUpFile (UBSetUpRec() AS ANY, UBSetupLen%)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'Qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  'INCLUDE: 'ubdraft.BI'
  '$INCLUDE: 'ubsetup.bi'
  'INCLUDE: 'ubtrans.bi'
  '$INCLUDE: 'newcust.bi'
  '$INCLUDE: 'ubrate.BI'
  '$INCLUDE: 'ubhusky.BI'

  CONST False = 0, True = NOT False

SUB AddEditHHCode (EDFlag)

  REDIM CodeRec(1) AS MtrNotesType
  CodeLen = LEN(CodeRec(1))

  LibName$ = "UBHH"
  ScrnName$ = "HHCODEED"
  NumScrns = 1

  '--define the multi-choice fields
  NumFlds = -1
  NumFlds = LibNumberOfFields(LibName$, ScrnName$) + 1

  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo

  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode

AddReEntry:
  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT

  IF EDFlag THEN
    SelectHHCode RecNo
    IF RecNo > 0 THEN
      GOSUB LoadCodeRec
    ELSE
      GOTO ExitAddEdHHCode
    END IF
  END IF
  
  BlockClear
  LibFile2Scrn LibName$, ScrnName$, MonoCode, Attribute%, ErrCode

  ShowCursor

  Action = 1
  FirstTime = True

  DO

    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      QPrintRC "[                       ]", 7, 17, 10
      IF EDFlag THEN
        LSET Form$(1, 0) = CodeRec(1).Note
        Action = 1
        QPrintRC "Edit an Existing Code", 7, 19, 14
      ELSE
        QPrintRC "Add New HandHeld Code", 7, 19, 14
      END IF
      FirstTime = False
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      ThisCode$ = QPTrim$(Form$(1, 0))
      IF LEN(ThisCode$) > 0 THEN
        CodeRec(1).Note = ThisCode$
        GOTO SaveCode
      END IF
    CASE EscKey
      ExitFlag = True
    END SELECT

  LOOP UNTIL ExitFlag

ExitAddEdHHCode:
EXIT SUB

SaveCode:
  UBFile = FREEFILE
  OPEN "UBHHCODE.DAT" FOR RANDOM SHARED AS UBFile LEN = CodeLen
  IF EDFlag THEN
    ThisRec = RecNo
  ELSE
    ThisRec = LOF(UBFile) \ CodeLen + 1
  END IF
  PUT UBFile, ThisRec, CodeRec(1)
  CLOSE
  DisplayUBScrn "UPDATEOK"
  WaitForAction
GOTO AddReEntry

LoadCodeRec:
  UBFile = FREEFILE
  OPEN "UBHHCODE.DAT" FOR RANDOM SHARED AS UBFile LEN = CodeLen
  GET UBFile, RecNo, CodeRec(1)
  CLOSE

RETURN

END SUB

SUB HHCodeMaint

'  UBLog " IN: Edit HH Codes."

  REDIM HHCodeRec(1) AS HHCodeRecType
  HHCodeLen = LEN(HHCodeRec(1))
  IF Exist("UBHHCODE.DAT") THEN

  END IF

  REDIM MChoice$(1 TO 5)

  MChoice$(1) = "Add a New HandHeld Code"
  MChoice$(2) = "Edit an Existing HandHeld Code"
  MChoice$(3) = "Print HandHeld Codes Listing"
  MChoice$(4) = "Send HandHeld Codes to Husky"
  MChoice$(5) = "Exit to OS"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2)

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    TitleBox 2, Col, MaxLen + 3, "HandHeld Code Maintenance", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN
      Choice = 0
      ExitFlag = True
    END IF

    SELECT CASE Choice
    CASE 1
      AddEditHHCode False
    CASE 2
      AddEditHHCode True
    CASE 3
      PrintHHCode
    CASE 4
      SendHHCode2Husky
    CASE 5
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP UNTIL ExitFlag

END SUB

SUB PrintHHCode

  FF$ = CHR$(12)
  MaxLines = 55

  REDIM CodeRec(1) AS MtrNotesType
  CodeLen = LEN(CodeRec(1))
  NumCodeRecs = FileSize("UBHHCODE.DAT") \ CodeLen
  
  IF NumCodeRecs = 0 THEN
    GOTO ExitCodeListing
  END IF

  REDIM UBSetUp(1) AS UBSetupRecType
  UBSetupLen = LEN(UBSetUp(1))
  UBFile = FREEFILE
  OPEN "UBSETUP.DAT" FOR RANDOM SHARED AS UBFile LEN = UBSetupLen
  IF LOF(UBFile) = 0 THEN
    OkFlag = False
  ELSE
    OkFlag = True
    GET UBFile, 1, UBSetUp(1)
  END IF
  CLOSE

  TownName$ = UBSetUp(1).UTILNAME
  Dash80$ = STRING$(78, "-")

  BlockClear
  ShowProcessingScrn "HandHeld Code Listing."

  UBFile = FREEFILE
  OPEN "UBHHCODE.DAT" FOR RANDOM SHARED AS UBFile LEN = CodeLen
  RPTFile = FREEFILE
  OPEN "CODELIST.RPT" FOR OUTPUT AS RPTFile
  GOSUB PrintCodeHeader
  FOR Cnt = 1 TO NumCodeRecs
    GET UBFile, Cnt, CodeRec(1)
    PRINT #RPTFile, CodeRec(1).Note
    LineCnt = LineCnt + 1
    ShowPctComp Cnt, NumCodeRecs
    IF AskAbandonPrint% THEN
      AbortFlag = True
      EXIT FOR
    END IF
    IF LineCnt > MaxLines THEN
      PRINT #RPTFile, FF$
      GOSUB PrintCodeHeader
    END IF
  NEXT
  PRINT #RPTFile, Dash80$
  PRINT #RPTFile, FF$
  CLOSE

  PrintRptFile "HandHeld Code Listing.", "CODELIST.RPT", 1, RetCode%, 1


ExitCodeListing:
EXIT SUB


PrintCodeHeader:
  PageNo = PageNo + 1
  PRINT #RPTFile, "Utility Billing HandHeld Code Listing."
  PRINT #RPTFile, TownName$; TAB(70); "Page:"; PageNo
  PRINT #RPTFile, "Report Date: "; DATE$
  PRINT #RPTFile, Dash80$
  LineCnt = 4
RETURN


END SUB

SUB SelectHHCode (RecNo)

  STATIC Choice
  IF Choice = 0 THEN Choice = 1

  REDIM CodeRec(1) AS MtrNotesType
  CodeLen = LEN(CodeRec(1))
  NumCodeRecs = FileSize("UBHHCODE.DAT") \ CodeLen

  IF NumCodeRecs = 0 THEN
    GOTO ExitCodeSelect
  END IF
  
  REDIM MChoice$(1 TO NumCodeRecs)  'setup array to number of rates
  UBFile = FREEFILE                   'open the rate code file
  OPEN "UBHHCODE.DAT" FOR RANDOM SHARED AS UBFile LEN = CodeLen
  FOR Cnt = 1 TO NumCodeRecs        'step thru all rate recs
    GET UBFile, Cnt, CodeRec(1)
    MChoice$(Cnt) = SPACE$(23)        'setup space in the array element
    LSET MChoice$(Cnt) = CodeRec(1).Note      'load rate info
    MID$(MChoice$(Cnt), 22) = MKI$(Cnt)                'set recno in array
  NEXT                                        'go till all processed
  CLOSE                                       'close rate file

  SortStr2 VARPTR(MChoice$(1)), Cnt - 1, 0    'sort the rate codes

  MaxLen = 20   'Set menu width to zero
  BoxBot = 16   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something

  TText$ = SPACE$(MaxLen + 4)
  LSET TText$ = "     HandHeld Codes"

  '--Center Menu within Screen

  Row = 6                                   'top title box

  Col = 29

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    ShowCursor
    TitleBox 20, Col, MaxLen + 3, " Use " + CHR$(24) + "-" + CHR$(25) + " to select ", Cnf
    QPrintRC TText$, Row - 1, Col, 112
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      RecNo = 0
      ExitFlag = True
    ELSE
      RecNo = CVI(MID$(MChoice$(Choice), 22, 2))
      ExitFlag = True
    END IF

  LOOP UNTIL ExitFlag

ExitCodeSelect:

END SUB

SUB SendHHCode2Husky

  IF Exist("UBHUSKY.CFG") THEN
    CFGFile = FREEFILE
    OPEN "UBHUSKY.CFG" FOR INPUT AS #CFGFile
    LINE INPUT #CFGFile, Port$
    CLOSE CFGFile
  ELSE
    CLS
    PRINT "CAN NOT FIND THE HUSKY CONFIGURATION FILE:  UBHUSKY.CFG"
    PRINT
    PRINT "Please call Southern Software for support. . ."
    PRINT
    PRINT "Press any key to EXIT.";
    WaitForAction
    GOTO ExitHHSend
  END IF

  IF NOT Exist("UBHHCODE.DAT") THEN
    GOTO ExitHHSend
  END IF

  Ok = MsgBox%("UBSETUP", "UBHUSK1C")
  CursorOff
  LibFile2Scrn "UBHH", "UBHUSK3C", MonoCode, Attribute%, ErrCode
  IF Ok = 1 THEN
    QPrintRC "Transfering file to the HUSKY.", 10, 25, -1
    LOCATE 14, 40
    IF WEnvTest = 0 THEN
      SHELL "cmdhcom " + Port$ + " /tx=UBHHCODE.DAT /noscreen /abort"
    END IF
    QPrintRC "    File Transfer Complete.   ", 10, 25, 15
  ELSE
    QPrintRC "    File Transfer Aborted.    ", 10, 25, -1
  END IF
  QPrintRC "  Press any key to continue.  ", 14, 25, -1
  CALL WaitForAction

ExitHHSend:

END SUB

