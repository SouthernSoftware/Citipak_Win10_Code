DEFINT A-Z
DECLARE SUB AdCostLabels ()
DECLARE FUNCTION Date2Num% (TheDate$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FLof& (Handle%)
DECLARE FUNCTION FUsing$ (Number$, Image$)
DECLARE FUNCTION FileSize& (FileName$)
DECLARE FUNCTION GetCustBalance# (RecNo&)
DECLARE FUNCTION GetTaxCustCnt& ()
DECLARE FUNCTION IsCustDeleted% (AcctNum&)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Num2Date$ (DateNumber%)
DECLARE FUNCTION ParseBillNum$ (Text$)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION QPValL& (Number$)
DECLARE FUNCTION Round# (DblNum#)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB MakeAdCostTrans ()
DECLARE SUB AdvLookUp (RecNo&, Text$, CLSFlag%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB ClearBack ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayTaxScrn (ScrnName$)
DECLARE SUB EditTrans ()
DECLARE SUB FClose (Handle%)
DECLARE SUB FGetA (Handle%, SEG Element AS ANY, NumBytes AS ANY)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FOpenS (FileName$, Handle)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FPutRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB GetPersRecList (PersRecs() AS LONG, CustRec&)
DECLARE SUB GetPropRecList (PropRecs() AS LONG, CustRec&)
DECLARE SUB HideCursor ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB Labels ()
DECLARE SUB MPaintBox (UlRow%, UlCol%, LRRow%, LRCol%, Colr%)
DECLARE SUB MakeCustIndex (IdxType)
DECLARE SUB PostTransactions ()
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PrintEditReport ()
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB RestScrn (Array%())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB Search4AdvAcct (AcctNum&, RecNo&)
DECLARE SUB Search4AdvCust (Search$, RecNo&, CLSFlag%)
DECLARE SUB ShowCursor ()
DECLARE SUB ShowCustHistory (CustRec&)
DECLARE SUB ShowPctComp (BYVAL RecNo%, BYVAL NumOfRecs%)
DECLARE SUB ShowPctCompL (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB ShowProcessingScrn (RptTitle$)
DECLARE SUB ShowSearchPCT (BYVAL RecNo&, BYVAL NumOfRecs&)
DECLARE SUB SmallPause ()
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB TransactionJournal ()
DECLARE SUB VertMenuT2 (Items() AS ANY, Choice, MaxLen%, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB WaitForAction ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
'DECLARE FUNCTION ParseBillNum$ (Text$)


  TYPE FLen2
    V AS STRING * 64
  END TYPE

  CONST False = 0, True = NOT False

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'
  '$INCLUDE: 'TAXCONST.BI'
  '$INCLUDE: 'PROPAbst.BI'
  '$INCLUDE: 'TAXRPTTY.BI'

  TYPE Struct
    who AS STRING * 14
    RecNum AS INTEGER
  END TYPE

  STACK 5000

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 6)

  MChoice$(1) = "Calculate Advertising Charges"
  MChoice$(2) = "Edit Advertising Transactions"
  MChoice$(3) = "Print Advertising Edit Report"
  MChoice$(4) = "Advertising Mailing Labels"
  MChoice$(5) = "Post Advertising Transactions"
  MChoice$(6) = "Exit to OS"

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
  Row = ((24 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 1

  IF Exist("TAXFROMM.DAT") THEN
    'KillFile "TAXFROMM.DAT"
    MainFlag = True
  END IF

  DO

    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0

    ClearBack

    TitleBox 2, Col, MaxLen + 3, "Tax Advertising Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor

    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    CASE 1   'Calc
      MakeAdCostTrans
    CASE 2   'Edit
      EditTrans
    CASE 3   'Print Report
      PrintEditReport
    CASE 4
      AdCostLabels
    CASE 5   'Post
      PostTransactions
    CASE 6
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP

  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF

  IF INSTR(COMMAND$, "TEST") THEN
    HideCursor
    ClearScrn
    END
  END IF

  RUN "TaxMenu"

  END

SUB AdCostLabels
  
  SHARED Choice$()
  
  REDIM Array(1 TO 1) AS Struct 'Template for the sort array
  ReportFile$ = "TAXLABEL.PRN"  'Report File Name
  Dash80$ = STRING$(80, "=")
  FF$ = CHR$(12)

  REDIM TaxCustRec(1)  AS TaxCustType
  TaxCustLen = LEN(TaxCustRec(1))

  REDIM TaxAdvRec(1 TO 1) AS InterestRecType
  TaxAdvRecLen = LEN(TaxAdvRec(1))

  TaxFile = FREEFILE
  OPEN TaxAdvFile FOR RANDOM AS TaxFile LEN = TaxAdvRecLen      'open data file
  NumAdvRec& = LOF(TaxFile) / TaxAdvRecLen
  CLOSE TaxFile

  MaxLines = 56
  LineCnt = 0
  CustCnt = 0
  Fill$ = SPACE$(25)
  LibName$ = "TAX"
  ScrnName$ = "VCUSTLAB"
  
  ' Define Fields
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  ' Define Quick Screen Form Editing Arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  ' Get 1st & Last Fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  ' Clear Fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT F
  
  REDIM Choice$(0 TO 2, 0 TO 2)
  
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "Name Order"
  Choice$(2, 0) = "Account Number"
  Choice$(0, 2) = "4"
  Choice$(1, 2) = "Screen"
  Choice$(2, 2) = "Printer"
  'Form$(2, 0) = "R"             'Default to Real
  'Form$(3, 0) = "N"             'Default to No
  Action = 1
  ClearBack
  
  FirstTime = True
  ShowCursor
  
  DisplayTaxScrn ScrnName$
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action

    IF FirstTime THEN
      FirstTime = False
      Fld(1).Protected = True
      Fld(2).Protected = True
      Fld(3).Protected = True
      QPrintRC Fill$, 10, 22, -1
      QPrintRC Fill$, 12, 22, -1
      QPrintRC "This Procedure Prints The Ad-Cost", 11, 23, 14
      QPrintRC "Mailing Labels Only!!!", 13, 23, 14
      QPrintRC Fill$, 14, 21, -1

      Action = 1
    END IF

    SELECT CASE Frm(1).KeyCode
    CASE F10Key
      DevSpec$ = LEFT$(Form$(4, 0), 1)
      ExitFlag = True
    CASE F5KEY
      GOSUB PrintAlign
    CASE EscKey
      AbortFlag = True
      ExitFlag = True           'EXIT DO
    END SELECT
  LOOP UNTIL ExitFlag
  
  IF AbortFlag THEN EXIT SUB

  CustFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM SHARED AS TaxFile LEN = TaxCustLen
  
  RptHandle = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #RptHandle

  TaxFile = FREEFILE
  OPEN TaxAdvFile FOR RANDOM AS TaxFile LEN = TaxAdvRecLen      'open data file
  NumAdvRec& = LOF(TaxFile) / TaxAdvRecLen
  
  
  ClearBack
  ShowProcessingScrn "Mailing Labels"
  
  FOR Cnt = 1 TO NumAdvRec&

    GET TaxFile, Cnt, TaxAdvRec(1)
    
    IF NOT TaxAdvRec(1).DelFlag THEN
      GET CustFile, TaxAdvRec(1).CustRec, TaxCustRec(1)
      NME$ = QPTrim$(TaxCustRec(1).FName) + " " + QPTrim$(TaxCustRec(1).LName)
      NME$ = QPTrim$(NME$)    'this one cleans up those with only last name
      PRINT #RptHandle, USING "#####"; TaxAdvRec(1).CustRec
      PRINT #RptHandle, NME$
      PRINT #RptHandle, TaxCustRec(1).ADDR1
      PRINT #RptHandle, TaxCustRec(1).ADDR2
      PRINT #RptHandle, TaxCustRec(1).CITY; " "; TaxCustRec(1).STATE; " "; TaxCustRec(1).ZIP
      PRINT #RptHandle,
      CustCnt = CustCnt + 1
    END IF
    ShowPctComp Cnt, NumAdvRec&
  NEXT
  
  PRINT #RptHandle, CHR$(18);   ' oki 320 10 cpi
  
  CLOSE         'Close all open files now
  
  IF DevSpec$ = "P" THEN
    EntryPoint = 4
  ELSEIF DevSpec$ = "S" THEN
    EntryPoint = 2
  ELSE
    EntryPoint = 1
  END IF
  
  ERASE Array, Frm, Form$, Fld, TaxCustRec
  
  PrintRptFile Header$, ReportFile$, LPTPort%, RetCode%, EntryPoint
  
  KILL ReportFile$
  
  EXIT SUB
  
PrintAlign:
  LPRINT STRING$(35, "X")
  LPRINT STRING$(35, "X")
  LPRINT STRING$(35, "X")
  LPRINT STRING$(35, "X")
  LPRINT STRING$(35, "X")
  LPRINT
  RETURN
  

END SUB

SUB AdvLookUp (RecNo&, Text$, CLSFlag%)
  
LookUpTop:
  
  REDIM Hlp$(1 TO 2)
  Hlp$(1) = "Enter an account number to look-up here."
  Hlp$(2) = "Enter all or part of the Customer Name here."
  
  SName$ = ""
  AcctNum& = 0
  LScrn = 2
  
  CursorOff
  
  REDIM ScrnArray(0)
  REDIM ScrnArray2(0)
  
  SaveScrn ScrnArray()
  
  REDIM LText(1 TO 2) AS STRING * 17
  
  MScrn = 2
  
  LText(1) = " Account Number:"
  LText(2) = "    Search Name:"
  
  LibName$ = "TAX"
  ScrnName$ = "LUPACCT"
  
  '--Initialize the form name array
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  Text$ = Text$ + " Look-Up"
  TextLen = LEN(Text$)
  TCol = ((80 - TextLen) \ 2)
  DisplayTaxScrn ScrnName$
  
  QPrintRC Text$, 8, TCol, -1
  
  GOSUB DisplayLookupText
  
  ShowCursor
  
  Action = 1
  FirstTime = True
  Frm(1).StayOnField = True
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = NOT FirstTime
      SELECT CASE LScrn
      CASE 1, 2
        LSET Form$(1, 0) = ""
        Fld(1).FType = 2
      END SELECT
      Form$(1, 1) = Hlp$(LScrn)
      Action = 1
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68, 13                'F10Key    Proceed with look up
      CursorOff
      SELECT CASE LScrn
      CASE 1    'account lookup
        AcctNum& = QPValL(Form$(1, 0))
        IF AcctNum& < 1 THEN
          Ok = MsgBox%("TAX.QSL", "BADACCTN")
        ELSE
          Search4AdvAcct AcctNum&, RecNo&
          IF RecNo& = 0 THEN
            Ok = MsgBox%("TAX.QSL", "BADACCTN")
          ELSE
            OKFlag = True
          END IF
        END IF
        Action = 1
      CASE 2    'Name lookup
        SName$ = LEFT$(QPTrim$(Form$(0, 0)), 10)
        IF LEN(SName$) = 0 THEN
          SName$ = SPACE$(10)
        END IF
        SaveScrn ScrnArray2()
        RestScrn ScrnArray()
        Search4AdvCust SName$, RecNo&, CLSFlag
        IF RecNo& > 0 THEN
          OKFlag = True
        ELSEIF RecNo& = 0 THEN
          Ok = MsgBox%("TAX.QSL", "NOMATCH")
        END IF
        RestScrn ScrnArray2()
        Action = 1
      END SELECT
    CASE -65    'F7Key
      IF LScrn < MScrn THEN
        LScrn = LScrn + 1
      ELSE
        LScrn = 1
      END IF
      LSET Form$(1, 0) = ""
      Action = 1
      FirstTime = True
      SaveField 0, Form$(), Fld(), BadField
      GOSUB DisplayLookupText
    CASE 27
      RecNo& = 0
      ExitFlag = True
    END SELECT
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 16
        SELECT CASE Frm(1).MCol
        CASE 22 TO 33           'ESC Cancel button
          PressButton 27, 16, 22, 33
        CASE 35 TO 45           'F7 Toggle Choice
          PressButton -65, 16, 35, 45
        CASE 47 TO 59           'F10 Save Button
          PressButton -68, 16, 47, 59
        END SELECT
      END SELECT
    END IF
    
  LOOP UNTIL ExitFlag OR OKFlag
  RestScrn ScrnArray()
  
  ERASE ScrnArray, ScrnArray2
  ERASE Frm, Form$, Fld, LText, Hlp$
  
  EXIT SUB
  
DisplayLookupText:
  QPrintRC LText(LScrn), 12, 15, -1
  RETURN
  
  
END SUB

SUB EditTrans
  
  RecDisp$ = SPACE$(14)
  SHARED Choice$()
  
  REDIM EdScrn(0)
  
  REDIM TaxAdvRec(1 TO 1) AS InterestRecType
  TaxAdvRecLen = LEN(TaxAdvRec(1))
  
  TaxFile = FREEFILE
  OPEN TaxAdvFile FOR RANDOM AS TaxFile LEN = TaxAdvRecLen      'open data file
  NumAdvRec& = LOF(TaxFile) / TaxAdvRecLen
  CLOSE TaxFile
  
  CursorOff
  ClearBack
  
  IF NumAdvRec& < 1 THEN
    CLOSE
    DisplayTaxScrn "ERRSCRN1"
    QPrintRC "No Ad-Cost Transaction to Edit!", 10, 25, -1
    QPrintRC "Press any key to continue.", 12, 28, -1
    WaitForAction
    GOTO EditIntExit
  ELSE
    GOSUB CheckAnyNotDeleted
    IF NoneToEdit THEN
      DisplayTaxScrn "ERRSCRN1"
      QPrintRC "No Active Interest Transactions!", 10, 25, -1
      QPrintRC "Press any key to continue.", 12, 28, -1
      WaitForAction
      GOTO EditIntExit
    END IF
  END IF
  
  LibName$ = "TAX"
  ScrnName$ = "TAXADVED"
  
  '--Initialize the form name array
  '--Get the total number of fields from all pages
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  
  '--define Quick Screen form editing arrays
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  '--for each screen, get first and last fields
  StartEl = 0
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  
  '--Clear all fields
  FOR F = 1 TO NumFlds
    LSET Form$(F, 0) = ""
  NEXT
  
  DisplayTaxScrn ScrnName$
  
  ShowCursor
  
  Action = 1
  FirstTime = True
  Frm(1).StayOnField = True
  
  DO
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF FirstTime THEN
      FirstTime = False
      Forward = True
      AdvRecNo& = 1
      QPrintRC "Editing Ad-Cost Information ", 6, 8, -1
      GOSUB GetAdvRec
    END IF
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE F10Key ', EnterKey    'F10   Save
      GOSUB UpDateAdvRec
    CASE F3Key
      Ok = MsgBox%("TAX", "DELADTRN")
      SELECT CASE Ok
      CASE 2    'if button 2 was pressed
        GOSUB DelAdvTrans
        GOSUB GetAdvRec
        'FirstTime = True
      CASE ELSE
      END SELECT
      Action = 1

    CASE F4KEY
      RecNo& = QPValL(Form$(1, 0))
      IF RecNo& > 0 THEN
        ShowCustHistory RecNo&
        Action = 1
      END IF
      
    CASE F5KEY  'F5 Look-up
      SaveScrn EdScrn()
      MPaintBox 6, 4, 18, 76, 8
      AdvLookUp FindRecNo&, "Ad-Cost Edit Transaction", CLSFlag
      RestScrn EdScrn()
      IF FindRecNo& <> AdvRecNo& AND FindRecNo& > 0 THEN
        AdvRecNo& = FindRecNo&
        GOSUB GetAdvRec
      END IF
      Action = 1
      
    CASE PgUpKey, -72               'previous rec
      Forward = False
      IF AdvRecNo& > 1 THEN
        AdvRecNo& = AdvRecNo& - 1
        GOSUB GetAdvRec
      END IF
    CASE PgDnKey, -80               'next rec
      Forward = True
      IF AdvRecNo& < NumAdvRec& THEN
        AdvRecNo& = AdvRecNo& + 1
        GOSUB GetAdvRec
      END IF

    CASE EscKey
      RecNo& = 0
      ExitFlag = True

    CASE -118, -117 'Ctrl-PgDn Ctrl-End
      Forward = True
      IF AdvRecNo& < NumAdvRec& THEN
        AdvRecNo& = NumAdvRec&
        GOSUB GetAdvRec
      END IF

    CASE -132, -119 'Ctrl-PgUp Ctrl-Home
      Forward = False
      IF AdvRecNo& > 1 THEN
        AdvRecNo& = 1
        GOSUB GetAdvRec
      END IF

    END SELECT
    
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 17
        SELECT CASE Frm(1).MCol
        CASE 10 TO 19           'ESC Cancel button
          PressButton EscKey, 17, 10, 19
        CASE 20 TO 30           'F3 Delete
          PressButton F3Key, 17, 20, 30
        CASE 31 TO 39           'F5 Search
          PressButton F4KEY, 17, 31, 39
        CASE 40 TO 48           'F5 Search
          PressButton F5KEY, 17, 40, 48
        CASE 49 TO 58           'F10 Save Button
          PressButton F10Key, 17, 49, 58
        CASE 59 TO 64           'PgUp Prev Rec
          PressButton PgUpKey, 17, 59, 64
        CASE 65 TO 70           'PgDn Next Rec
          PressButton PgDnKey, 17, 65, 70
        END SELECT
      END SELECT
    END IF
    
  LOOP UNTIL ExitFlag OR OKFlag
  
  GOTO EditIntExit
  
GetAdvRec:
  DidFirst = False
  DidLast = False
  TaxFile = FREEFILE
  OPEN TaxAdvFile FOR RANDOM AS TaxFile LEN = TaxAdvRecLen      'open data file
  DO
    GET TaxFile, AdvRecNo&, TaxAdvRec(1)
    IF TaxAdvRec(1).DelFlag <> 0 THEN
      SELECT CASE Forward
      CASE True
        IF AdvRecNo& < NumAdvRec& THEN
          AdvRecNo& = AdvRecNo& + 1
        ELSE
          Forward = False
          DidLast = True
        END IF
      CASE False
        IF AdvRecNo& > 1 THEN
          AdvRecNo& = AdvRecNo& - 1
        ELSE
          Forward = True
          DidFirst = True
        END IF
      END SELECT
    ELSE
      EXIT DO
    END IF
    IF DidFirst AND DidLast THEN
      CursorOff
      CLOSE TaxFile
      DisplayTaxScrn "ERRSCRN1"
      QPrintRC "No Active Ad-Cost Transactions!", 10, 25, -1
      QPrintRC "Press any key to continue.", 12, 28, -1
      WaitForAction
      GOTO EditIntExit
    END IF
  LOOP
  CLOSE TaxFile
  
  BCopy VARSEG(TaxAdvRec(1)), VARPTR(TaxAdvRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), LEN(Form$(0, 0)), 0
  UnPackBuffer 0, 0, Form$(), Fld()
  QPrintRC TaxAdvRec(1).InfoTxt, 14, 20, -1
  Action = 1
  GOSUB DisplayRecCnt
RETURN
  
CheckAnyNotDeleted:
  ShowProcessingScrn "Scanning for Active Transactions"
  NoneToEdit = True
  TaxFile = FREEFILE
  OPEN TaxAdvFile FOR RANDOM AS TaxFile LEN = TaxAdvRecLen      'open data file
  NumAdvRec& = LOF(TaxFile) / TaxAdvRecLen
  FOR Cnt& = 1 TO NumAdvRec&
    GET TaxFile, Cnt&, TaxAdvRec(1)
    IF TaxAdvRec(1).DelFlag = 0 THEN
      NoneToEdit = False
    END IF
    ShowPctCompL Cnt&, NumAdvRec&
  NEXT
  CLOSE TaxFile
  SmallPause
RETURN
  
UpDateAdvRec:
  SaveScrn EdScrn()
  TaxFile = FREEFILE
  OPEN TaxAdvFile FOR RANDOM AS TaxFile LEN = TaxAdvRecLen      'open data file
  GET TaxFile, AdvRecNo&, TaxAdvRec(1)
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(TaxAdvRec(1)), VARPTR(TaxAdvRec(1)), LEN(Form$(0, 0)), 0
  PUT TaxFile, AdvRecNo&, TaxAdvRec(1)
  CLOSE TaxFile
  DisplayTaxScrn "UPDATEOK"
  WaitForAction
  RestScrn EdScrn()
RETURN
  
DelAdvTrans:
  SaveScrn EdScrn()
  TaxFile = FREEFILE
  OPEN TaxAdvFile FOR RANDOM AS TaxFile LEN = TaxAdvRecLen      'open data file
  GET TaxFile, AdvRecNo&, TaxAdvRec(1)
  TaxAdvRec(1).DelFlag = True
  PUT TaxFile, AdvRecNo&, TaxAdvRec(1)
  CLOSE TaxFile
  DisplayTaxScrn "UPDATEOK"
  WaitForAction
  RestScrn EdScrn()
RETURN
  
DisplayRecCnt:
  LSET RecDisp$ = QPTrim$(STR$(AdvRecNo&)) + " of " + QPTrim$(STR$(NumAdvRec&))
  QPrintRC RecDisp$, 8, 56, -1
RETURN
  
EditIntExit:
  
END SUB

SUB PostTransactions

  REDIM TempScrn(0)
  
  REDIM TaxCust(1)  AS TaxCustType
  REDIM TaxTrans(1) AS TaxTransactionType
  REDIM TaxAdvRec(1 TO 1) AS InterestRecType
  TaxAdvRecLen = LEN(TaxAdvRec(1))
  AdvRecLen = LEN(TaxAdvRec(1))
  
  ClearBack
  ShowProcessingScrn "Verifying Ad-Cost Transactions"

  AdvFile = FREEFILE
  OPEN TaxAdvFile FOR RANDOM AS AdvFile LEN = AdvRecLen
  NumOfRec& = LOF(AdvFile) / AdvRecLen
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #AdvFile, Cnt&, TaxAdvRec(1)
    IF TaxAdvRec(1).DelFlag = 0 THEN
      DidSome = DidSome + 1
    END IF
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  CLOSE
  SmallPause

  IF DidSome < 1 THEN
    GOSUB ShowPostError
    GOTO PostExit
  END IF
  
  PostFlag = MsgBox%("TAX", "POSTINTR")
  
  SELECT CASE PostFlag
  CASE 1
    GOTO Postem
  CASE 2
    CursorOff
    ClearBack
    DisplayTaxScrn "POSTCAN"
    WaitForAction
    GOTO PostExit
  END SELECT
  
Postem:
  ClearBack
  ShowProcessingScrn "Posting Tax AddCost Transactions"
  
  'Open All Needed Files Here

  AdvFile = FREEFILE
  OPEN TaxAdvFile FOR RANDOM AS AdvFile LEN = AdvRecLen
  NumOfRec& = LOF(AdvFile) / AdvRecLen
  
  CustFile = FREEFILE
  OPEN TaxCustFile FOR RANDOM AS CustFile LEN = LEN(TaxCust(1))
  
  TransFile = FREEFILE
  OPEN TaxTransFile FOR RANDOM AS TransFile LEN = LEN(TaxTrans(1))
  
  FOR Cnt& = 1 TO NumOfRec&
    GET #AdvFile, Cnt&, TaxAdvRec(1)
    IF TaxAdvRec(1).DelFlag = 0 THEN
      'Update the Bill transaction first
      'TaxAdvRec(1).BillRec
      GET #TransFile, TaxAdvRec(1).BillRec, TaxTrans(1) 'get bill trans
      TaxTrans(1).Revenue.Collection = Round#(TaxTrans(1).Revenue.Collection + TaxAdvRec(1).Amount)
      PUT #TransFile, TaxAdvRec(1).BillRec, TaxTrans(1) 'put it back
      'Now make a new clean transaction
      REDIM TaxTrans(1) AS TaxTransactionType

      TaxTrans(1).TransDate = Date2Num%(DATE$)
      TaxTrans(1).TaxYear = TaxAdvRec(1).TaxYear
      TaxTrans(1).TranType = 6       '4=Interest
      TaxTrans(1).BillType = "A"     'R=Real P=Personal Property C=Combined (NC/GA)
      TaxTrans(1).Amount = TaxAdvRec(1).Amount  'Total Transaction Amount
      TaxTrans(1).Description = "Col/Ad Charge Bill#" + QPTrim$(TaxAdvRec(1).BillNumber)
      TaxTrans(1).Revenue.Principle1 = 0
      TaxTrans(1).Revenue.Principle2 = 0
      TaxTrans(1).Revenue.Principle3 = 0
      TaxTrans(1).Revenue.Principle4 = 0
      TaxTrans(1).Revenue.Principle5 = 0
      TaxTrans(1).Revenue.INTEREST = 0
      TaxTrans(1).Revenue.PENALTY = 0
      TaxTrans(1).Revenue.Collection = TaxAdvRec(1).Amount
      TaxTrans(1).Revenue.Future1 = 0
      TaxTrans(1).Revenue.Future2 = 0
      TaxTrans(1).Revenue.Principle1Pd = 0
      TaxTrans(1).Revenue.Principle2Pd = 0
      TaxTrans(1).Revenue.Principle3Pd = 0
      TaxTrans(1).Revenue.Principle4Pd = 0
      TaxTrans(1).Revenue.Principle5Pd = 0
      TaxTrans(1).Revenue.InterestPd = 0
      TaxTrans(1).Revenue.PenaltyPd = 0
      TaxTrans(1).Revenue.CollectionPd = 0
      TaxTrans(1).Revenue.Future1Pd = 0
      TaxTrans(1).Revenue.Future2Pd = 0
      TaxTrans(1).Posted2GL = "N"
      TaxTrans(1).CustomerRec = TaxAdvRec(1).CustRec
      TaxTrans(1).LastTrans = 0
      TaxTrans(1).BelongTo = TaxAdvRec(1).BillRec
      LSET TaxTrans(1).Padding = ""
      'Increment Transaction File Record Count
      NextRecord& = (LOF(TransFile) / LEN(TaxTrans(1))) + 1
      PUT #TransFile, NextRecord&, TaxTrans(1)
      'Update the Customer Pointers Now
      GET #CustFile, TaxAdvRec(1).CustRec, TaxCust(1)
      
      IF TaxCust(1).LastTrans = 0 THEN
        TaxCust(1).LastTrans = NextRecord&
        PUT #CustFile, TaxAdvRec(1).CustRec, TaxCust(1)
      ELSE
        Previous& = TaxCust(1).LastTrans
        TaxCust(1).LastTrans = NextRecord&
        PUT #CustFile, TaxAdvRec(1).CustRec, TaxCust(1)
        GET #TransFile, NextRecord&, TaxTrans(1)
        TaxTrans(1).LastTrans = Previous&
        PUT #TransFile, NextRecord&, TaxTrans(1)
      END IF
    END IF
    ShowPctComp Cnt&, NumOfRec&
  NEXT
  CLOSE

  ClearBack
  DisplayTaxScrn "UPDATEOK"
  WaitForAction
  
  'Now Delete the Tax Bill File so Duplicate's Cannot Be Reproduced

  KillFile TaxAdvFile
  
PostExit:
  EXIT SUB
  
  
ShowPostError:
  ClearBack
  DisplayTaxScrn "ERRSCRN1"
  QPrintRC "No Active Ad-Cost Transactions!", 10, 25, -1
  QPrintRC "Press any key to continue.", 12, 28, -1
  WaitForAction
RETURN
END SUB

SUB Search4AdvAcct (AcctNum&, RecNo&)
  
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()

  WazzWind 10, 22, 14, 58, 10, 2, True
  QPrintRC "Searching:    % Completed.", 12, 28, 14
  
  REDIM ITaxCust(1 TO 1) AS InterestRecType

  ITaxCustLen = LEN(ITaxCust(1))

  RecNo& = 0
  FOpenS TaxAdvFile, TaxFile    'open data file
  NumOfCust& = FLof&(TaxFile) / ITaxCustLen

  FOR CCnt& = 1 TO NumOfCust&
    FGetRTA TaxFile, ITaxCust(1), CCnt&, ITaxCustLen
    IF ITaxCust(1).CustRec = AcctNum& THEN
      RecNo& = CCnt&
      EXIT FOR
    END IF
    ShowSearchPCT CCnt&, NumOfCust&
  NEXT

  FClose TaxFile
  RestScrn ScrnArray()

  ERASE ScrnArray, ITaxCust

END SUB

SUB Search4AdvCust (Search$, RecNo&, CLSFlag%)
  
  STATIC Choice, LastSEARCH$
  ACCT$ = SPACE$(5)
  
  BNum$ = SPACE$(7)
  
  IF LastSEARCH$ <> Search$ THEN
    LastSEARCH$ = Search$
    Choice = 1
  END IF
  
  REDIM TScrnArray(0)
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  
  WazzWind 10, 22, 14, 58, 10, 2, True
  QPrintRC "Searching:    % Completed.", 12, 28, 14
  
  '091598 Found a bug that caused the lookup to get erroneous records
  '       if this was changed to greater than 32???
  CustBlock = 1
  
  REDIM MChoice(1 TO 1) AS FLen2
  REDIM ITaxCust(1 TO 1) AS InterestRecType
  
  ITaxCustLen = LEN(ITaxCust(1))
  
  SearchLen = LEN(Search$)
  Match = False
  
  FOpenS TaxAdvFile, TaxFile    'open data file
  NumOfCust& = FLof&(TaxFile) / ITaxCustLen
  
  FOR CCnt& = 1 TO NumOfCust&
    FGetRTA TaxFile, ITaxCust(1), CCnt&, ITaxCustLen
    IF INSTR(ITaxCust(1).CUSTNAME, Search$) > 0 AND ITaxCust(1).DelFlag = 0 THEN
      GOSUB CustLoadEM2
    END IF
DelSkip3:
    ShowSearchPCT CCnt&, NumOfCust&
  NEXT
  
  FClose TaxFile
  
  IF DCnt = 0 THEN
    RecNo& = 0
    GOTO ExitSearch2
  ELSE
    
    'SortT MChoice(1), DCnt, Direction%, LEN(MChoice(1).V), 0, 18
    
    MaxLen = 59 'Set menu width to zero
    Action = 0  '0 means stay in the menu until they select something
    
    IF Choice < 1 THEN
      Choice = 1                'Pre-load choice to highlight
    END IF
    
    Title$ = SPACE$(MaxLen + 4)
    LSET Title$ = "  Customer Name                  Year      Amt    BillNo  Acct"
    '--Find max menu width
    '--Center Menu within Screen
    Row = 4
    Col = ((80 - 60) \ 2) - 1
    
    IF CLSFlag THEN
      Row = 4
      BoxBot = 17               'limit the box length
      ClearBack
    ELSE
      Row = 6
      BoxBot = 14               'limit the box length to go no lower than line 20
      RestScrn ScrnArray()
    END IF
    
LoopRestart:
    LOCATE Row, Col, 0
    DO
      TitleBox BoxBot + 3, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
      QPrintRC "Matched:" + STR$(DCnt), BoxBot + 4, Col + 2, 15
      QPrintRC Title$, Row - 1, Col, 112
      MPaintBox Row, Col + MaxLen + 4, Row, Col + MaxLen + 5, 8
      VertMenuT2 MChoice(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo& = -1
        EXIT DO 'choice = 0
      END IF
      RecNo& = CVL(MID$(MChoice(Choice).V, 61, 4))
    LOOP UNTIL RecNo& > 0
  END IF
  
ExitSearch2:
  RestScrn ScrnArray()
  
  ERASE ScrnArray, MChoice, ITaxCust
  
  EXIT SUB
  
CustLoadEM2:
  
  DCnt = DCnt + 1
  REDIM PRESERVE MChoice(1 TO DCnt) AS FLen2
  LSET MChoice(DCnt).V = LEFT$(QPTrim$(ITaxCust(1).CUSTNAME), 30)
  
  MID$(MChoice(DCnt).V, 32) = FUsing$(STR$(ITaxCust(1).TaxYear), "####")
  MID$(MChoice(DCnt).V, 38) = FUsing$(STR$(ITaxCust(1).Amount), ",####.##")
  RSET BNum$ = QPTrim$(ITaxCust(1).BillNumber)
  MID$(MChoice(DCnt).V, 48) = BNum$
  RSET ACCT$ = QPTrim$(STR$(ITaxCust(1).CustRec))
  MID$(MChoice(DCnt).V, 55, 5) = ACCT$
  MID$(MChoice(DCnt).V, 61) = MKL$(CCnt&)
  RETURN
  
  
END SUB

