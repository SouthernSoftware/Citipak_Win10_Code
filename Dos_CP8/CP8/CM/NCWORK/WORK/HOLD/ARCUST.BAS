 DEFINT A-Z
DECLARE FUNCTION GetNumOfAcct% ()
DECLARE SUB SaveOldCustRec (Form$(), RecNo%)
DECLARE SUB CustLookUP (RecNo%)
DECLARE SUB SearchGetCust (Search$, RecNo%)
DECLARE SUB PrintCustList ()
DECLARE SUB SaveNewCustRec (Form$())
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE FUNCTION Compare3% (BYVAL Seg1, BYVAL Adr1, BYVAL Seg2, BYVAL Adr2, NumBytes)
DECLARE SUB InsertT (SEG StartElement AS ANY, ElSize%, NumEls%)
DECLARE SUB EditCustomer ()
DECLARE SUB AddCustomer ()
DECLARE SUB LoadMeterRec (Form$())
DECLARE SUB AddNewMeter (Cust$, EdFlag)
DECLARE SUB EditMeter ()
DECLARE SUB SaveMeterRec (Form$())
DECLARE SUB SaveScrn (Array%())
DECLARE SUB RestScrn (Array%())
DECLARE SUB AddEditCustomer (EdFlag%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB SaveCustRec (Form$())
DECLARE SUB AddNewCustomer ()
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB SaveSetUpRec (Form$())
DECLARE SUB LoadSetUpRec (Form$())
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB EditCust ()
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()
DECLARE SUB PressButton (BYVAL KeyCode, BYVAL ButtonRow, BYVAL ButtonLCol, BYVAL ButtonRCol)
DECLARE SUB StuffBuf (Ky$)
DECLARE FUNCTION FileSize& (FileName$)
  
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
  
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB WaitForAction ()
  
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB FOpenS (FileName$, Handle%)
DECLARE SUB FClose (Handle%)
  
DECLARE SUB FGetRTA (Handle%, SEG Dest AS ANY, RecNo&, RecSize%)
DECLARE SUB FPutAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB HideCursor ()
DECLARE FUNCTION QPValL& (Number$)
DECLARE SUB LoadCustRec (Form$(), RecNo%, TFormLen%, LockedFlag%)
  
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'PageInfo.BI'                  'Form Page info
  
  '$INCLUDE: 'QScr.BI'                      'QuickScreen Declarations
  
  '$INCLUDE: 'SetCnf.bi'
  '$INCLUDE: 'ubcust.bi'
  ''$INCLUDE: 'ubmeter.BI'
  
  CONST False = 0, True = NOT False
  
  STACK 8000
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 5)
  
  MChoice$(1) = "Add a New Customer"
  MChoice$(2) = "Edit Existing Customer  "
  MChoice$(3) = "Print Customer Listing  "
  'MChoice$(4) = "Print Customer Listing  "
  MChoice$(4) = "Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 20   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  Help$ = "Customer Maintenance Menu"
  
  IF COMMAND$ = "ADD" THEN GOTO There
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    LibFile2Scrn "UB.QSL", "MENUBAK", MonoCode, -1, ErrorCode
    
    TitleBox 3, Col, MaxLen + 3, "Customer Menu ", Cnf
    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    PrintTitle user$
    PrintHelp Help$
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      AddEditCustomer 0
    CASE 2
      CustLookUP RecNo
      IF RecNo > 0 THEN
        AddEditCustomer RecNo
      END IF
    CASE 3
      'ove
'    CASE 4
'      PrintCustList
    CASE 4
      CLS
      END
    END SELECT
  LOOP
  
  RUN "armenu"
  
There:
  
  RANDOMIZE TIMER
  CLS
  REDIM Form$(0, 0)
  REDIM UBCustRec(1) AS UBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  Stop$ = CHR$(27)
  
  Form$(0, 0) = STRING$(UBCustRecLen - 2, 0)
  LINE INPUT "Enter number or recs:"; recs$
  Amt = VAL(recs$)
  FOR zz = 1 TO Amt
    '      zzz$ = SPACE$(50)
    FOR cnt = 1 TO 10
      MID$(Form$(0, 0), cnt + 8, 1) = CHR$(INT(26 * RND) + 65)
    NEXT
    St! = TIMER
    SaveNewCustRec Form$()
    '      LOCATE 1, 1: PRINT zzz$;
    '      LOCATE 1, 1
    PRINT MID$(Form$(0, 0), 9, 10), GetNumOfAcct%, zz, USING "#.########"; TIMER - St!
    a$ = INKEY$
    IF a$ = Stop$ THEN EXIT FOR
  NEXT
  '  END

SUB AddEditCustomer (RecNo)
  
  REDIM ScrnArray(0)
  
  LibName$ = "UB"
  FormName$ = "ARCUST"
  ScrnName$ = "ARCUST"
  NumScrns = 4



  
  '--define the multi-choice fields
  SHARED Choice$()
'  REDIM Choice$(0 TO 4, 0)

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

  Action = 1


  
  
  
  'IF RecNo > 0 THEN
  '  LoadCustRec Form$(), RecNo, TFormLen, LockedFlag
  '  UnPackBuffer 0, 0, Form$(), Fld()
  'END IF
  
  '--Set screen number to one and display screen
  Scr = 1
  LibFile2Scrn LibName$, FormName$, MonoCode, Attribute%, ErrCode
  
  ShowCursor
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      
      SaveScrn ScrnArray()
      IF RecNo = 0 THEN
        SaveNewCustRec Form$()
      ELSE
        SaveOldCustRec Form$(), RecNo
      END IF
      DisplayUBScrn "UPDATEOK"
      WaitForAction
      ExitFlag = True
      RestScrn ScrnArray()
      
    END SELECT
    
    '--check for mouse clicks on buttons not attached to the form
    IF Frm(1).Presses THEN
      SELECT CASE Frm(1).MRow
      CASE 21   'Look for the f10 or esc button
        SELECT CASE Frm(1).MCol
        CASE 43 TO 54           '--Save Button
          PressButton -68, 21, 43, 54
        CASE 62 TO 72           '--cancel button
          PressButton 27, 21, 57, 70
        END SELECT
      END SELECT                'row
    END IF
    
    'add
    '    IF Frm(1).PrevFld = BillWhoFld AND Frm(1).FldNo <> BillWhoFld THEN
    '      SELECT CASE QPTrim$(Form$(BillWhoFld, 0))
    '      CASE "O"
    '        GOSUB UnProPage3
    '      CASE ELSE
    '        GOSUB ProPage3
    '      END SELECT
    '    END IF
    
    '--Check screen page
    
      LibFile2Scrn LibName$, FormName$, MonoCode, Attribute, ErrCode
    
    
  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag
  
  '  IF Frm(1).KeyCode = 27 THEN GOTO ExitCustEdit
  '  REDIM MTChoice$(1 TO 4)
  '
  '  MTChoice$(1) = "Add Metered Service"
  '  MTChoice$(2) = "Edit Existing Metered Service"
  '  MTChoice$(3) = "Flat Rate Only"
  '
  '  MaxLen = 0    'Set menu width to zero
  '  BoxBot = 20   'limit the box length to go no lower than line 20
  '  Action = 0    '0 means stay in the menu until they select something
  '  Choice = 1    'Pre-load choice to highlight
  '
  '  '--Find max menu width
  '  FOR cnt = 1 TO UBOUND(MTChoice$)
  '    TLen = LEN(MTChoice$(cnt))
  '    IF TLen > MaxLen THEN
  '      MaxLen = TLen
  '    END IF
  '  NEXT
  '
  '  '--Center Menu within Screen
  '  Row = ((25 - (UBOUND(MTChoice$))) \ 2) - 1
  '  Col = ((80 - MaxLen) \ 2) - 2
  '
  '  DO
  '    LOCATE Row, Col, 0
  '    ShowCursor
  '    VertMenu MTChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
  '
  '    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
  '
  '    SELECT CASE Choice
  '    CASE 1
  '      AddNewMeter Form$(2, 0), FALSE
  '      RestScrn ScrnArray()
  '      ExitFlag = FALSE
  '    CASE 2
  '      AddNewMeter Form$(2, 0), True
  '      RestScrn ScrnArray()
  '      ExitFlag = FALSE
  '    CASE 3
  '    END SELECT
  '  LOOP UNTIL Choice = 3 OR ExitFlag = True
  
ExitCustEdit:
  HideCursor
  EXIT SUB
  
  'UnProPage3:
  '  FOR TCnt = PageInfo(3).FirstFld TO PageInfo(3).LastFld
  '    Fld(TCnt).Protected = FALSE
  '  NEXT
  '  RETURN
  
  'ProPage3:
  '  FOR TCnt = PageInfo(3).FirstFld TO PageInfo(3).LastFld
  '    Fld(TCnt).Protected = True
  '  NEXT
  '  RETURN
  
END SUB

SUB AddNewMeter (Cust$, EdFlag)
  
  LibName$ = "UB"
  ScrnName$ = "UBMETER"
  
  '--define the multi-choice fields
  SHARED Choice$()
  REDIM Choice$(0 TO 4, 0)
  REDIM ScrnArray(0)
  '--Initialize the form name array
  '--Get the total number of fields from all pages
  
  NumFlds = LibNumberOfFields(LibName$, ScrnName$)
  '--define Quick Screen form editing arrays
  
  REDIM Frm(1) AS FormInfo
  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  
  StartEl = 0
  
  LibGetFldDef LibName$, ScrnName$, StartEl, Fld(), Form$(), ErrCode
  LibFile2Scrn LibName$, ScrnName$, MonoCode, Attribute%, ErrCode
  
  ShowCursor
  QPrintRC QPTrim$(Cust$), 4, 30, -1
  Action = 1
  
  IF EdFlag THEN
    'LoadMeterRec Form$()
    UnPackBuffer 0, 0, Form$(), Fld()
  END IF
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      SaveScrn ScrnArray()
      SaveMeterRec Form$()
      ExitFlag = True
      RestScrn ScrnArray()
      
    END SELECT
    
    
    'add
    
  LOOP UNTIL Frm(1).KeyCode = 27 OR ExitFlag
  
  
  
  
END SUB

SUB CustLookUP (RecNo)
  REDIM ScrnArray(0)
  
  REDIM LText$(4)
  
  LScrn = 1
  MScrn = 1
  
  LText$(1) = " Account Number:"
  LText$(2) = "    Search Name:"
  LText$(3) = "   Meter Number:"
  LText$(4) = "Service Address:"
  
  LibName$ = "UB"
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
  
  DisplayUBScrn "MENUBAK"
  DisplayUBScrn ScrnName$
  
  GOSUB DisplayLookupText
  
  ShowCursor
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    '--Check for Key presses
    SELECT CASE Frm(1).KeyCode
    CASE -68    'F10Key
      Search$ = QPTrim$(Form$(0, 0))
      SearchAcct& = QPValL(Search$)
      
      SELECT CASE LScrn
      CASE 1    'Customer account lookup
        IF SearchAcct& < 1 OR SearchAcct& > GetNumOfAcct THEN
          OK = MsgBox%("UB.QSL", "BADACCTN")
        ELSE
          RecNo = SearchAcct&
          OKFlag = True
        END IF
      CASE 2    'Customer Name lookup
        IF LEN(Search$) < 1 THEN
          OK = MsgBox%("UB.QSL", "BADSEARH")
          Action = 1
          Frm(1).FldNo = 1
        ELSE
          SearchGetCust Search$, RecNo
          IF RecNo > 0 THEN
            OKFlag = True
          ELSEIF RecNo = 0 THEN
            OK = MsgBox%("UB.QSL", "NOMATCH")
          END IF
          
        END IF
      END SELECT
      
    CASE -65    'F7Key
      IF LScrn < MScrn THEN
        LScrn = LScrn + 1
      ELSE
        LScrn = 1
      END IF
      LSET Form$(1, 0) = ""
      Action = 1
      SaveField 0, Form$(), Fld(), BadField
      GOSUB DisplayLookupText
      
    CASE 27
      RecNo = 0
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
  EXIT SUB
  
DisplayLookupText:
  QPrintRC LText$(LScrn), 12, 22, -1
  RETURN
  
  
END SUB

SUB EditMeter
  
END SUB

SUB LoadCustRec (Form$(), RecNo, TFormLen, LockedFlag)
  
  REDIM UBCustRec(1) AS UBCustRecType
  
  UBCustRecLen = LEN(UBCustRec(1))
  
  Form$(0, 0) = SPACE$(TFormLen)
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM AS UBFile LEN = UBCustRecLen
  GET UBFile, RecNo, UBCustRec(1)
  CLOSE
  'IF UBCustRec(1).LockedFlag THEN
  '  LockedFlag = True
  '  GOTO ExitCustLoad
  'ELSE
  '  LockedFlag = False
  'END IF
  
  BCopy VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), TFormLen, 0
  
ExitCustLoad:
  ERASE UBCustRec
  
END SUB

SUB LoadMeterRec (Form$())
  
  'REDIM UBMeterRec(1) AS UBMeterRecType
  
  UBMeterRecLen = LEN(UBMeterRec(1))
  
  Form$(0, 0) = SPACE$(UBMeterRecLen)
  '
  UBFile = FREEFILE
  OPEN "UBMETER.DAT" FOR RANDOM AS UBFile LEN = UBMeterRecLen
  GET UBFile, 1, UBMeterRec(1)
  CLOSE
  
  BCopy VARSEG(UBMeterRec(1)), VARPTR(UBMeterRec(1)), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), UBMeterRecLen, 0
  
  ERASE UBMeterRec
  
END SUB

SUB PrintCustList
  REDIM UBCustRec(1) AS UBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  Stop$ = CHR$(27)
  
  IdxRecLen = 2 'we are using a integer
  IdxFileSize& = FileSize("UBCUSTBK.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM AS UBFile LEN = UBCustRecLen
  
  FOR cnt = 1 TO IdxNumOfRecs
    
    GET UBFile, IdxBuff(cnt).RecNum, UBCustRec(1)
    'PRINT UBCustRec(1).Book; "-"; UBCustRec(1).SEQNUMB
    'IF INSTR(UBCustRec(1).SearchN, "WATSON") THEN STOP
    '  SLEEP
    a$ = INKEY$
    IF a$ = Stop$ THEN EXIT FOR
    
  NEXT
  
  
  CLOSE
  
  
END SUB

SUB SaveMeterRec (Form$())
  
  'REDIM UBMeterRec(1) AS UBMeterRecType
  UBMeterRecLen = LEN(UBMeterRec(1))
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBMeterRec(1)), VARPTR(UBMeterRec(1)), UBMeterRecLen, 0
  
  UBFile = FREEFILE
  OPEN "UBMETER.DAT" FOR RANDOM AS UBFile LEN = UBMeterRecLen
  PUT UBFile, 1, UBMeterRec(1)
  CLOSE
  LibFile2Scrn "UB", "UPDATEOK", MonoCode, Attribute%, ErrCode
  WaitForAction
  
END SUB

SUB SaveNewCustRec (Form$())
  
  REDIM UBCustRec(1) AS UBCustRecType
  
  FormLen = LEN(Form$(0, 0))
  
  UBCustRecLen = LEN(UBCustRec(1))
  
  NumOfRecs& = FileSize("UBCUST.DAT") \ UBCustRecLen
  
  NextRec& = NumOfRecs& + 1
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1)), FormLen, 0
  
  'make sure the lockedflag is false
  'UBCustRec(1).LockedFlag = False
  
  UBFile = FREEFILE
  OPEN "UBCUST.DAT" FOR RANDOM AS UBFile LEN = UBCustRecLen
  PUT UBFile, NextRec&, UBCustRec(1)
  CLOSE UBFile
  
  'Update Cust Name Search index -\-\-\-\-\-\-\-
  
  IdxRecLen = 2 'we are using a integer
  
  IdxFileSize& = FileSize("UBCUSTNM.IDX")
  
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  NextFreeRec = IdxNumOfRecs + 1
  
  REDIM IdxBuff(1 TO IdxNumOfRecs + 1) AS UBCustIndexRecType
  'dim array to num of recs +1
  
  REDIM TempUBCustRec(1) AS UBCustRecType
  
  IF IdxNumOfRecs > 0 THEN      'if this is not the first
    FGetAH "UBCUSTNM.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs  'load it
    
    FOpenS "UBCUST.DAT", C1Handle               'open data file
    
    LowOffSet = 0               '
    TopOffSet = IdxNumOfRecs + 1
    RecOffSet = IdxNumOfRecs \ 2
    
    IF RecOffSet < 1 THEN RecOffSet = 1
    
    DO
      
      FGetRTA C1Handle, TempUBCustRec(1), CLNG(IdxBuff(RecOffSet).RecNum), UBCustRecLen
      'get the rec from the data file based off the index rec num
      
      SELECT CASE Compare3%(VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1).SEARCHN), VARSEG(TempUBCustRec(1)), VARPTR(TempUBCustRec(1).SEARCHN), 10)
        'compare the rec we just read to new cust we just added
      CASE -1   'New search name comes before this one
        TopOffSet = RecOffSet
        RecOffSet = RecOffSet \ 2
        IF RecOffSet < LowOffSet THEN
          RecOffSet = LowOffSet + ((TopOffSet - LowOffSet) \ 2)
        END IF
      CASE 1    'New search name comes after this one
        LowOffSet = RecOffSet
        RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
        
      CASE 0    'New search name is the same as this one
        SELECT CASE Compare3%(VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1).FULLNAME), VARSEG(TempUBCustRec(1)), VARPTR(TempUBCustRec(1).FULLNAME), 35)
          'so we will examine the whole name
        CASE -1 'Lower
          TopOffSet = RecOffSet
          RecOffSet = RecOffSet \ 2
          IF RecOffSet < LowOffSet THEN
            RecOffSet = LowOffSet + ((TopOffSet - LowOffSet) \ 2)
          END IF
        CASE 1  'Higher
          LowOffSet = RecOffSet
          RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
        CASE 0
          LowOffSet = RecOffSet
          RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
        END SELECT
      END SELECT
    LOOP WHILE TopOffSet - LowOffSet > 1
    
    InsertRecAt = TopOffSet
    
    AmtToMove = (IdxNumOfRecs + 1) - TopOffSet
    
    IF AmtToMove THEN
      InsertT IdxBuff(InsertRecAt), IdxRecLen, AmtToMove
    END IF
    
    IdxBuff(InsertRecAt).RecNum = NextFreeRec
    
    FPutAH "UBCUSTNM.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs + 1
    FClose C1Handle
    
  ELSE
    IdxBuff(1).RecNum = NextFreeRec
    FPutAH "UBCUSTNM.IDX", IdxBuff(1), IdxRecLen, 1
  END IF
  
  ERASE IdxBuff
  'Done with name index
  
  'Update Cust Book/Route index -\-\-\-\-\-\-\-
  
  IdxRecLen = 2 'we are using a integer
  
  IdxFileSize& = FileSize("UBCUSTBK.IDX")
  
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  NextFreeRec = IdxNumOfRecs + 1
  
  REDIM IdxBuff(1 TO IdxNumOfRecs + 1) AS UBCustIndexRecType
  'dim array to num of recs +1
  
  REDIM TempUBCustRec(1) AS UBCustRecType
  
  IF IdxNumOfRecs > 0 THEN      'if this is not the first
    FGetAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs  'load it
    
    FOpenS "UBCUST.DAT", C1Handle               'open data file
    
    LowOffSet = 0               '
    TopOffSet = IdxNumOfRecs + 1
    RecOffSet = IdxNumOfRecs \ 2
    
    IF RecOffSet < 1 THEN RecOffSet = 1
    
    DO
      
      FGetRTA C1Handle, TempUBCustRec(1), CLNG(IdxBuff(RecOffSet).RecNum), UBCustRecLen
      'get the rec from the data file based off the index rec num
      
      SELECT CASE Compare3%(VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1).Book), VARSEG(TempUBCustRec(1)), VARPTR(TempUBCustRec(1).Book), 8)
        'compare the rec we just read to new cust we just added
      CASE -1   'New Route is before this one
        TopOffSet = RecOffSet
        RecOffSet = RecOffSet \ 2
        IF RecOffSet < LowOffSet THEN
          RecOffSet = LowOffSet + ((TopOffSet - LowOffSet) \ 2)
        END IF
      CASE 1    'New Route is after this one
        LowOffSet = RecOffSet
        RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
        
      CASE 0    'This Should never occure
        'STOP
        LowOffSet = RecOffSet
        RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
        
        '        SELECT CASE Compare3%(VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1).FULLNAME), VARSEG(TempUBCustRec(1)), VARPTR(TempUBCustRec(1).FULLNAME), 35)
        '          'so we will examine the whole name
        '        CASE -1 'Lower
        '          TopOffSet = RecOffSet
        '          RecOffSet = RecOffSet \ 2
        '          IF RecOffSet < LowOffSet THEN
        '            RecOffSet = LowOffSet + ((TopOffSet - LowOffSet) \ 2)
        '          END IF
        '        CASE 1  'Higher
        '          LowOffSet = RecOffSet
        '          RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
        '        CASE 0
        '          LowOffSet = RecOffSet
        '          RecOffSet = RecOffSet + ((TopOffSet - RecOffSet) \ 2)
        '        END SELECT
      END SELECT
    LOOP WHILE TopOffSet - LowOffSet > 1
    
    InsertRecAt = TopOffSet
    
    AmtToMove = (IdxNumOfRecs + 1) - TopOffSet
    
    IF AmtToMove THEN
      InsertT IdxBuff(InsertRecAt), IdxRecLen, AmtToMove
    END IF
    
    IdxBuff(InsertRecAt).RecNum = NextFreeRec
    
    FPutAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs + 1
    FClose C1Handle
    
  ELSE
    IdxBuff(1).RecNum = NextFreeRec
    FPutAH "UBCUSTBK.IDX", IdxBuff(1), IdxRecLen, 1
  END IF
  
  
  ERASE UBCustRec, TempUBCustRec, IdxBuff
  
END SUB

SUB SaveOldCustRec (Form$(), RecNo)
  
  REDIM UBCustRec(1) AS UBCustRecType
  
  UBCustRecLen = LEN(UBCustRec(1))
  
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(UBCustRec(1)), VARPTR(UBCustRec(1)), UBCustRecLen, 0
  
  UBFile = FREEFILE
  
  OPEN "UBCUST.DAT" FOR RANDOM AS UBFile LEN = UBCustRecLen
  PUT UBFile, RecNo, UBCustRec(1)
  CLOSE
  
END SUB

SUB SearchGetCust (Search$, RecNo)
  
  REDIM ScrnArray(0)
  SaveScrn ScrnArray()
  wheel$ = "/|\|"
  WPos = 1
  
  DisplayUBScrn "SHOWSCRH"
  
  REDIM UBCustRec(1) AS UBCustRecType
  UBCustRecLen = LEN(UBCustRec(1))
  
  IdxRecLen = 2 'we are using a integer
  IdxFileSize& = FileSize("UBCUSTNM.IDX")
  IdxNumOfRecs = IdxFileSize& \ IdxRecLen
  
  REDIM IdxBuff(1 TO IdxNumOfRecs) AS UBCustIndexRecType
  
  FGetAH "UBCUSTNM.IDX", IdxBuff(1), IdxRecLen, IdxNumOfRecs    'load it
  
  SearchLen = LEN(Search$)
  
  Match = False
  FirstRec = 1
  LastRec = IdxNumOfRecs
  
  BotOffSet = 1
  TopOffSet = IdxNumOfRecs
  
  FOpenS "UBCUST.DAT", C1Handle 'open data file
  
  'Find matching record
  
  MidRec = (LastRec + FirstRec) \ 2
  
  DO
    
    FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(MidRec).RecNum), UBCustRecLen
    UBSearchN$ = LEFT$(UBCustRec(1).SEARCHN, SearchLen)
    GOSUB ShowSearchThing
    
    IF Search$ = UBSearchN$ THEN
      IF MidRec - BotOffSet > 1 THEN
        MidRec = MidRec - 1
      ELSE
        FirstMatchRec = MidRec
      END IF
    ELSEIF Search$ < UBSearchN$ THEN            'lower
      TopOffSet = MidRec
      MidRec = TopOffSet - ((TopOffSet - BotOffSet) \ 2)
    ELSE        'higher
      BotOffSet = MidRec
      MidRec = BotOffSet + ((TopOffSet - BotOffSet) \ 2)
      IF MidRec = BotOffSet THEN EXIT DO
    END IF
    IF TopOffSet = BotOffSet + 1 THEN EXIT DO
  LOOP UNTIL FirstMatchRec
  
  IF FirstMatchRec > 0 THEN
    MidRec = FirstMatchRec
    DO UNTIL MidRec + 1 > IdxNumOfRecs
      MidRec = MidRec + 1
      FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(MidRec).RecNum), UBCustRecLen
      UBSearchN$ = LEFT$(UBCustRec(1).SEARCHN, SearchLen)
      IF Search$ <> UBSearchN$ THEN
        LastMatchRec = MidRec - 1
        EXIT DO
      END IF
    LOOP
  ELSE
    RecNo = 0
    GOTO ExitSearch
  END IF
  
  TotalMatch = (LastMatchRec - FirstMatchRec) + 1
  
  IF TotalMatch > 1 THEN
    REDIM MChoice$(1 TO TotalMatch)
    FOR cnt = FirstMatchRec TO LastMatchRec
      GOSUB ShowSearchThing
      DCnt = DCnt + 1
      FGetRTA C1Handle, UBCustRec(1), CLNG(IdxBuff(cnt).RecNum), UBCustRecLen
      MChoice$(DCnt) = SPACE$(64)
      LSET MChoice$(DCnt) = UBCustRec(1).SEARCHN
      MID$(MChoice$(DCnt), 14, 30) = UBCustRec(1).CADDR1
      MID$(MChoice$(DCnt), 49) = MKI$(IdxBuff(cnt).RecNum)
    NEXT
    
    MaxLen = 46 'Set menu width to zero
    BoxBot = 20 'limit the box length to go no lower than line 20
    Action = 0  '0 means stay in the menu until they select something
    Choice = 1  'Pre-load choice to highlight
    
    '--Find max menu width
    
    '--Center Menu within Screen
    Row = 7
    Col = 16
    BoxBot = 19
    DisplayUBScrn "MENUBAK"
    DO
      LOCATE Row, Col, 0
      ShowCursor
      VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        RecNo = -1
        EXIT DO 'choice = 0
      END IF
      RecNo = CVI(MID$(MChoice$(Choice), 49, 2))
    LOOP UNTIL RecNo > 0
  ELSE
    RecNo = IdxBuff(FirstMatchRec).RecNum
  END IF
  
ExitSearch:
  RestScrn ScrnArray()
  FClose C1Handle
  EXIT SUB
  
ShowSearchThing:
  WPos = WPos + 1
  IF WPos > 4 THEN WPos = 1
  QPrintRC MID$(wheel$, WPos, 1), 12, 44, -1
  RETURN
  
END SUB

