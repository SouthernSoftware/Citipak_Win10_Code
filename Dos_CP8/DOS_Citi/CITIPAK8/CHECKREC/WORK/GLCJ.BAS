DECLARE SUB PostTrans2 (CJType%)
DECLARE SUB GetBankList (BankList$(), NumBanks%)
DEFINT A-Z

DECLARE SUB ClrBG ()
DECLARE SUB MnuTitle (Title$, X%, Y%, Wdth%, TitColor%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
'DECLARE SUB Prn2Scrn2 (TitleText$(), FileName$, TopRow%, BotRow%, Clr%, PrnFlag%)
DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB FileView (FileName$, Ky, Action, FVI AS ANY, SEG Array)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (T$, R%, c%, Clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()

'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'pageinfo.BI'                  'Form Page info

'$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
'$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'RPTPRN.BI'                    'Report printer routines
'$INCLUDE: 'GL.BI'                        'GL File Types
'$INCLUDE: 'CASHJRNL.BI'
'$INCLUDE: 'OSCHK.BI'
'$INCLUDE: 'SCRAUX.bi'
  
DECLARE SUB CJMenu (JrnlType%)
DECLARE SUB EdEntry (JrnlCode%)
DECLARE SUB OpenCJEditFile (CJEditFile%, NumEdTrans%, CJType)
DECLARE SUB NewCJEntry ()
DECLARE FUNCTION GetNextRec% (FileNum%, NumRecs%, CurrRec%, NextRec%)
DECLARE SUB PrnEditList (CJType)
DECLARE SUB DisplayTotals ()
DECLARE SUB PostTrans (CJType)
DECLARE SUB Check4BlankFlds (Form$(), BadFld%, LPDate%, HPDate%, CJType)
DECLARE FUNCTION ListRecs% (CJType)
DECLARE SUB ResetForm (Form$(), NumFlds%)
DECLARE SUB ShowEdMode (EdMode%)
DECLARE SUB Rec2Form (Form$(), RecordNumber%, Fld() AS ANY, CJType)
DECLARE SUB SaveTrans (Form$(), RecordNumber%, EdMode%, CJType)
'DECLARE SUB PressButton2 (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB PressButton (BYVAL KeyCode%, BYVAL ButtonRow%, BYVAL ButtonLCol%, BYVAL ButtonRCol%)
DECLARE SUB GetFundList (FundList$(), NumFunds%)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
DECLARE SUB GetPostDates (LPDate, HPDate)
DECLARE FUNCTION EDCheck4Acct% (Form$(), AcctFld%)
DECLARE FUNCTION GetBankGLAcct$ (BankNum)

  TYPE FLen
    V AS STRING * 42
  END TYPE

DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)

CONST False = 0, True = NOT False

DIM SHARED AcctIdx AS GLAcctIndexType
DIM SHARED acct AS GLAcctRecType
DIM SHARED CJEdit AS CJEditRecType

  'DATA 8,12,16,20,24,28,32,36,40,44,48,52
  'DATA 57,61,65,69,73,77,81,85,89,93,97,101
  'DATA 106,110,114,118,122,126,130,134,138,142,146,150

DIM SHARED User$, CashAcct$, APAcct$, EncAcct$, FBAcct$, CDCash$, CDDue$, CDActive$, CRCashAcct$
DIM SHARED TotLen, FundLen, AcctLen, DetLen

''$INCLUDE: 'init.BI'

SUB Check4BlankFlds (Form$(), BadFld, LPDate, HPDate, CJType)
  
  TDate = Date2Num(Form$(2, 0))
  IF TDate < LPDate OR TDate > HPDate THEN
    BadFld = 2
    EXIT SUB
  END IF
  
  FOR F = 3 TO 5
    IF LEN(QPTrim$(Form$(F, 0))) = 0 THEN
      BadFld = F
      EXIT FOR
    END IF
  NEXT

  IF BadFld = 4 AND CJType = 1 THEN
    BadFld = 0
  END IF

  IF LEN(QPTrim$(Form$(7, 0))) = 0 THEN
    BadFld = 7
  END IF




END SUB

SUB CJMenu (JrnlType)

  '$INCLUDE: 'init.bi'

  TextCursor -2, -2             'Define the Mouse cursor
  DIM MenuText$(1 TO 4)
  SELECT CASE JrnlType
  CASE 1
    Help$ = "Cash Receipts Journal"
    MenuText$(1) = "Enter/Edit Cash Receipts"
    MenuText$(2) = "Print Cash Receipts Journal"
    MenuText$(3) = "Post Entries"
    MenuText$(4) = "Quit"
  CASE 2
    Help$ = "Cash Disbursements Journal"
    MenuText$(1) = "Enter/Edit Cash Disbursements"
    MenuText$(2) = "Print Cash Disbursements Journal"
    MenuText$(3) = "Post Entries"
    MenuText$(4) = "Quit"
  END SELECT

  MaxLen = 0  'Set menu width to zero
  BoxBot = 18   'limit the box to line 20
  Action = 0    '0 means stay in the menu
  Choice = 1    'Pre-load choice to highlight
  
  FOR Cnt = 1 TO UBOUND(MenuText$)              'Find max menu width
    TLen = LEN(MenuText$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  'Center Menu within Screen
  Row = ((25 - (UBOUND(MenuText$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  ShowCursor

  DO
    LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, -1, ErrorCode
    MnuTitle Help$, Row, Col, MaxLen, 15
    PrintTitle User$
    PrintHelp Help$

    LOCATE Row, Col, 0
    VertMenu MenuText$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

    IF Ky$ = CHR$(27) THEN EXIT DO
    SELECT CASE Choice
    CASE 1
      ClrBG
      EdEntry JrnlType
    CASE 2
      ClrBG
      PrnEditList JrnlType
    CASE 3
      ClrBG
      PostTrans2 JrnlType
    CASE 4
      CLS : END
    CASE ELSE
    END SELECT
  LOOP

  ClrBG
  DispLoad

  
END SUB

FUNCTION EDCheck4Acct% (Form$(), AcctFld%)

  SHARED TotLen

  List$ = SPACE$(65)

  Lookup$ = QPTrim$(UCASE$(Form$(AcctFld, 0)))
  Lookuplen = LEN(Lookup$)

  REDIM ScrBuf(2000)
  MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen

  OpenAcctIdx AcctIdxFileNum, NumIdxRecs
  CLOSE AcctIdxFileNum
  REDIM TempAcctIdx(1 TO NumIdxRecs) AS GLAcctIndexType

  ElSize = LEN(TempAcctIdx(1))
  FGetAH "GLACCT.IDX", TempAcctIdx(1), ElSize, NumIdxRecs

  OpenAcctFile AcctFileNum, NumAccts

  REDIM AcctList(1 TO NumIdxRecs) AS FLen       'Dim array to hold trans list
  FOR Cnt = 1 TO NumIdxRecs     'Read the file
    IF LEFT$(TempAcctIdx(Cnt).AcctNum, Lookuplen) = Lookup$ THEN
      GET AcctFileNum, TempAcctIdx(Cnt).RecNum, acct
      MatchAccts = MatchAccts + 1
      LSET AcctList(MatchAccts).V = List$
      AcctList(MatchAccts).V = TempAcctIdx(Cnt).AcctNum
      MID$(AcctList(MatchAccts).V, TotLen + 4) = acct.Title
      MID$(AcctList(MatchAccts).V, 41) = MKI$(TempAcctIdx(Cnt).RecNum)
    END IF
  NEXT

  CLOSE AcctFileNum

  SELECT CASE MatchAccts
  CASE 0
    PLAY "L32O3EC"              'Acct NOT found
    OK = MsgBox("GL.QSL", "NOACCT")
    EDCheck4Acct = False
  CASE 1
    AcctRecNum = CVI(RIGHT$(AcctList(1).V, 2))
    GOSUB GetAcctInfo
    EDCheck4Acct = True
  CASE ELSE
    Row = 5
    Col = 15
    LOCATE Row, Col
    Action = 0
    MaxLen = 39
    BoxBot = 18
    ListDescTxt$ = "Account No.   Description"
    ListDesc$ = SPACE$(MaxLen + 4)
    MID$(ListDesc$, 3) = ListDescTxt$
    QPrintRC ListDesc$, Row - 1, Col, 112       'Cnf.ActivCh
    DO          'call vertmenu to display matching recs to choose from
      VertMenuT AcctList(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
      IF Ky$ = CHR$(27) THEN
        EDCheck4Acct = False
        EXIT DO
      END IF
      IF Choice > 0 THEN
        AcctRecNum = CVI(RIGHT$(AcctList(Choice).V, 2))
        GOSUB GetAcctInfo
        EDCheck4Acct = True
        EXIT DO
      END IF
    LOOP
  END SELECT

  MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen

  ERASE AcctList, TempAcctIdx, ScrBuf


  EXIT FUNCTION

GetAcctInfo:
  OpenAcctFile AcctFileNum, NumAccts
  GET AcctFileNum, AcctRecNum, acct
  CLOSE AcctFileNum
  LSET Form$(AcctFld, 0) = acct.Num
  LSET Form$(AcctFld + 1, 0) = acct.Title
  LSET Form$(AcctFld + 2, 0) = MKI$(AcctRecNum)
RETURN



END FUNCTION

SUB EdEntry (CJType)

  IF CDActive$ <> "Y" THEN
    REDIM BankList$(1)
    GetBankList BankList$(), NumBanks
    IF NumBanks > 0 THEN
      MaxChoice = NumBanks
    ELSE
      OK = MsgBox("GL", "NOBANKS")
      EXIT SUB
    END IF
  ELSE
    MaxChoice = 1
  END IF

  GetPostDates LPDate, HPDate
  
  SHARED Choice$()
  REDIM Choice$(0 TO MaxChoice, 1)
  Choice$(0, 0) = "7"
  IF NumBanks > 0 THEN
    FOR BCnt = 1 TO NumBanks
      Choice$(BCnt, 0) = BankList$(BCnt)
    NEXT
  ELSE
    Choice$(1, 0) = "99 System"
  END IF
  

  NumOfScrns = 3
  REDIM Frm(1)    AS FormInfo
  REDIM FormName$(1 TO NumOfScrns)
  CommaFormat$ = "#########,.##"

  SELECT CASE CJType
  CASE 1        'Cash Reciepts
    FOR Scr = 1 TO NumOfScrns
      FormName$(Scr) = "CR" + LTRIM$(STR$(Scr))
    NEXT
    Header = 7
    TotalAmtFld = 5
  CASE 2        'Cash Disbursements
    FOR Scr = 1 TO NumOfScrns
      FormName$(Scr) = "CD" + LTRIM$(STR$(Scr))
    NEXT
    Header = 7
    TotalAmtFld = 5
    'CASE 3  'AP
    'AP stuff here
  END SELECT

  '-- Get total number of fields
  NumFlds = -1
  FOR Scr = 1 TO NumOfScrns
    NumFlds = NumFlds + LibNumberOfFields("GL", FormName$(Scr)) + 1
  NEXT

  REDIM Form$(NumFlds, 2)
  REDIM Fld(NumFlds) AS FieldInfo
  REDIM PageInfo(1 TO NumOfScrns) AS PageInfoType

  'for each page, get first and last fields
  StartEl = 0   'Load first form at array start
  FOR Scr = 1 TO NumOfScrns     'Get field defs
    LibGetFldDef "GL", FormName$(Scr), StartEl, Fld(), Form$(), ErrCode
    PageInfo(Scr).PageNo = Scr
    PageInfo(Scr).FirstFld = StartEl + 1
    StartEl = StartEl + Fld(StartEl).Fields + 1 'Inc StartEl to next
    PageInfo(Scr).LastFld = StartEl - 1
  NEXT

  FOR F = 1 TO NumFlds          'For all fields on all pages
    LSET Form$(F, 0) = ""       'Clear each field
  NEXT

  ' This is broken in my formedit
  REDIM AcctFlds(1 TO 36)
  FOR F = 1 TO NumFlds
    IF QPTrim$(Fld(F).ScratchS) = "A" THEN
      Cnt = Cnt + 1
      AcctFlds(Cnt) = F
    END IF
  NEXT
  Frm(1).PageDisable = True

  
  '--get first active record number
  OpenCJEditFile CJEditFile, NumEdTrans, CJType
  CurrRec = 0: NextRec = 1
  RecNum = GetNextRec(CJEditFile, NumEdTrans, CurrRec, NextRec)
  CLOSE CJEditFile
  IF RecNum = 0 THEN
    EdMode = 1
    'Record$ = "New "
  ELSE
    EdMode = 0
  END IF
  '===
  
  Scr = 1       'Set screen number to one
  LibFile2Scrn "GL", FormName$(Scr), MonoCode, -1, ErrCode      'display screen

  GOSUB ShowRecNo

  IF EdMode = 0 THEN
    Rec2Form Form$(), RecNum, Fld(), CJType
  END IF
  
  ShowEdMode EdMode
  
  Action = 1
  
  DO

    'IF Action = 1 THEN
    '  Scr = 1
    'END IF

    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    'This chunk of code deletes a distribution line when Alt+F3 is pressed
    'This needs to occur before prev fld test
    'Gets a list of acct fields from disk file.

    SELECT CASE Frm(1).KeyCode
    CASE -106   'Alt+F3
      
      IF Frm(1).FldNo > Header THEN             'if we're on an account field
        
        'for each accounting distribution
        FOR Cnt = 1 TO 36
          'if this acct field is less than the current field
          'set the row counter
          IF AcctFlds(Cnt) <= Frm(1).FldNo THEN
            FldRow = FldRow + 1 '+AcctFlds(Cnt)
          ELSE
            FOR Cnt3 = AcctFlds(FldRow) TO AcctFlds(FldRow) + 3
              LSET Form$(Cnt3, 0) = ""
            NEXT

            FOR Cnt3 = FldRow TO 35
              LSET Form$(AcctFlds(Cnt3), 0) = Form$(AcctFlds(Cnt3 + 1), 0)
              LSET Form$(AcctFlds(Cnt3) + 1, 0) = Form$((AcctFlds(Cnt3 + 1) + 1), 0)
              LSET Form$(AcctFlds(Cnt3) + 2, 0) = Form$((AcctFlds(Cnt3 + 1) + 2), 0)
              LSET Form$(AcctFlds(Cnt3) + 3, 0) = Form$((AcctFlds(Cnt3 + 1) + 3), 0)
            NEXT
            'Cnt is at 36 after 35 times thru above loop

            LSET Form$(AcctFlds(Cnt3), 0) = ""
            LSET Form$(AcctFlds(Cnt3) + 1, 0) = ""
            LSET Form$(AcctFlds(Cnt3) + 2, 0) = ""
            LSET Form$(AcctFlds(Cnt3) + 3, 0) = ""

            Frm(1).FldNo = AcctFlds(FldRow)
            Frm(1).PrevFld = Frm(1).FldNo - 1
            FldRow = 0
            Action = 1
            EXIT FOR
          END IF
        NEXT
      END IF
    CASE F2Key  'Change Editing Mode
      
      GOSUB CheckFormEdit
      IF CanChange THEN
        GOSUB InitForm
        SELECT CASE EdMode
        CASE 0  'were in Edit Mode
          EdMode = 1            ' switch to add mode
          ShowEdMode EdMode
          ResetForm Form$(), NumFlds
          Frm(1).FldNo = 1
          Action = 1
          QPrintRC "New  ", 25, 68, -1
          GOSUB ShowDistTot
          
        CASE 1  'currently adding records
          'check to see if there are any active transactions
          OpenCJEditFile CJEditFile, NumEdTrans, CJType
          CurrRec = 0: NextRec = 1
          RecNum = GetNextRec(CJEditFile, NumEdTrans, CurrRec, NextRec)
          CLOSE CJEditFile
          
          IF RecNum = 0 THEN    'No active records can't edit
            EdMode = 1
          ELSE
            EdMode = 0          'switch to edit mode
            ShowEdMode EdMode
            RecNum = ListRecs(CJType)
            IF RecNum > 0 THEN
              Rec2Form Form$(), RecNum, Fld(), CJType
              Frm(1).FldNo = 1
              Action = 1
            ELSE
              CLOSE
              EXIT SUB
            END IF
          END IF
          GOSUB ShowDistTot
        END SELECT              'Edmode
      END IF
    CASE ELSE
      
    END SELECT
    
    IF Frm(1).PrevFld <> Frm(1).FldNo THEN      'User is moving around the form
      
      GOSUB ShowDistTot         'Update Distribution totals
      
      '--Check the field we're in to see if it is an acct field.
      '  if so and if there is data in the field, remove dashes for user
      '  set stripped acct string in form
      
      IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN         '= AcctField
        LSET Form$(Frm(1).FldNo, 0) = StripAcct$(Form$(Frm(1).FldNo, 0))
        Action = 1
      END IF
      
      '--the following checks the field we have just left to see
      '  if it is an account field. if it is then
      '  if the field is not blank, the acct number is formatted with dashes
      '  if the field is blank the acct name and distribution amt are blanked
      '  and the amt field is protected to prevent user from entering a
      '  distribution without an account.
      
      IF QPTrim$(Fld(Frm(1).PrevFld).ScratchS) = "A" THEN       '"A" = AcctField
        IF LEN(QPTrim$(Form$(Frm(1).PrevFld, 0))) THEN
          TempAcct$ = FmtAcct$(RTRIM$(Form$(Frm(1).PrevFld, 0)), FundLen%, AcctLen%, DetLen%)
          LSET Form$(Frm(1).PrevFld, 0) = TempAcct$
          Fld(Frm(1).PrevFld + 3).Protected = False
        ELSE
          FOR Cnt2 = 1 TO 3
            LSET Form$(Frm(1).PrevFld + Cnt2, 0) = ""
            Fld(Frm(1).PrevFld + Cnt2).Protected = True
          NEXT
          IF Frm(1).FldNo > Frm(1).PrevFld THEN
            Frm(1).FldNo = Frm(1).PrevFld
          END IF
        END IF
        Action = 1
        
      END IF
      
      '--if we've just left an account field, check to see if
      '  we have a valid acct number
      
      IF LEN(QPTrim$(Form$(Frm(1).PrevFld, 0))) THEN
        IF QPTrim$(Fld(Frm(1).PrevFld).ScratchS) = "A" THEN     '= AcctField
          GOSUB CheckForValidAcctNum
        END IF
      END IF
      
    END IF      'End of test for PrevFld <> current FldNo
    
    IF SaveFlag THEN GOSUB SaveRec
    
    '--Process key strokes
    
    SELECT CASE Frm(1).KeyCode
    CASE F3Key
      IF EMode = 0 THEN
        OK = MsgBox%("GL.QSL", "OK2DELTR")
        IF OK = 2 THEN          'Pressed Button #2
          Form$(1, 0) = "-1"
          SaveField 1, Form$(), Fld(), BadFld
          SaveTrans Form$(), RecNum, EdMode, CJType
          ResetForm Form$(), NumFlds
          RecNum = ListRecs(CJType)
          IF RecNum > 0 THEN
            Rec2Form Form$(), RecNum, Fld(), CJType
          ELSE
            EXIT SUB
          END IF
        ELSE
          Action = 2
        END IF
        Frm(1).FldNo = 1
      END IF
      
    CASE F5KEY
      GOSUB CheckFormEdit
      IF CanChange THEN
        GOSUB InitForm
        IF EdMode = 0 THEN
          RecNum = ListRecs(CJType)
          IF RecNum > 0 THEN
            Rec2Form Form$(), RecNum, Fld(), CJType
            Frm(1).FldNo = 1
            Action = 1
          END IF
        END IF
        
      END IF
      
    CASE F7KEY
      IF QPTrim$(Fld(Frm(1).PrevFld).ScratchS) = "A" THEN       '= AcctField
        BEEP
      END IF
      
    CASE F10Key
      '--Make sure acct fields are formatted correctly if user
      '  Preses F10 with cursor in an account field
      IF QPTrim$(Fld(Frm(1).FldNo).ScratchS) = "A" THEN         '= AcctField
        IF LEN(QPTrim$(Form$(Frm(1).FldNo, 0))) > 0 THEN
          TempAcct$ = FmtAcct$(RTRIM$(Form$(Frm(1).FldNo, 0)), FundLen%, AcctLen%, DetLen%)
          LSET Form$(Frm(1).FldNo, 0) = TempAcct$
          'Fld(Frm(1).PrevFld + 3).Protected = False
          Frm(1).FldNo = 1
          'Action = 1
        END IF
      END IF
      'send user back thru loop to verify acct before saving
      Action = 1
      SaveFlag = True
      
    CASE CtrlL  'Ctrl+LeftArrow               'Get the next record
      
      GOSUB CheckFormEdit
      IF CanChange THEN
        IF RecNum > 0 THEN
          Frm(1).FldNo = 1
          OpenCJEditFile CJEditFile, NumEdTrans, CJType
          CurrRec = RecNum: NextRec = -1
          RecNum = GetNextRec(CJEditFile, NumEdTrans, CurrRec, NextRec)
          CLOSE CJEditFile
          GOSUB InitForm
          Rec2Form Form$(), RecNum, Fld(), CJType
          GOSUB ShowDistTot
          ShowEdMode EdMode
          'Frm(1).PrevFld = 1
          Scr = 1
          Action = 1
        END IF
      END IF
      
    CASE CtrlR  'Ctrl+RightArrow             'Get the previous record

      GOSUB CheckFormEdit
      IF CanChange THEN
        IF RecNum > 0 THEN
          Frm(1).FldNo = 1
          OpenCJEditFile CJEditFile, NumEdTrans, CJType
          CurrRec = RecNum: NextRec = 1
          RecNum = GetNextRec(CJEditFile, NumEdTrans, CurrRec, NextRec)
          CLOSE CJEditFile
          GOSUB InitForm
          Rec2Form Form$(), RecNum, Fld(), CJType
          GOSUB ShowDistTot
          ShowEdMode EdMode
         ' Frm(1).PrevFld = 1
          Action = 1
          Scr = 1
        END IF
      END IF
      
    END SELECT
    'End of Checking key codes
    
    '--Set screen pages
    'IF Frm(1).FldNo > PageInfo(Scr).LastFld OR Frm(1).FldNo < PageInfo(Scr).FirstFld AND NOT StayOnPage THEN
    IF Frm(1).FldNo > PageInfo(Scr).LastFld OR Frm(1).FldNo < PageInfo(Scr).FirstFld THEN

      SELECT CASE Frm(1).FldNo
      CASE PageInfo(1).FirstFld TO PageInfo(1).LastFld
        Scr = 1
      CASE PageInfo(2).FirstFld TO PageInfo(2).LastFld
        Scr = 2
      CASE PageInfo(3).FirstFld TO PageInfo(3).LastFld
        Scr = 3
      END SELECT
    '  Action = 1
      LibFile2Scrn "GL", FormName$(Scr), MonoCode, -2, ErrCode
      
      GOSUB ShowDistTot
      ShowEdMode EdMode
      GOSUB ShowRecNo
    END IF
    
    '--Buttons are not attached to the form, the following processes
    '--the rodent
    IF Frm(1).Presses THEN
      
      SELECT CASE Frm(1).MRow
      CASE 8
        SELECT CASE Frm(1).MCol
        CASE 71 TO 73
          'PressButton2 -72, 8, 72, 72
          PressButton -72, 8, 72, 72
        END SELECT
        
      CASE 19
        SELECT CASE Frm(1).MCol
        CASE 71 TO 73
          'PressButton2 -80, 19, 72, 72
          PressButton -80, 19, 72, 72
        END SELECT
        
      CASE 21
        SELECT CASE Frm(1).MCol
        CASE 9 TO 21
          'PressButton2 F10Key, 21, 9, 21
          PressButton F10Key, 21, 9, 21
        CASE 22 TO 37
          'PressButton2 F3Key, 21, 22, 37        'Alt+F3
          PressButton F3Key, 21, 22, 37        'Alt+F3
        CASE 38 TO 56
          'PressButton2 -106, 21, 38, 56
          PressButton -106, 21, 38, 56
        CASE 57 TO 71
          'PressButton2 ESC, 21, 57, 71
          PressButton ESC, 21, 57, 71
        END SELECT
        
      CASE 25
        SELECT CASE Frm(1).MCol
        CASE 27 TO 39
          'PressButton2 F2Key, 25, 27, 39
          PressButton F2Key, 25, 27, 39
        CASE 40 TO 52
          'PressButton2 F5KEY, 25, 40, 52
          PressButton F5KEY, 25, 40, 52
        CASE 53 TO 59
          'PressButton2 CtrlL, 25, 53, 59
          PressButton CtrlL, 25, 53, 59
        CASE 73 TO 79
          'PressButton2 CtrlR, 25, 73, 79
          PressButton CtrlR, 25, 73, 79
        END SELECT
      END SELECT
    END IF
    
  LOOP UNTIL Frm(1).KeyCode = ESC
  
  'End of FormEdit Loop
  EXIT SUB
  
  
CheckForValidAcctNum:
'  AcctFld = Frm(1).PrevFld
'  TempAcct$ = QPTrim$(Form$(AcctFld, 0))
'  IF LEN(TempAcct$) THEN
'    TempAcct$ = Form$(AcctFld, 0)
'    RecordNum = FindAcct(TempAcct$)
'    IF RecordNum > 0 THEN       'and RecordNum <> PreviousRecNum THEN         'Eureka!
'      Form$(AcctFld + 1, 0) = GetAcctTitle$(RecordNum)
'      Action = 1
'    ELSEIF RecordNum = 0 THEN
'      PLAY "L32O3EC"            'Acct NOT found
'      OK = MsgBox("GL.QSL", "NOACCT")
'      Frm(1).FldNo = Frm(1).PrevFld             'Put user back on acct field
'      LSET Form$(Frm(1).FldNo, 0) = ""
'      LSET Form$(Frm(1).FldNo + 1, 0) = ""
'      LSET Form$(Frm(1).FldNo + 3, 0) = ""
'      Action = 1
'    END IF
'  END IF
  
'  RETURN

  AcctFld = Frm(1).PrevFld
  SELECT CASE EDCheck4Acct%(Form$(), AcctFld%)
  CASE True
    SaveField AcctFld, Form$(), Fld(), BadFld   'Write the flag to the form$
    SaveField AcctFld + 1, Form$(), Fld(), BadFld               'Write the flag to the form$
    Action = 2
  CASE ELSE
    Frm(1).FldNo = AcctFld      'Put user back on acct field
    LSET Form$(Frm(1).FldNo, 0) = ""
    LSET Form$(Frm(1).FldNo + 1, 0) = ""
    Action = 2
  END SELECT

  RETURN
  
  
ShowDistTot:
  TotDist# = 0
  FOR Cnt = 1 TO NumFlds        'look thru acct fld list
    IF QPTrim$(Fld(Cnt).ScratchS) = "S" THEN    '= Distribution Field
      TotDist# = Round#(TotDist# + Value(Form$(Cnt, 0), ECode))
    END IF
  NEXT
  
  Amt# = Round#(Value(Form$(TotalAmtFld, 0), ECode))
  UnDist# = Amt# - TotDist#
  
  DIST$ = SPACE$(14)
  UnDist$ = SPACE$(14)
  LSET DIST$ = FUsing(STR$(TotDist#), CommaFormat$)
  LSET UnDist$ = FUsing(STR$(UnDist#), CommaFormat$)
  QPrintRC UnDist$, 20, 24, -1
  QPrintRC DIST$, 20, 56, -1
  
  RETURN
  
  
SaveRec:
  BadFld = 0
  Check4BlankFlds Form$(), BadFld, LPDate, HPDate, CJType              'make sure no blank fields
  IF BadFld = 0 THEN            'if no blanks, make sure
    GOSUB ShowDistTot           'distribution are in bal
    
    IF Amt# <> 0 AND UnDist# = 0 THEN
    'IF UnDist# = 0 THEN

      Dt$ = Form$(2, 0)
      Descript$ = Form$(3, 0)
     
      Form$(1, 0) = "0"         'Set deleted flag to false
      SaveField 1, Form$(), Fld(), BadFld       'Write the flag to the form$
      SaveTrans Form$(), RecNum, EdMode, CJType 'Save the transaction
      OK = MsgBox%("GL.QSL", "SAVED")           'tell user we saved it
      
      IF EdMode = 1 THEN        'if in Add Mode, clear
        ResetForm Form$(), NumFlds              ' the form for the next entry
        Form$(2, 0) = Dt$
        Form$(3, 0) = Descript$
        Frm(1).FldNo = 1
      END IF
      
    ELSE
      OK = MsgBox%("GL.QSL", "BADDIST")         'Distributions out of balance
      'tell user
    END IF
    
  ELSE
    SELECT CASE BadFld
    CASE 2
      OK = MsgBox%("GL.QSL", "BADDATE")           'Tell user no blank fields
    CASE ELSE
      OK = MsgBox%("GL.QSL", "NOBLANK")           'Tell user no blank fields
    END SELECT
    Frm(1).FldNo = BadFld

  END IF
  

  Action = 1
  SaveFlag = False
  
RETURN
  
  
CheckFormEdit:
  
  IF Frm(1).Edited THEN
    BEEP
    CanChange = False
    OK = MsgBox("CMNDLG", "CHANGREC")           'returns 2 for Abandon
    IF OK = 2 THEN CanChange = True
    Action = 2
  ELSE
    CanChange = True
  END IF
  
  RETURN
  
InitForm:
  REDIM Frm(1)    AS FormInfo
  Frm(1).PageDisable = True
  Action = 1
RETURN

ShowRecNo:
  IF RecNum > 0 THEN
    QPrintRC STR$(RecNum), 25, 67, -1
  ELSE
    QPrintRC "New  ", 25, 67, -1
  END IF

RETURN

END SUB

  '***************************************************************************
  'Finds the next undeleted record.
  'Call with NextRec value of -1 for previous record, +1 for the next record.
  'If a record is not found, the function returns the value of CurrRec.
  'File should be opened before function is called
  '***************************************************************************
FUNCTION GetNextRec (FileNum, NumRecs, CurrRec, NextRec)
  
  Found = 0
  Rec = CurrRec
  
  DO
    
    Rec = Rec + NextRec         'Set file pointer to next record
    
    IF Rec > NumRecs OR Rec <= 0 THEN           'test for beg or end of file
      Found = 0 'if no more records then get out
      EXIT DO
    END IF
    
    GET FileNum, Rec, CJEdit    'Get the record
    
    IF CJEdit.DELFLAG <> -1 THEN                'Ok if not deleted
      Found = 1
      EXIT DO   'Get out of loop when we find one
    END IF
    
  LOOP
  
  IF Found = 0 THEN
    GetNextRec = CurrRec
  ELSE
    GetNextRec = Rec
  END IF
  
END FUNCTION

FUNCTION ListRecs (CJType)
  
  SHARED Cnf AS Config
  REDIM ScrBuf(2000)
  
  MScrnSave 1, 1, 25, 80, ScrBuf(0)             'Save the existing screen
  
  QPrintRC "List ", 25, 68, -1
  
  OpenCJEditFile CJEditFile, NumEdTrans, CJType
  
  REDIM TransList$(1 TO NumEdTrans)             'Dim array to hold trans list
  
  Active = 0    'Counter for Active Transactions
  
  FOR Cnt = 1 TO NumEdTrans     'Read the file
    GET CJEditFile, Cnt, CJEdit
    IF NOT CJEdit.DELFLAG THEN  'if active add to list
      Active = Active + 1
      
      List$ = SPACE$(75)
      
      LSET List$ = ConvDateStr1(Num2Date(CJEdit.TRDATE))
      MID$(List$, 11) = CJEdit.DESC
      MID$(List$, 36) = STR$(Cnt)               'CJEdit.Ref
      MID$(List$, 53) = FUsing$(STR$(CJEdit.Amt), "#########,.##")
      MID$(List$, 73) = MKI$(Cnt)
      
      TransList$(Active) = List$
      
    END IF
    
  NEXT
  CLOSE
  
  IF Active = 0 THEN
    'KILL "CJED.DAT"                     '..init file to zero
    EXIT FUNCTION
  END IF
  
  REDIM PRESERVE TransList$(1 TO Active)        'Redim list to only active trans
  
  LOCATE 4, 6   'set up vertmenu
  Action = 0
  MaxLen = 65
  BoxBot = 22
  
  ListDesc$ = "   Date     Description               Ref                      Amt   "
  QPrintRC ListDesc$, 3, 6, 112 'Cnf.ActivCh
  'MPaintBox 6, 75, 7, 76, 8              'to add shadow to list
  
  DO            'call vertmenu
    VertMenu TransList$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    IF Ky$ = CHR$(27) THEN
      CLOSE
      Abort = True
      EXIT DO
      EXIT FUNCTION
    END IF
    IF Choice > 0 THEN EXIT DO
  LOOP
  
  IF Abort THEN
    Choice = 0
  ELSE
    RecNum = CVI(MID$(TransList$(Choice), 73, 2))               'Get the record number
  END IF        'off the end of the $
  
  CLOSE CJEditFile              'Clean up after ourselves
  ERASE TransList$
  MScrnRest 1, 1, 25, 80, ScrBuf(0)             'Restore Screen
  
  ListRecs = RecNum
  
  
END FUNCTION

SUB OpenCJEditFile (CJEditFile, NumEdTrans, CJType)
  
  SELECT CASE CJType
  CASE 1
    FileName$ = "GLCRED.DAT"
  CASE 2
    FileName$ = "GLCDED.DAT"
  END SELECT
  
  CJEdLen = LEN(CJEdit)
  CJEditFile = FREEFILE
  OPEN FileName$ FOR RANDOM ACCESS READ WRITE SHARED AS CJEditFile LEN = CJEdLen
  NumEdTrans = LOF(CJEditFile) \ CJEdLen
  
END SUB

SUB PostTrans (CJType)
  
  '--Paint a background
  FOR Cnt = 2 TO 24
    QPrintRC STRING$(80, CHR$(177)), Cnt, 1, 3
  NEXT
  
  '--Verify that there are active transactions to post
  OpenCJEditFile CJEditFile, NumEdTrans, CJType
  FOR Cnt = 1 TO NumEdTrans     'summarize the file totals
    GET CJEditFile, Cnt, CJEdit
    IF NOT CJEdit.DELFLAG THEN
      Active = Active + 1
    END IF
  NEXT
  CLOSE
  
  '--If there are no active trans, tell user and exit
  IF Active = 0 THEN
    '--No active transactions - tell user and get out
    OK = MsgBox%("GL.QSL", "NOTRANS")
    EXIT SUB
  END IF
  
  '--make sure we're ready to post
  OK = MsgBox%("GL.QSL", "POSTGJ")
  IF OK = 1 THEN EXIT SUB       '1=No
  
  '--Create a fund array to hold transaction totals by fund
  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  REDIM TrFundSum#(1 TO NumFunds)
  
  SELECT CASE CJType
  CASE 1
    CJEditFile$ = "GLCRED.DAT"
    TmpIFFile$ = "CJRPOST.DAT"
    MSrc$ = "CR" + ConvDateStr$(DATE$)
    'CshAcct$ = CRCashAcct$
  CASE 2
    CJEditFile$ = "GLCDED.DAT"
    TmpIFFile$ = "CJDPOST.DAT"
    MSrc$ = "CD" + ConvDateStr$(DATE$)
    'CshAcct$ = CashAcct$
  END SELECT
  
  '--For Central Depository - number of characters to pad detail code with
  IF CDActive$ = "Y" THEN
    PadChars = DetLen - FundLen
    IF PadChars > 0 THEN
      DetPad$ = STRING$(PadChars, "0")
    END IF
  END IF

  '--open the edit file
  OpenCJEditFile CJEditFile, NumEdTrans, CJType

  '--open the interface file
  DIM Tr2Post AS GLTransRecType
  RecLen = LEN(Tr2Post)
  IFFile = FREEFILE
  OPEN TmpIFFile$ FOR RANDOM AS IFFile LEN = RecLen
  RecordNum = LOF(IFFile) \ RecLen
  IF RecordNum > 0 THEN
    CLOSE
    CLS
    QPrintRC "WARNING", 9, 1, 12
    QPrintRC "Interface file already exists.  Possible problems with last posting.", 10, 1, 12
    QPrintRC "Contact Software support. Press any key to continue.", 11, 1, 12
    k$ = INPUT$(1)
    EXIT SUB
  END IF
  
  RecordNum = 0
  
  FOR Cnt = 1 TO NumEdTrans
    
    '--Reinitialize transaction fund summary elements for next transaction
    FOR Fund = 1 TO NumFunds
      TrFundSum#(Fund) = 0
    NEXT
    
    GET CJEditFile, Cnt, CJEdit
    
    IF NOT CJEdit.DELFLAG THEN

      '--Assign each distribution to the interface file
      FOR AcctDist = 1 TO 36

        '--Check for the last accounting distribution.
        IF LEN(QPTrim$(CJEdit.DIST(AcctDist).DACN)) = 0 THEN

          '--No more distributions so create the "Other" side of the
          '--transaction.
          FOR Fund = 1 TO NumFunds
            IF TrFundSum#(Fund) <> 0 THEN
              
              '--create the object code for the cent dep due to acct
              IF CDActive$ = "Y" THEN
                IF PadChars > 0 THEN
                  FundDue$ = FundList$(Fund) + DetPad$
                ELSE
                  FundDue$ = FundList$(Fund)
                END IF
              END IF
              
              RecordNum = RecordNum + 1
              Tr2Post.AcctNum = FundList$(Fund) + CshAcct$
              Tr2Post.TRDATE = CJEdit.TRDATE
              Tr2Post.DESC = CJEdit.DESC
              Tr2Post.Ref = CJEdit.DOCREF
              SELECT CASE CJType
                '--Cash Reciepts - cash is debited
                CASE 1
                  Tr2Post.DrAmt = TrFundSum#(Fund)
                  Tr2Post.CrAmt = 0
                  '-Update Central Depository
                  IF CDActive$ = "Y" THEN
                    CDDrAcct$ = CDCash$
                    CDCrAcct$ = CDDue$ + FundDue$
                  END IF
                '--Cash Disbursements - cash is credited
                CASE 2
                  Tr2Post.DrAmt = 0
                  Tr2Post.CrAmt = TrFundSum#(Fund)
                  '-Update Central Depository
                  IF CDActive$ = "Y" THEN
                    CDDrAcct$ = CDDue$ + FundDue$
                    CDCrAcct$ = CDCash$
                  END IF
              END SELECT
              Tr2Post.Src = MSrc$
              PUT IFFile, RecordNum, Tr2Post

              '--make entries to Central Depository Here.
              IF CDActive$ = "Y" THEN

                '--Debit to central depository
                RecordNum = RecordNum + 1
                Tr2Post.AcctNum = CDDrAcct$
                Tr2Post.TRDATE = CJEdit.TRDATE
                Tr2Post.DESC = CJEdit.DESC
                Tr2Post.Ref = CJEdit.DOCREF
                Tr2Post.DrAmt = TrFundSum#(Fund)
                Tr2Post.CrAmt = 0
                Tr2Post.Src = MSrc$
                PUT IFFile, RecordNum, Tr2Post

                '--Credit to central depository
                RecordNum = RecordNum + 1
                Tr2Post.AcctNum = CDCrAcct$
                Tr2Post.TRDATE = CJEdit.TRDATE
                Tr2Post.DESC = CJEdit.DESC
                Tr2Post.Ref = CJEdit.DOCREF
                Tr2Post.DrAmt = 0
                Tr2Post.CrAmt = TrFundSum#(Fund)
                Tr2Post.Src = MSrc$
                PUT IFFile, RecordNum, Tr2Post

              END IF

            END IF              'Fund summary <> 0
          NEXT  'fund
          
          EXIT FOR              'NumFunds
          
        ELSE
          '--Not the last distribution so add this distribution to the
          '--interface file.
          RecordNum = RecordNum + 1

          '--Assign the fields to the interface file
          Tr2Post.AcctNum = CJEdit.DIST(AcctDist).DACN
          Tr2Post.TRDATE = CJEdit.TRDATE
          Tr2Post.DESC = CJEdit.DESC
          Tr2Post.Ref = CJEdit.DOCREF
          
          '--Set Debit and credit amounts and source ref
          '--according to Journal Type
          SELECT CASE CJType
            '--Cash Reciepts - distributions are credits
            CASE 1
              Tr2Post.DrAmt = 0
              Tr2Post.CrAmt = CJEdit.DIST(AcctDist).DAMT
            '--Cash Disbursements - distributions are debits
            CASE 2
              Tr2Post.DrAmt = CJEdit.DIST(AcctDist).DAMT
              Tr2Post.CrAmt = 0
          END SELECT
          Tr2Post.Src = MSrc$
          PUT IFFile, RecordNum, Tr2Post
          
          '--Add this distribution to proper fund
          FOR Fund = 1 TO NumFunds
            FundNum$ = LEFT$(CJEdit.DIST(AcctDist).DACN, FundLen)
            IF FundNum$ = FundList$(Fund) THEN
              TrFundSum#(Fund) = TrFundSum#(Fund) + Round#(CJEdit.DIST(AcctDist).DAMT)
              'FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(CJEdit.DIST(AcctDist).DAMT)
              EXIT FOR
            END IF
          NEXT
          
        END IF  'test for blank distribution line
      NEXT    'Acct'g Distribution
    END IF  'not deleted
  NEXT    'transaction
  
  '--Post cash disbursements to o/s check file
  IF CJType = 2 THEN

    '--open o/s check file
    DIM OSChk AS OSChkRecType
    OSChkRecLen = LEN(OSChk)
    OSChkFile = FREEFILE
    OPEN "crchk.dat" FOR RANDOM AS OSChkFile LEN = OSChkRecLen
    NumOSChks = LOF(OSChkFile) \ OSChkRecLen

    '--assign each edit transaction to the o/s check file
    FOR Cnt = 1 TO NumEdTrans
      GET CJEditFile, Cnt, CJEdit

      '--Make sure trans is not deleted
      IF NOT CJEdit.DELFLAG THEN
         NumOSChks = NumOSChks + 1
         OSChk.ChkNum = VAL(CJEdit.DOCREF)
         OSChk.ChkDate = ConvDateStr1$(Num2Date(CJEdit.TRDATE))
         OSChk.DESC = CJEdit.DESC
         OSChk.Amt = CJEdit.Amt
         OSChk.Src = 0 'code as apcheck
         OSChk.Cleared = 0
         PUT OSChkFile, NumOSChks, OSChk
      END IF
    NEXT
  END IF

  CLOSE

  '--Post to G/L via common post & link sub in GLAUX.BAS
  Post2GL TmpIFFile$, BadTrans%
  IF BadTrans > 0 THEN
    CLS
    PRINT "Error: One or more transactions were not posted."
    PRINT "Press any key to view posting log"
    k$ = INPUT$(1)
    PrintRptFile "Posting Log", "GLUTIL.LOG", LPTNo, RetCode%, 2
  END IF
  
  '--kill the temp files
  KILL CJEditFile$
  KILL TmpIFFile$
  
  '--Posting Report
  'DIM Trans2Post AS GLTransRecType
  'RecLen = LEN(Trans2Post)
  '
  'IFFile = FREEFILE
  'OPEN TmpIFFile$ FOR RANDOM AS IFFile LEN = RecLen
  '
  'NumTrans = LOF(IFFile) \ RecLen
  '
  'PrnFile = FREEFILE
  'OPEN "Trans.prn" FOR OUTPUT AS #PrnFile
  '
  ''print field names
  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Date"
  'MID$(ToPrint$, 11) = "Acct Number"
  'MID$(ToPrint$, 29) = "Description"
  'MID$(ToPrint$, 46) = "Reference"
  'MID$(ToPrint$, 61) = "Debit"
  'MID$(ToPrint$, 74) = "Credit"
  'PRINT #PrnFile, ToPrint$
  '
  'FOR Cnt = 1 TO NumTrans
  '  GET IFFile, Cnt, Trans2Post
  '  ToPrint$ = SPACE$(80)
  '  LSET ToPrint$ = ConvDateStr1(Num2Date$(Trans2Post.TRDATE))
  '  MID$(ToPrint$, 11) = Trans2Post.AcctNum
  '  MID$(ToPrint$, 29) = LEFT$(Trans2Post.DESC, 15)
  '  MID$(ToPrint$, 46) = Trans2Post.Ref
  '  MID$(ToPrint$, 56) = FUsing(STR$(Trans2Post.DrAmt), "#########,.##")
  '  MID$(ToPrint$, 69) = FUsing(STR$(Trans2Post.CrAmt), "#########,.##")
  '  MID$(ToPrint$, 80) = STR$(Cnt&)
  ' MID$(ToPrint$, 90) = STR$(Trans2Post.NextTran)
  '  PRINT #PrnFile, ToPrint$
  'NEXT
  
  'CLOSE
  'Prn2Scrn "Trans.prn", 1, 25, 11, PrnFlag
  'KILL "trans.prn"

END SUB

SUB PostTrans2 (CJType)
  
  '--Paint a background
  FOR Cnt = 2 TO 24
    QPrintRC STRING$(80, CHR$(177)), Cnt, 1, 3
  NEXT

  '--Verify that there are active transactions to post
  OpenCJEditFile CJEditFile, NumEdTrans, CJType
  FOR Cnt = 1 TO NumEdTrans
    GET CJEditFile, Cnt, CJEdit
    IF NOT CJEdit.DELFLAG THEN
      Active = Active + 1
    END IF
  NEXT
  CLOSE

  '--If there are no active trans, tell user and exit
  IF Active = 0 THEN
    '--No active transactions - tell user and get out
    OK = MsgBox%("GL.QSL", "NOTRANS")
    EXIT SUB
  END IF

  '--make sure we're ready to post
  OK = MsgBox%("GL.QSL", "POSTGJ")
  IF OK = 1 THEN EXIT SUB       '1=No

  '--Create a fund list
  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds

  '--Holds a summary of the transaction's distributions by fund
  REDIM TrFundSum#(1 TO NumFunds)

  '--Set vars according to who called me to post
  SELECT CASE CJType
  CASE 1
    CJEditFile$ = "GLCRED.DAT"
    TmpIFFile$ = "CJRPOST.DAT"
    MSrc$ = "CR" + ConvDateStr$(DATE$)
    CJPrnFile$ = "CRTRANS.PRN"
  CASE 2
    CJEditFile$ = "GLCDED.DAT"
    TmpIFFile$ = "CJDPOST.DAT"
    MSrc$ = "CD" + ConvDateStr$(DATE$)
    CJPrnFile$ = "CDTRANS.PRN"
  END SELECT

  ToPrint$ = SPACE$(80)

  '--For Central Depository - number of characters to pad detail code with
  IF CDActive$ = "Y" THEN
    PadChars = DetLen - FundLen
    IF PadChars > 0 THEN
      DetPad$ = STRING$(PadChars, "0")
    END IF
  END IF

  '--open the edit file
  OpenCJEditFile CJEditFile, NumEdTrans, CJType

  '--open the interface file
  DIM Tr2Post AS GLTransRecType
  RecLen = LEN(Tr2Post)
  IFFile = FREEFILE
  OPEN TmpIFFile$ FOR RANDOM AS IFFile LEN = RecLen
  RecordNum = LOF(IFFile) \ RecLen
  IF RecordNum > 0 THEN
    CLOSE
    CLS
    QPrintRC "Error:", 9, 1, 12
    QPrintRC "Interface file already exists.  Possible problems with last cash journal posting.", 10, 1, 12
    QPrintRC "Contact Software support. Press any key to continue.", 11, 1, 12
    k$ = INPUT$(1)
    EXIT SUB
  END IF

  RecordNum = 0
  FOR Cnt = 1 TO NumEdTrans

    '--Reinitialize transaction fund summary elements for next transaction
    FOR Fund = 1 TO NumFunds
      TrFundSum#(Fund) = 0
    NEXT

    GET CJEditFile, Cnt, CJEdit

    IF NOT CJEdit.DELFLAG THEN

      '--if we're using a central depository post one entry to
      '--central depository cash per transaction
      IF CDActive$ = "Y" THEN

        CshAcct$ = CashAcct$
        RecordNum = RecordNum + 1

        Tr2Post.AcctNum = CDCash$
        Tr2Post.TRDATE = CJEdit.TRDATE
        Tr2Post.DESC = CJEdit.DESC
        Tr2Post.Ref = CJEdit.DOCREF
        SELECT CASE CJType
          '--Cash Reciepts - cash is debited
          CASE 1
            Tr2Post.DrAmt = Round(CJEdit.Amt)
            Tr2Post.CrAmt = 0
          '--Cash Disbursements - cash is credited
          CASE 2
            Tr2Post.DrAmt = 0
            Tr2Post.CrAmt = Round(CJEdit.Amt)
          CASE ELSE
        END SELECT
        Tr2Post.Src = MSrc$
        PUT IFFile, RecordNum, Tr2Post

      ELSE
        '--Find the cash account noumber for this trans from the bank code
        CshAcct$ = GetBankGLAcct$(VAL(CJEdit.RecCode))
      END IF
      
      '--Assign each distribution to the interface file
      FOR AcctDist = 1 TO 36

        '--Skip blank account numbers.
        IF LEN(QPTrim$(CJEdit.DIST(AcctDist).DACN)) > 0 THEN

          RecordNum = RecordNum + 1

          '--Assign the fields to the interface file
          Tr2Post.AcctNum = CJEdit.DIST(AcctDist).DACN
          Tr2Post.TRDATE = CJEdit.TRDATE
          Tr2Post.DESC = CJEdit.DESC
          Tr2Post.Ref = CJEdit.DOCREF
          '--Set Dr and Cr amts and source ref according to Journal Type
          SELECT CASE CJType
            '--Cash Reciepts - distributions are credits
            CASE 1
              Tr2Post.DrAmt = 0
              Tr2Post.CrAmt = CJEdit.DIST(AcctDist).DAMT
            '--Cash Disbursements - distributions are debits
            CASE 2
              Tr2Post.DrAmt = CJEdit.DIST(AcctDist).DAMT
              Tr2Post.CrAmt = 0
          END SELECT
          Tr2Post.Src = MSrc$
          PUT IFFile, RecordNum, Tr2Post

          '--Add this distribution to proper fund
          FOR Fund = 1 TO NumFunds
            FundNum$ = LEFT$(CJEdit.DIST(AcctDist).DACN, FundLen)
            IF FundNum$ = FundList$(Fund) THEN
              TrFundSum#(Fund) = TrFundSum#(Fund) + Round#(CJEdit.DIST(AcctDist).DAMT)
              'FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(CJEdit.DIST(AcctDist).DAMT)
              EXIT FOR
            END IF
          NEXT 'Fund

        END IF 'test for blank distribution line
      NEXT 'Distrubution

      '--No more distributions so create the "Opposite" entries.
      ' -One entry to cash or central depository for each fund
      ' -in this transaction
      FOR Fund = 1 TO NumFunds
        IF TrFundSum#(Fund) <> 0 THEN

          '--if we're using a central depository then create the
          ' -object code for the central depository "due to" acct
          IF CDActive$ = "Y" THEN
            IF PadChars > 0 THEN
              FundDue$ = FundList$(Fund) + DetPad$
            ELSE
              FundDue$ = FundList$(Fund)
            END IF
          END IF

          A$ = FundList$(Fund) + CshAcct$
          B = FindAcct(A$)
          IF OK = False THEN
            BadCashAcct = True
          END IF
          
          '--Fund's Cash or Central Depository entry
          RecordNum = RecordNum + 1
          Tr2Post.AcctNum = A$  'FundList$(Fund) + CshAcct$
          Tr2Post.TRDATE = CJEdit.TRDATE
          Tr2Post.DESC = CJEdit.DESC
          Tr2Post.Ref = CJEdit.DOCREF
          SELECT CASE CJType
            CASE 1   'Cash Reciepts - cash is debited
              Tr2Post.DrAmt = TrFundSum#(Fund)
              Tr2Post.CrAmt = 0
            CASE 2    'Cash Disb - cash is credited
              Tr2Post.DrAmt = 0
              Tr2Post.CrAmt = TrFundSum#(Fund)
          END SELECT
          Tr2Post.Src = MSrc$
          PUT IFFile, RecordNum, Tr2Post

          '--Entry to the Central Dep to update liability to this fund
          IF CDActive$ = "Y" THEN
            RecordNum = RecordNum + 1
            Tr2Post.AcctNum = CDDue$ + FundDue$
            Tr2Post.TRDATE = CJEdit.TRDATE
            Tr2Post.DESC = CJEdit.DESC
            Tr2Post.Ref = CJEdit.DOCREF
            Tr2Post.Src = MSrc$
            SELECT CASE CJType
              CASE 1   'Cash Reciepts - credit due to fund
                Tr2Post.DrAmt = 0
                Tr2Post.CrAmt = TrFundSum#(Fund)
              CASE 2    'Cash Disb - debit due to fund
                Tr2Post.DrAmt = TrFundSum#(Fund)
                Tr2Post.CrAmt = 0
            END SELECT
            PUT IFFile, RecordNum, Tr2Post
          END IF

        END IF 'fund summary test
      NEXT  'next tr fund summary
    END IF 'trx not deleted
  NEXT  'next trx

  IF BadCashAcct THEN
    CLOSE
    KILL TmpIFFile$
    OK = MsgBox("GL", "BADCASH")
    EXIT SUB
  END IF


  '--Post cash disbursements to o/s check file
  IF CJType = 2 THEN

    '--open o/s check file
    DIM OSChk AS OSChkRecType
    OSChkRecLen = LEN(OSChk)
    OSChkFile = FREEFILE
    OPEN "crchk.dat" FOR RANDOM AS OSChkFile LEN = OSChkRecLen
    NumOSChks = LOF(OSChkFile) \ OSChkRecLen

    '--assign each edit transaction to the o/s check file
    FOR T = 1 TO NumEdTrans
      GET CJEditFile, T, CJEdit

      '--Make sure trans is not deleted
      IF NOT CJEdit.DELFLAG THEN
         NumOSChks = NumOSChks + 1
         OSChk.ChkNum = VAL(CJEdit.DOCREF)
         OSChk.ChkDate = ConvDateStr1$(Num2Date(CJEdit.TRDATE))
         OSChk.DESC = CJEdit.DESC
         OSChk.Amt = CJEdit.Amt
         OSChk.Src = 0 'code as apcheck
         OSChk.Cleared = 0
         PUT OSChkFile, NumOSChks, OSChk
      END IF
    NEXT
  END IF

  CLOSE

  GOSUB PrintFile

  '---unREM me to stop before actual posting. this is a test this is only..
  'KILL TmpIFFile$
  'END
  '===============================
  

  '--Post to G/L via common post & link sub in GLAUX.BAS
  Post2GL TmpIFFile$, BadTrans%
  IF BadTrans > 0 THEN
    CLS
    PRINT "Error: One or more transactions were not posted."
    PRINT "Press any key to view posting log"
    k$ = INPUT$(1)
    PrintRptFile "Posting Log", "GLUTIL.LOG", LPTNo, RetCode%, 2
  END IF

  '--kill the temp files
  KILL CJEditFile$
  KILL TmpIFFile$

EXIT SUB


PrintFile:
'--Posting Journal Entry Report
  DIM Trans2Post AS GLTransRecType
  RecLen = LEN(Trans2Post)
  IFFile = FREEFILE
  OPEN TmpIFFile$ FOR RANDOM AS IFFile LEN = RecLen
  NumTrans = LOF(IFFile) \ RecLen
  
  REDIM SortTrans(NumTrans) AS GLTransRecType
  FOR Cnt = 1 TO NumTrans
    GET IFFile, Cnt, Trans2Post
    SortTrans(Cnt) = Trans2Post
  NEXT

  CLOSE

  SortT SortTrans(1), NumTrans, 0, 96, 2, 14

  IFFile = FREEFILE
  OPEN TmpIFFile$ FOR RANDOM AS IFFile LEN = RecLen

  FOR Cnt = 1 TO NumTrans
    PUT IFFile, Cnt, SortTrans(Cnt)
  NEXT
  
  PrnFile = FREEFILE
  PrnFileName$ = "CJPOST.PRN"
  OPEN PrnFileName$ FOR OUTPUT AS #PrnFile
  
  GOSUB CJJEHeader
  
  FOR Cnt = 1 TO NumTrans
    GET IFFile, Cnt, Trans2Post

    JEDebits# = JEDebits# + Round(Trans2Post.DrAmt)
    JECredits# = JECredits# + Round(Trans2Post.CrAmt)

    LSET ToPrint$ = ""
    LSET ToPrint$ = ConvDateStr1(Num2Date$(Trans2Post.TRDATE))
    MID$(ToPrint$, 11) = Trans2Post.AcctNum
    MID$(ToPrint$, 29) = LEFT$(Trans2Post.DESC, 15)
    MID$(ToPrint$, 46) = Trans2Post.Ref
    MID$(ToPrint$, 56) = FUsing(STR$(Trans2Post.DrAmt), "#########,.##")
    MID$(ToPrint$, 69) = FUsing(STR$(Trans2Post.CrAmt), "#########,.##")
    PRINT #PrnFile, ToPrint$
    LineCnt = LineCnt + 1
    IF LineCnt > 55 THEN
      PRINT #PrnFile, CHR$(12)
      GOSUB CJJEHeader
    END IF
  NEXT

  PRINT #PrnFile,
  PRINT #PrnFile, "Posting Totals:";
  PRINT #PrnFile, TAB(56); FUsing(STR$(JEDebits#), "############,.##");
  PRINT #PrnFile, TAB(69); FUsing(STR$(JECredits#), "############,.##")
  PRINT #PrnFile, CHR$(12)

  CLOSE
  PrintRptFile "Cash Journal Entries", PrnFileName$, LPTPort, RetCode%, EntryPoint
  RETURN



CJJEHeader:
  PRINT #PrnFile, "Cash Journal Entries"
  PRINT #PrnFile, "Module: " + MSrc$

  PRINT #PrnFile,

  LSET ToPrint$ = ""
  MID$(ToPrint$, 1) = "Date"
  MID$(ToPrint$, 11) = "Acct No"
  MID$(ToPrint$, 29) = "Description"
  MID$(ToPrint$, 46) = "Reference"
  MID$(ToPrint$, 57) = "     Debit"
  MID$(ToPrint$, 69) = "    Credit"
  PRINT #PrnFile, ToPrint$

  PRINT #PrnFile, STRING$(78, "=")
  LineCnt = 5
  RETURN


END SUB

SUB PrnEditList (CJType)
  
  REDIM Title$(1 TO 4)
  SHARED Choice$()              'Multiple choice array
  SHARED CashAcct$, User$, CDCash$, CDDue$, CDActive$, CRCashAcct$

  WarningFlag = False           'Check for valid cash account
  
  REDIM Form$(0, 0)             'Holds data from form
  REDIM Fld(0) AS FieldInfo     'Field editing array
  
  REDIM Choice$(0 TO 4, 1)      'Choices for entry type field
  Choice$(0, 0) = "2"           'Field Number
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn #1"
  Choice$(3, 0) = "Prn #2"
  
  '--call sub in GLAUX module to load screen & form
  LibLoadDisplayForm "GL.QSL", "PRNGJ", Form$(), Fld()
  PrintHelp "Print Register"
  
  DIM Frm AS FormInfo
  Action = 1
  
  '--Get the report options from user
  DO
    
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    '--Check the user's key presses
    SELECT CASE Frm.KeyCode
      
    CASE F10Key 'SaveButton
      IF LEN(LTRIM$(RTRIM$(Form$(2, 0)))) = 0 THEN              'Fld#2 = OutputDevice
        Dev$ = "S"
        LPTNo = 1
      ELSE
        Dev$ = LEFT$(Form$(2, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(2, 0)), 1))
      END IF
      
    CASE ESC
      EXIT SUB
      
    END SELECT
    
  LOOP UNTIL Frm.KeyCode = F10Key
  
  '==Start processing register=>
  LOCATE 1, 1, 0

  TranCol = 58
  CommaFmt$ = "#########,.##"   'ten millions
  DistSumLine$ = "-------------"
  TransTotal# = 0
  TranCnt = 0
  FF$ = CHR$(12)
  MaxLines = 55

  OpenCJEditFile CJEditFile, NumEdTrans, CJType
  
  SELECT CASE CJType
  CASE 1
    FileName$ = "CRREG.PRN"
    RegTitle$ = "Cash Receipts Journal"
  CASE 2
    FileName$ = "CDREG.PRN"
    RegTitle$ = "Cash Disbursements Journal"
  END SELECT
  
  IF CDActive$ = "Y" THEN
    CshAcct$ = CashAcct$
  END IF
  
  PrnFileNum = FREEFILE
  OPEN FileName$ FOR OUTPUT AS #PrnFileNum
  
  '--Get a list of active funds
  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  REDIM FundSum#(1 TO NumFunds)
  REDIM FundGrdTot#(1 TO NumFunds)
  
  GOSUB PrnCJPageHeader

  FOR transaction = 1 TO NumEdTrans
    GET CJEditFile, transaction, CJEdit
    
    IF NOT CJEdit.DELFLAG THEN
      TranCnt = TranCnt + 1
      TransTotal# = TransTotal# + Round#(CJEdit.Amt)
      
      '--If we're NOT using a central depository fund then use
      '--the bank gl account number
      'IF CDActive$ <> "Y" THEN
      '  BankNum = VAL(CJEdit.RECCODE)
      '  IF BankNum > 0 THEN
      '    CshAcct$ = GetBankGLAcct$(BankNum)
      '  END IF
      'END IF

      '--Print 1st Line - Transaction details
      ToPrint$ = SPACE$(80)
      LSET ToPrint$ = ConvDateStr1$(Num2Date(CJEdit.TRDATE))
      MID$(ToPrint$, 10) = CJEdit.DESC
      MID$(ToPrint$, 32) = CJEdit.DOCREF
      MID$(ToPrint$, 44) = CJEdit.RecCode
      MID$(ToPrint$, 68) = FUsing$(STR$(CJEdit.Amt), CommaFmt$)
      PRINT #PrnFileNum, ToPrint$
      LineCnt = LineCnt + 1
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrnCJPageHeader
      END IF

      '--Blank line between detail and acct'g distributions
      PRINT #PrnFileNum,
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrnCJPageHeader
      END IF
      
      '--Print Distribution Label
      ToPrint$ = SPACE$(80)
      MID$(ToPrint$, 2) = "Accounting Distribution:"
      PRINT #PrnFileNum, ToPrint$
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrnCJPageHeader
      END IF
      
      '--Print Field Titles
      ToPrint$ = SPACE$(80)
      MID$(ToPrint$, 4) = "Account Number   Name                                       Amount"
      PRINT #PrnFileNum, ToPrint$
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrnCJPageHeader
      END IF
      
      '--Print Accounting Distributions
      TotTranDist# = 0
      
      '--Loop Thru distributions to print and summarize
      FOR AcctDist = 1 TO 36
        
        '--no more distributions when we find a blank Acct Number field
        IF LEN(QPTrim$(CJEdit.DIST(AcctDist).DACN)) = 0 THEN
          
          TranCashTot# = 0
          FOR Fund = 1 TO NumFunds
            '--Test each fund to see if it has any distributions.
            IF FundSum#(Fund) <> 0 THEN

              'TranCashTot# = TranCashTot# + FundSum#(Fund)
              'FundCashAcct$ = FundList$(Fund) + CshAcct$
              'FundCashAcctRec = FindAcct(FundCashAcct$)
              'IF FundCashAcctRec > 0 THEN
              '  FundCashAcctName$ = GetAcctTitle(FundCashAcctRec)
              'ELSE
              '  FundCashAcctName$ = "Incorrect Cash Acct"
              '  WarningFlag = True
              'END IF
              
              'ToPrint$ = SPACE$(80)
              'MID$(ToPrint$, 4) = FundCashAcct$
              'MID$(ToPrint$, 21) = FundCashAcctName$
              'MID$(ToPrint$, CashCol) = FUsing$(STR$(FundSum#(Fund)), CommaFmt$)
              'PRINT #PrnFileNum, ToPrint$
              'IF LineCnt > MaxLines THEN
              '  PRINT #PrnFileNum, FF$
              '  GOSUB PrnCJPageHeader
              'END IF
              
              '--print Central Depository Entries here
              'IF CDActive$ = "Y" THEN
              '  ToPrint$ = SPACE$(80)
              '  MID$(ToPrint$, 4) = CDDue$ + FundList$(Fund)
              '  MID$(ToPrint$, 21) = "Due to Fund " + FundList$(Fund)
              '  MID$(ToPrint$, TranCol) = FUsing$(STR$(FundSum#(Fund)), CommaFmt$)
              '  PRINT #PrnFileNum, ToPrint$
              '  TotTranDist# = TotTranDist# + FundSum#(Fund)
              '  IF LineCnt > MaxLines THEN
              '    PRINT #PrnFileNum, FF$
              '    GOSUB PrnCJPageHeader
              '  END IF
              'END IF
            
            END IF
          NEXT  'Fund
          
          'IF CDActive$ = "Y" THEN
          '  ToPrint$ = SPACE$(80)
          '  MID$(ToPrint$, 4) = CDCash$ '+ FundList$(Fund)
          '  MID$(ToPrint$, 21) = "Central Depository"
          '  MID$(ToPrint$, CashCol) = FUsing$(STR$(CJEdit.Amt), CommaFmt$)
          '  PRINT #PrnFileNum, ToPrint$
          '  TranCashTot# = TranCashTot# + CJEdit.Amt
          '  IF LineCnt > MaxLines THEN
          '    PRINT #PrnFileNum, FF$
          '    GOSUB PrnCJPageHeader
          '  END IF
          '
          'END IF

          '--Summary line after last distributions
          ToPrint$ = SPACE$(80)
          'MID$(ToPrint$, DebitCol) = DistSumLine$
          'MID$(ToPrint$, CreditCol) = DistSumLine$
          MID$(ToPrint$, TranCol) = DistSumLine$
          PRINT #PrnFileNum, ToPrint$
          IF LineCnt > MaxLines THEN
            PRINT #PrnFileNum, FF$
            GOSUB PrnCJPageHeader
          END IF
          
          '--Transaction Distribution Totals
          ToPrint$ = SPACE$(80)
          MID$(ToPrint$, 4) = "Total Distributed"
          'MID$(ToPrint$, DebitCol) = FUsing$(STR$(TranCashTot#), CommaFmt$)
          'MID$(ToPrint$, CreditCol) = FUsing$(STR$(TotTranDist#), CommaFmt$)
          MID$(ToPrint$, TranCol) = FUsing$(STR$(TotTranDist#), CommaFmt$)
          PRINT #PrnFileNum, ToPrint$
          IF LineCnt > MaxLines THEN
            PRINT #PrnFileNum, FF$
            GOSUB PrnCJPageHeader
          END IF
          
          FOR Cnt = 1 TO NumFunds
            FundSum#(Cnt) = 0
          NEXT
          
          '--Last distribution for this transaction - exit loop
          EXIT FOR
          
        ELSE    '--we have active transactions
          '--Add distribution to total
          TotTranDist# = TotTranDist# + Round#(CJEdit.DIST(AcctDist).DAMT)
          
          '--Add distribution to proper fund
          Found = False
          FOR Fund = 1 TO NumFunds
            FundNum$ = LEFT$(CJEdit.DIST(AcctDist).DACN, FundLen)
            IF FundNum$ = FundList$(Fund) THEN
              Found = True
              FundSum#(Fund) = FundSum#(Fund) + Round#(CJEdit.DIST(AcctDist).DAMT)
              FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(CJEdit.DIST(AcctDist).DAMT)
              EXIT FOR
            END IF
          NEXT
          
          IF NOT Found THEN     'OShit = True
            OK = MsgBox("GL", "NOFUND")
          END IF
          
          '--Print this distribution
          ToPrint$ = SPACE$(80)
          MID$(ToPrint$, 4) = CJEdit.DIST(AcctDist).DACN
          MID$(ToPrint$, 21) = CJEdit.DIST(AcctDist).DACNM
          MID$(ToPrint$, TranCol) = FUsing$(STR$(CJEdit.DIST(AcctDist).DAMT), CommaFmt$)
          PRINT #PrnFileNum, ToPrint$
          IF LineCnt > MaxLines THEN
            PRINT #PrnFileNum, FF$
            GOSUB PrnCJPageHeader
          END IF

        END IF  'Active transaction test
      NEXT      'Distribution
      
      '--2 blank lines before next distribution
      ToPrint$ = STRING$(80, "=")
      PRINT #PrnFileNum, ToPrint$
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrnCJPageHeader
      END IF
      
      PRINT #PrnFileNum,
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrnCJPageHeader
      END IF

    END IF      'Not deleted test
  NEXT          'Transaction
  
  'PRINT #PrnFileNum, STRING$(80, "-")
  
  '--Summary
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "File Totals:"
  PRINT #PrnFileNum, ToPrint$
  IF LineCnt > MaxLines THEN
    PRINT #PrnFileNum, FF$
    GOSUB PrnCJPageHeader
  END IF
  
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "Number of Transactions"
  MID$(ToPrint$, 31) = FUsing$(STR$(TranCnt), "####")
  PRINT #PrnFileNum, ToPrint$
  IF LineCnt > MaxLines THEN
    PRINT #PrnFileNum, FF$
    GOSUB PrnCJPageHeader
  END IF
  
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "Grand Totals"
  MID$(ToPrint$, 25) = FUsing$(STR$(TransTotal#), CommaFmt$)
  PRINT #PrnFileNum, ToPrint$
  IF LineCnt > MaxLines THEN
    PRINT #PrnFileNum, FF$
    GOSUB PrnCJPageHeader
  END IF
  

  PRINT #PrnFileNum,
  IF LineCnt > MaxLines THEN
    PRINT #PrnFileNum, FF$
    GOSUB PrnCJPageHeader
  END IF
  
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "Summary by Fund:"
  PRINT #PrnFileNum, ToPrint$
  IF LineCnt > MaxLines THEN
    PRINT #PrnFileNum, FF$
    GOSUB PrnCJPageHeader
  END IF
  
  FOR Cnt = 1 TO NumFunds
    IF FundGrdTot#(Cnt) <> 0 THEN
      ToPrint$ = SPACE$(80)
      LSET ToPrint$ = "Fund" + " " + FundList$(Cnt)
      MID$(ToPrint$, 25) = FUsing$(STR$(FundGrdTot#(Cnt)), CommaFmt$)
      PRINT #PrnFileNum, ToPrint$
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrnCJPageHeader
      END IF
      GrdTot# = GrdTot# + Round#(FundGrdTot#(Cnt))
    END IF
  NEXT
  
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "Total All Funds"
  MID$(ToPrint$, 25) = FUsing$(STR$(GrdTot#), CommaFmt$)
  PRINT #PrnFileNum, ToPrint$
  PRINT #PrnFileNum, FF$

  CLOSE
  
  SELECT CASE Dev$
  CASE "S"
    EntryPoint = 2
  CASE "P"
    EntryPoint = 5
  END SELECT

  PrintRptFile RegTitle$, FileName$, LPTNo, RetCode%, EntryPoint
  'KILL FileName$
  
  IF WarningFlag THEN
    OK = MsgBox%("GL.QSL", "ERRSETUP")
  END IF
  
  EXIT SUB

PrnCJPageHeader:
  PRINT #PrnFileNum, User$
  PRINT #PrnFileNum, RegTitle$
  PRINT #PrnFileNum,
  PRINT #PrnFileNum, "Date      Description          Reference   Bank                             Amt"
  PRINT #PrnFileNum, STRING$(80, "=")
  LineCnt = 4
RETURN


END SUB

SUB PrnEdList

  SHARED Choice$()              'Multiple choice array
  REDIM Choice$(0 TO 4, 1)      'Choices for entry type field
  Choice$(0, 0) = "2"           'Field Number
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn #1"
  Choice$(3, 0) = "Prn #2"
  
  REDIM Form$(0, 0)             'Holds data from form
  REDIM Fld(0) AS FieldInfo     'Field editing array
  DIM Frm AS FormInfo
  LibLoadDisplayForm "GL.QSL", "PRNGJ", Form$(), Fld()

  PrintHelp "Print Register"
  Action = 1

  '--Get the report options from user
  DO

    EditForm Form$(), Fld(), Frm, Cnf, Action

    '--Check the user's key presses
    SELECT CASE Frm.KeyCode

    CASE F10Key 'SaveButton
      IF LEN(LTRIM$(RTRIM$(Form$(2, 0)))) = 0 THEN              'Fld#2 = OutputDevice
        Dev$ = "S"
        LPTNo = 1
      ELSE
        Dev$ = LEFT$(Form$(2, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(2, 0)), 1))
      END IF

    CASE ESC
      EXIT SUB

    END SELECT

  LOOP UNTIL Frm.KeyCode = F10Key

  '==Start processing register=>
  LOCATE 1, 1, 0
  
  'SHARED CashAcct$, User$, CDCash$, CDDue$, CDActive$, CRCashAcct$

  '--For Central Depository
  IF CDActive$ = "Y" THEN
    'assign
    CshAcct$ = CashAcct$

    '--chars to pad object code if obj code > fund code
    PadChars = DetLen - FundLen
    IF PadChars > 0 THEN
      DetPad$ = STRING$(PadChars, "0")
    END IF
  END IF

  '--Check for valid cash account
  WarningFlag = False

  REDIM Title$(1 TO 4)
  
  TranCol = 58
  CommaFmt$ = "#########,.##"   'ten millions
  DistSumLine$ = "-------------"
  TransTotal# = 0
  TranCnt = 0
  FF$ = CHR$(12)
  MaxLines = 55
  
  '--Create a fund array to hold transaction totals by fund
  REDIM FundList$(1)
  GetFundList FundList$(), NumFunds
  REDIM TrFundSum#(1 TO NumFunds)
  REDIM FundGrdTot#(1 TO NumFunds)
  REDIM FundSum#(1 TO NumFunds)

  '--open the interface file
  DIM Tr2Post AS GLTransRecType
  GLIFRecLen = LEN(Tr2Post)
  IFFile = FREEFILE
  OPEN TmpIFFile$ FOR RANDOM AS IFFile LEN = GLIFRecLen
  NumGLIFRecs = LOF(IFFile) \ GLIFRecLen
  IF NumGLIFRecs > 0 THEN
    CLOSE
    CLS
    QPrintRC "WARNING", 9, 1, 12
    QPrintRC "Interface file already exists.  Possible problems with last posting.", 10, 1, 12
    QPrintRC "Contact Software support. Press any key to continue.", 11, 1, 12
    k$ = INPUT$(1)
    EXIT SUB
  END IF
  
  '--open a report file to print to
  PrnFileNum = FREEFILE
  OPEN FileName$ FOR OUTPUT AS #PrnFileNum
  
  '--open the edit file
  OpenCJEditFile CJEditFile, NumEdTrans, CJType

  SELECT CASE CJType
  CASE 1
    FileName$ = "CRREG.PRN"
    RegTitle$ = "Cash Receipts Journal"
    'TranCol = CreditCol
    'CashCol = DebitCol
    'CshAcct$ = CRCashAcct$
    CJEditFile$ = "GLCRED.DAT"
    TmpIFFile$ = "CJRPOST.DAT"
    MSrc$ = "CR" + ConvDateStr$(DATE$)
    'CshAcct$ = CRCashAcct$
  CASE 2
    FileName$ = "CDREG.PRN"
    RegTitle$ = "Cash Disbursements Journal"
    'TranCol = DebitCol
    'CashCol = CreditCol
    'CshAcct$ = CashAcct$
    CJEditFile$ = "GLCDED.DAT"
    TmpIFFile$ = "CJDPOST.DAT"
    MSrc$ = "CD" + ConvDateStr$(DATE$)
    'CshAcct$ = CashAcct$
  END SELECT

  GOSUB PrintCJPageHeader

  FOR transaction = 1 TO NumEdTrans
    GET CJEditFile, transaction, CJEdit
    IF NOT CJEdit.DELFLAG THEN
      TranCnt = TranCnt + 1
      TransTotal# = TransTotal# + Round#(CJEdit.Amt)

      IF CDActive$ <> "Y" THEN
        '--If we're NOT using a central depository fund then use
        '--the bank gl account number
        BankNum = VAL(CJEdit.RecCode)
        IF BankNum > 0 THEN
          CshAcct$ = GetBankGLAcct$(BankNum)
        END IF
      END IF
      
      '--Print 1st Line - Transaction details
      ToPrint$ = SPACE$(80)
      LSET ToPrint$ = ConvDateStr1$(Num2Date(CJEdit.TRDATE))
      MID$(ToPrint$, 10) = CJEdit.DESC
      MID$(ToPrint$, 32) = CJEdit.DOCREF
      MID$(ToPrint$, 68) = FUsing$(STR$(CJEdit.Amt), CommaFmt$)
      PRINT #PrnFileNum, ToPrint$
      LineCnt = LineCnt + 1
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrintCJPageHeader
      END IF
      
      '--Blank line between detail and acct'g distributions
      PRINT #PrnFileNum,
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrintCJPageHeader
      END IF

      '--Print Distribution Label
      ToPrint$ = SPACE$(80)
      MID$(ToPrint$, 2) = "Accounting Distribution:"
      PRINT #PrnFileNum, ToPrint$
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrintCJPageHeader
      END IF

      '--Print Field Titles
      ToPrint$ = SPACE$(80)
      MID$(ToPrint$, 4) = "Account Number   Name                                       Amount"
      PRINT #PrnFileNum, ToPrint$
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrintCJPageHeader
      END IF

      '--Print Accounting Distributions
      TotTranDist# = 0

      '--Loop Thru distributions to print and summarize
      FOR AcctDist = 1 TO 36

        IF LEN(QPTrim$(CJEdit.DIST(AcctDist).DACN)) > 0 THEN

          '--Add distribution to total
          TotTranDist# = TotTranDist# + Round#(CJEdit.DIST(AcctDist).DAMT)
          
          '--Add distribution to proper fund
          FOR Fund = 1 TO NumFunds
            FundNum$ = LEFT$(CJEdit.DIST(AcctDist).DACN, FundLen)
            IF FundNum$ = FundList$(Fund) THEN
              TrFundSum#(Fund) = TrFundSum#(Fund) + Round#(CJEdit.DIST(AcctDist).DAMT)
              FundGrdTot#(Fund) = FundGrdTot#(Fund) + Round#(CJEdit.DIST(AcctDist).DAMT)
              EXIT FOR
            END IF
          NEXT
          
          '--Print this distribution
          ToPrint$ = SPACE$(80)
          MID$(ToPrint$, 4) = CJEdit.DIST(AcctDist).DACN
          MID$(ToPrint$, 21) = CJEdit.DIST(AcctDist).DACNM
          MID$(ToPrint$, TranCol) = FUsing$(STR$(CJEdit.DIST(AcctDist).DAMT), CommaFmt$)
          PRINT #PrnFileNum, ToPrint$
          IF LineCnt > MaxLines THEN
            PRINT #PrnFileNum, FF$
            GOSUB PrintCJPageHeader
          END IF
         
          '--write the distribution to the interface file
          RecordNum = RecordNum + 1
          Tr2Post.AcctNum = CJEdit.DIST(AcctDist).DACN
          Tr2Post.TRDATE = CJEdit.TRDATE
          Tr2Post.DESC = CJEdit.DESC
          Tr2Post.Ref = CJEdit.DOCREF
          '--Set debit and credit amts and src ref according to Jrn Type
          SELECT CASE CJType
            CASE 1 'Cash Reciepts - distributions are credits
              Tr2Post.DrAmt = 0
              Tr2Post.CrAmt = CJEdit.DIST(AcctDist).DAMT
            CASE 2 'Cash Disbursements - distributions are debits
              Tr2Post.DrAmt = CJEdit.DIST(AcctDist).DAMT
              Tr2Post.CrAmt = 0
          END SELECT
          Tr2Post.Src = MSrc$
          PUT IFFile, RecordNum, Tr2Post
          
        END IF  'valid distribution test
      NEXT  'Distribution

      '--No more distributions
      ToPrint$ = STRING$(80, "=")
      PRINT #PrnFileNum, ToPrint$
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrintCJPageHeader
      END IF

      PRINT #PrnFileNum,
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrintCJPageHeader
      END IF

      TranCashTot# = 0
      FOR Fund = 1 TO NumFunds
        '--Test each fund to see if it has any distributions.
        IF TrFundSum#(Fund) <> 0 THEN
          TranCashTot# = TranCashTot# + TrFundSum#(Fund)
          FundCashAcct$ = FundList$(Fund) + CshAcct$
          FundCashAcctRec = FindAcct(FundCashAcct$)
          IF FundCashAcctRec > 0 THEN
            FundCashAcctName$ = GetAcctTitle(FundCashAcctRec)
          ELSE
            FundCashAcctName$ = "Incorrect Cash Acct"
            WarningFlag = True
          END IF

          ToPrint$ = SPACE$(80)
          MID$(ToPrint$, 4) = FundCashAcct$
          MID$(ToPrint$, 21) = FundCashAcctName$
          MID$(ToPrint$, CashCol) = FUsing$(STR$(FundSum#(Fund)), CommaFmt$)
          PRINT #PrnFileNum, ToPrint$
          IF LineCnt > MaxLines THEN
            PRINT #PrnFileNum, FF$
            GOSUB PrintCJPageHeader
          END IF

          '--print Central Depository Entries here
          IF CDActive$ = "Y" THEN
            ToPrint$ = SPACE$(80)
            MID$(ToPrint$, 4) = CDDue$ + FundList$(Fund)
            MID$(ToPrint$, 21) = "Due to Fund " + FundList$(Fund)
            MID$(ToPrint$, TranCol) = FUsing$(STR$(FundSum#(Fund)), CommaFmt$)
            PRINT #PrnFileNum, ToPrint$
            TotTranDist# = TotTranDist# + FundSum#(Fund)
            IF LineCnt > MaxLines THEN
              PRINT #PrnFileNum, FF$
              GOSUB PrintCJPageHeader
            END IF
          END IF

        END IF
      NEXT  'Fund

      IF CDActive$ = "Y" THEN
        ToPrint$ = SPACE$(80)
        MID$(ToPrint$, 4) = CDCash$ '+ FundList$(Fund)
        MID$(ToPrint$, 21) = "Central Depository"
        MID$(ToPrint$, CashCol) = FUsing$(STR$(CJEdit.Amt), CommaFmt$)
        PRINT #PrnFileNum, ToPrint$
        TranCashTot# = TranCashTot# + CJEdit.Amt
        IF LineCnt > MaxLines THEN
          PRINT #PrnFileNum, FF$
          GOSUB PrintCJPageHeader
        END IF
      END IF

      '--Summary line after last distributions
      ToPrint$ = SPACE$(80)
      MID$(ToPrint$, TranCol) = DistSumLine$
      PRINT #PrnFileNum, ToPrint$
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrintCJPageHeader
      END IF

      '--Transaction Distribution Totals
      ToPrint$ = SPACE$(80)
      MID$(ToPrint$, 4) = "Total Distributed"
      MID$(ToPrint$, TranCol) = FUsing$(STR$(TotTranDist#), CommaFmt$)
      PRINT #PrnFileNum, ToPrint$
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrintCJPageHeader
      END IF

      FOR Cnt = 1 TO NumFunds
        FundSum#(Cnt) = 0
      NEXT
      
    END IF      'Not deleted test
  NEXT          'Transaction

  'PRINT #PrnFileNum, STRING$(80, "-")

  '--Summary
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "File Totals:"
  PRINT #PrnFileNum, ToPrint$
  IF LineCnt > MaxLines THEN
    PRINT #PrnFileNum, FF$
    GOSUB PrintCJPageHeader
  END IF

  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "Number of Transactions"
  MID$(ToPrint$, 31) = FUsing$(STR$(TranCnt), "####")
  PRINT #PrnFileNum, ToPrint$
  IF LineCnt > MaxLines THEN
    PRINT #PrnFileNum, FF$
    GOSUB PrintCJPageHeader
  END IF

  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "Grand Totals"
  MID$(ToPrint$, 25) = FUsing$(STR$(TransTotal#), CommaFmt$)
  PRINT #PrnFileNum, ToPrint$
  IF LineCnt > MaxLines THEN
    PRINT #PrnFileNum, FF$
    GOSUB PrintCJPageHeader
  END IF
  
  PRINT #PrnFileNum,
  IF LineCnt > MaxLines THEN
    PRINT #PrnFileNum, FF$
    GOSUB PrintCJPageHeader
  END IF

  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "Summary by Fund:"
  PRINT #PrnFileNum, ToPrint$
  IF LineCnt > MaxLines THEN
    PRINT #PrnFileNum, FF$
    GOSUB PrintCJPageHeader
  END IF

  FOR Cnt = 1 TO NumFunds
    IF FundGrdTot#(Cnt) > 0 THEN
      ToPrint$ = SPACE$(80)
      LSET ToPrint$ = "Fund" + " " + FundList$(Cnt)
      MID$(ToPrint$, 25) = FUsing$(STR$(FundGrdTot#(Cnt)), CommaFmt$)
      PRINT #PrnFileNum, ToPrint$
      IF LineCnt > MaxLines THEN
        PRINT #PrnFileNum, FF$
        GOSUB PrintCJPageHeader
      END IF
      GrdTot# = GrdTot# + Round#(FundGrdTot#(Cnt))
    END IF
  NEXT

  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "Total All Funds"
  MID$(ToPrint$, 25) = FUsing$(STR$(GrdTot#), CommaFmt$)
  PRINT #PrnFileNum, ToPrint$
  PRINT #PrnFileNum, FF$

  CLOSE

  SELECT CASE Dev$
  CASE "S"
    EntryPoint = 2
  CASE "P"
    EntryPoint = 5
  END SELECT

  PrintRptFile RegTitle$, FileName$, LPTNo, RetCode%, EntryPoint
  'KILL FileName$

  IF WarningFlag THEN
    OK = MsgBox%("GL.QSL", "ERRSETUP")
  END IF

  EXIT SUB

PrintCJPageHeader:
  PRINT #PrnFileNum, User$
  PRINT #PrnFileNum, RegTitle$
  PRINT #PrnFileNum,
  PRINT #PrnFileNum, "Date      Description          Reference                                    Amt"
  PRINT #PrnFileNum, STRING$(80, "=")
  LineCnt = 4
RETURN

'STOP
'---------
'              IF CDActive$ = "Y" THEN
'                IF PadChars > 0 THEN
'                  FundDue$ = FundList$(Fund) + DetPad$
'                ELSE
'                  FundDue$ = FundList$(Fund)
'                END IF
'              END IF
'
'              RecordNum = RecordNum + 1
'              Tr2Post.AcctNum = FundList$(Fund) + CshAcct$
'              Tr2Post.TRDATE = CJEdit.TRDATE
'              Tr2Post.DESC = CJEdit.DESC
'              Tr2Post.Ref = CJEdit.DOCREF
'              SELECT CASE CJType
'                '--Cash Reciepts - cash is debited
'                CASE 1
'                 Tr2Post.DrAmt = TrFundSum#(Fund)
'                 Tr2Post.CrAmt = 0
'                 '-Update Central Depository
'                 IF CDActive$ = "Y" THEN
'                   CDDrAcct$ = CDCash$
'                   CDCrAcct$ = CDDue$ + FundDue$
'                 END IF
'                '--Cash Disbursements - cash is credited
'                CASE 2
'                  Tr2Post.DrAmt = 0
'                  Tr2Post.CrAmt = TrFundSum#(Fund)
'                  '-Update Central Depository
'                  IF CDActive$ = "Y" THEN
'                    CDDrAcct$ = CDDue$ + FundDue$
'                    CDCrAcct$ = CDCash$
'                  END IF
'              END SELECT
'              Tr2Post.Src = MSrc$
'              PUT IFFile, RecordNum, Tr2Post
'
'              '--make entries to Central Depository Here.
'              IF CDActive$ = "Y" THEN
'
'                '--Debit to central depository
'                RecordNum = RecordNum + 1
'                Tr2Post.AcctNum = CDDrAcct$
'                Tr2Post.TRDATE = CJEdit.TRDATE
'                Tr2Post.DESC = CJEdit.DESC
'                Tr2Post.Ref = CJEdit.DOCREF
'                Tr2Post.DrAmt = TrFundSum#(Fund)
'                Tr2Post.CrAmt = 0
'                Tr2Post.Src = MSrc$
'                PUT IFFile, RecordNum, Tr2Post
'
'                '--Credit to central depository
'                RecordNum = RecordNum + 1
'                Tr2Post.AcctNum = CDCrAcct$
'                Tr2Post.TRDATE = CJEdit.TRDATE
'                Tr2Post.DESC = CJEdit.DESC
'                Tr2Post.Ref = CJEdit.DOCREF
'                Tr2Post.DrAmt = 0
'                Tr2Post.CrAmt = TrFundSum#(Fund)
'                Tr2Post.Src = MSrc$
'                PUT IFFile, RecordNum, Tr2Post
'
'              END IF
'
'            'END IF              'Fund summary <> 0
'          'NEXT  'fund
'
'          'EXIT FOR              'NumFunds
'
'        ELSE
'          '--Not the last distribution so add this distribution to the
'          '--interface file.
'
'
'        END IF  'test for blank distribution line
'      NEXT    'Acct'g Distribution
'    END IF  'not deleted
'  NEXT    'transaction
'
END SUB

SUB Rec2Form (Form$(), RecordNumber%, Fld() AS FieldInfo, CJType)
  
  OpenCJEditFile CJEditFile, NumEdTrans, CJType
  GET CJEditFile, RecordNumber, CJEdit
  Form$(0, 0) = SPACE$(LEN(CJEdit))
  BCopy VARSEG(CJEdit), VARPTR(CJEdit), SSEG(Form$(0, 0)), SADD(Form$(0, 0)), LEN(Form$(0, 0)), 0
  UnPackBuffer 0, 0, Form$(), Fld()
  CLOSE CJEditFile              'close the file
  
END SUB

SUB ResetForm (Form$(), NumFlds)
  
  FOR F = 1 TO NumFlds
    Form$(F, 0) = ""
  NEXT
  

END SUB

SUB SaveTrans (Form$(), RecordNumber, EdMode, CJType)
  
  OpenCJEditFile CJEditFile, NumEdTrans, CJType
  
  IF EdMode = 1 THEN
    Rec = NumEdTrans + 1
  ELSE
    Rec = RecordNumber
  END IF
  
 'DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
  BCopy SSEG(Form$(0, 0)), SADD(Form$(0, 0)), VARSEG(CJEdit), VARPTR(CJEdit), LEN(Form$(0, 0)), 0

  PUT CJEditFile, Rec, CJEdit   'write it
  CLOSE CJEditFile              'close the file
END SUB

SUB ShowEdMode (EdMode%)
  
  Status$ = SPACE$(24)
  F2Ed$ = "  F2=Edit  "
  F2New$ = "  F2=New   "
  ActiveColor = 112
  InActiveColor = 115
  
  SELECT CASE EdMode
  CASE 0
    LSET Status$ = "Edit an Existing Record"
    F2BtnCap$ = F2New$
    F5BtnColor = ActiveColor
    
  CASE 1
    LSET Status$ = "Add a New Record"
    F2BtnCap$ = F2Ed$
    F5BtnColor = InActiveColor
    
  END SELECT
  
  QPrintRC Status$, 25, 2, -1
  QPrintRC F2BtnCap$, 25, 28, -1
  MPaintBox 25, 41, 25, 51, F5BtnColor          'F5 Button enable / disable
  
END SUB

