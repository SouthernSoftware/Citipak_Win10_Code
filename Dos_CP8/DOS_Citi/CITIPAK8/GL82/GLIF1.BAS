  'GLIF.BAS
  DEFINT A-Z
  
DECLARE SUB ViewIFFile ()
DECLARE SUB Interface ()
DECLARE SUB PackIFFile (IFFileName$)
DECLARE SUB Prn2Scrn2 (TitleText$(), FileName$, TopRow%, BotRow%, Clr%, PrnFlag%)
  
  'DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
  'DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
  'DECLARE FUNCTION Exist% (FileName$)
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
DECLARE SUB QPrintRC (t$, r%, c%, Clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
  
  '$INCLUDE: 'DefCnf.bi'                    'monitor and color info
  '$INCLUDE: 'FieldInf.bi'                  'field information
  '$INCLUDE: 'FormEdit.bi'                  'form editing information
  '$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
  '$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
  '$INCLUDE: 'RPTPRN.BI'                    'Report printer routines
  '$INCLUDE: 'GL.bi'                        'GL File Types
  '$INCLUDE: 'SCRAUX.BI'
  
  DIM SHARED TotLen%, FundLen%, AcctLen%, DetLen%
  DIM SHARED Acct AS GLAcctRecType
  DIM SHARED AcctIdx AS GLAcctIndexType
  
  '$INCLUDE: 'SetCnf.bi'                    'Set monitor and color information
  '$INCLUDE: 'init.bi'                      'Get GL system info
  ''$INCLUDE: 'keycode.bi'                   'Define keycodes
  
  CONST False = 0
  CONST True = NOT False
  
  REDIM SHARED Choice$(0, 0)    'Used to hold multiple choices
  
  REDIM MChoice$(1 TO 2)
  
  MChoice$(1) = "View Un-Posted Transactions"
  MChoice$(2) = "Post Interface File to G/L"
  
  MaxWidth = 0  'Set menu width to zero
  BoxBot = 20   'limit the box to line 20
  Action = 0    'Stay on menu til a choice is made
  Choice = 1    'Pre-load choice to highlight
  
  FOR Cnt = 1 TO UBOUND(MChoice$)               'Find max menu width
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxWidth THEN
      MaxWidth = TLen
    END IF
  NEXT
  
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1     'Center Menu within Screen
  Col = ((80 - MaxWidth) \ 2) - 2
  
  TextCursor -2, -2             'Define the Mouse cursor
  ShowCursor    'Show it
  
  DO
    
    LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, Attribute, ErrorCode
    PrintTitle User$
    PrintHelp "General Ledger Interface"
    
    LOCATE Row, Col, 0
    VertMenu MChoice$(), Choice, MaxWidth, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO
    
    SELECT CASE Choice
    CASE 1
      ViewIFFile
    CASE 2
      Interface
    END SELECT
    
  LOOP
  
  HideCursor    'Be nice and turn off mouse cursor
  DispLoad
  RUN "glmenu"
  
SUB Interface
  
  DIM Trans AS GLTransRecType
  SHARED FundLen%, AcctLen%, DetLen%
  SHARED Choice$()
  
  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "AR-Accts Receivable"
  Choice$(2, 0) = "AP-Accts Payable"
  Choice$(3, 0) = "PR-Payroll"
  
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  
  LibName$ = "GL.QSL"
  FormName$ = "CPIF"
  
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    SELECT CASE Frm.KeyCode     'Check the user's key presses
    CASE F10Key 'SaveButton
      'ExitFlag = True
      IF LEN(QPTrim$(Form$(1, 0))) = 2 THEN
        LowDate = GetBegDate
        EndDate = GetEndDate
        InfDate = Date2Num(Form$(2, 0))
        IF InfDate >= LowDate AND InfDate <= EndDate THEN
          EXIT DO
        ELSE
          BEEP
          QPrintRC "Invalid date or date out of Fiscal Year.", 25, 2, -1
          Frm.FldNo = 2
          Action = 1
        END IF
      ELSE
        BEEP
        QPrintRC "No module selected.                          ", 25, 2, -1
        Frm.FldNo = 1
        Action = 1
      END IF
      
    CASE EscKey
      ExitFlag = True
      EXIT DO
      'Check to see if form was edited before ESC was pressed
      
    END SELECT
    
  LOOP          'UNTIL ExitFlag
  
  IF ExitFlag THEN EXIT SUB
  
  '--Get the input from the form
  Module$ = LEFT$(Form$(1, 0), 2)
  EndDate = Date2Num(Form$(2, 0))
  
  '--Assign Report Strings
  'RptTitle$ = "Posting Report: " + ModuleFile$
  REDIM Desc$(1)
  Desc$(1) = "Date     Account    Description         Ref            Debit     Credit  Record"
  
  ModuleFile$ = Module$ + "IF.DAT"
  DIM IFRec AS IFRecType
  IFRecLen = LEN(IFRec)
  
  IFFile = FREEFILE
  OPEN ModuleFile$ FOR RANDOM AS IFFile LEN = IFRecLen
  NumOFRecs = LOF(1) \ IFRecLen
  
  OpenAcctFile AcctFileNum, NumAccts
  OpenTransFile TransFileNum, NumTrans&
  
  PrnFileName$ = "ifrpt.prn"
  PrnFile = FREEFILE
  OPEN PrnFileName$ FOR OUTPUT AS PrnFile
  
  ErrLogName$ = "errlog.prn"
  ErrLogFile = FREEFILE
  OPEN ErrLogName$ FOR OUTPUT AS ErrLogFile
  
  TotDr# = 0
  TotCr# = 0
  ErrTotDr# = 0
  ErrTotCr# = 0
  
  FOR Cnt = 1 TO NumOFRecs
    QPrintRC STR$(Cnt), 25, 1, -1
    GET IFFile, Cnt, IFRec
    IF NOT IFRec.Posted THEN
      IF Date2Num(IFRec.TrDate) <= EndDate THEN
        AN$ = QPTrim$(IFRec.AcctNum)
        AN$ = FmtAcct$(AN$, FundLen%, AcctLen%, DetLen%)
        AcctRec = FindAcct(AN$)
        IF AcctRec > 0 THEN     'good acct
          '--tell user what's going on
          QPrintRC " Posting Account Number: ", 25, 1, 112
          QPrintRC IFRec.AcctNum, 25, 26, 112
          
          GET AcctFileNum, AcctRec, Acct        'Get the account
          
          '--depending on account type, update running balance
          SELECT CASE Acct.Typ
          CASE "A", "E"         'asset, exp accts
            Acct.Bal = Round#(Acct.Bal) + Round#(IFRec.DrAmt) - Round#(IFRec.CrAmt)
            PUT AcctFileNum, AcctRec, Acct
            
          CASE "L", "R"         'liab, rev accts
            Acct.Bal = Round#(Acct.Bal) + Round#(IFRec.CrAmt) - Round#(IFRec.DrAmt)
            PUT AcctFileNum, RecNum, Acct
            
          END SELECT
          
          NumTrans& = NumTrans& + 1             'increment record pointer
          GET TransFileNum, NumTrans&, Trans
          Trans.AcctNum = IFRec.AcctNum
          Trans.TrDate = Date2Num(IFRec.TrDate)
          Trans.Desc = IFRec.Desc
          Trans.CrAmt = IFRec.CrAmt
          Trans.DrAmt = IFRec.DrAmt
          Trans.Ref = IFRec.Ref
          Trans.Src = IFRec.Src
          Trans.NextTran = 0
          PUT TransFileNum, NumTrans&, Trans
          
          GOSUB ReportIt
          IFRec.Posted = -1
          PUT IFFile, Cnt, IFRec
          
          '---------------------------------Start linking here
          IF Acct.FrstTran = 0 THEN             'if first trans for this acct,
            Acct.FrstTran = NumTrans&           'assign first & last pointers to
            Acct.LastTran = NumTrans&           'this transaction
            PUT AcctFileNum, RecNum, Acct
            
          ELSE  'otherwise
            'in the account file..
            Prev& = Acct.LastTran               'remember the prev trans pointer,
            Acct.LastTran = NumTrans&           'reset last trans to this trans
            PUT AcctFileNum, AcctRec, Acct
            
            'In the trans file...
            GET TransFileNum, Prev&, Trans      'Get the last transaction
            Trans.NextTran = NumTrans&          'reset pointer to this trans
            PUT TransFileNum, Prev&, Trans
            
          END IF
          
        ELSE
          '--Invalid Account Number!
          BadTrans = BadTrans + 1
          GOSUB ErrorLog
          
        END IF  'Valid Acct Check
      END IF    'Date Check
    END IF      'Posted Check
    
  NEXT
  
  PRINT #PrnFile, ""
  PRINT #PrnFile, USING "Total Debits   :#########,.##"; Round#(TotDr#)
  PRINT #PrnFile, USING "Total Credits  :#########,.##"; Round#(TotCr#)
  
  Diff# = TotDr# - TotCr#
  PRINT #PrnFile, USING "Difference   : :#########,.##"; Round#(Diff#)
  
  IF BadTrans THEN
    PRINT #ErrLogFile, ""
    PRINT #ErrLogFile, USING "Total Debits   :#########,.##"; Round#(ErrTotDr#)
    PRINT #ErrLogFile, USING "Total Credits  :#########,.##"; Round#(ErrTotCr#)
  END IF
  
  CLOSE
  
  CLS
  QPrintRC "Posting Report:" + ModuleFile$, 1, 1, 3
  QPrintRC "Esc to exit. F7 to Print.", 25, 1, 3
  Prn2Scrn2 Desc$(), PrnFileName$, 6, 21, 15, PrnFlag%
  IF PrnFlag = 1 THEN
    ReportFile$ = PrnFileName$
    GOSUB PrintIFReport
  END IF
  
  IF BadTrans THEN
    CLS
    QPrintRC "Error Log:" + ModuleFile$, 1, 1, 6
    QPrintRC "Esc to exit. F7 to Print.", 25, 1, 6
    Prn2Scrn2 Desc$(), ErrLogName$, 6, 21, 79, PrnFlag%
    IF PrnFlag = 1 THEN
      ReportFile$ = ErrLogName$
      GOSUB PrintIFReport
    END IF
  END IF
  
  'KILL PrnFileName$
  'KILL ErrLogName$
  
  PackIFFile ModuleFile$
  
  EXIT SUB
  
  ReportIt:
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = IFRec.TrDate
  MID$(ToPrint$, 10) = IFRec.AcctNum
  MID$(ToPrint$, 21) = IFRec.Desc
  MID$(ToPrint$, 41) = IFRec.Ref
  MID$(ToPrint$, 51) = FUsing$(STR$(IFRec.DrAmt), "######,.##")
  MID$(ToPrint$, 62) = FUsing$(STR$(IFRec.CrAmt), "######,.##")
  MID$(ToPrint$, 73) = STR$(Cnt)
  PRINT #PrnFile, ToPrint$
  
  'Add transaction totals to running totals
  TotDr# = TotDr# + IFRec.DrAmt
  TotCr# = TotCr# + IFRec.CrAmt
  
  RETURN
  
  
  ErrorLog:
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = IFRec.TrDate
  MID$(ToPrint$, 10) = IFRec.AcctNum
  MID$(ToPrint$, 21) = IFRec.Desc
  MID$(ToPrint$, 41) = IFRec.Ref
  MID$(ToPrint$, 51) = FUsing$(STR$(IFRec.DrAmt), "######,.##")
  MID$(ToPrint$, 62) = FUsing$(STR$(IFRec.CrAmt), "######,.##")
  MID$(ToPrint$, 73) = STR$(Cnt)
  PRINT #ErrLogFile, ToPrint$
  
  'Add transaction totals to running totals
  ErrTotDr# = ErrTotDr# + IFRec.DrAmt
  ErrTotCr# = ErrTotCr# + IFRec.CrAmt
  
  RETURN
  
  
  PrintIFReport:
  REDIM Title$(1 TO 4)
  Title$(1) = RptTitle$
  Title$(2) = ""
  Title$(2) = Desc$(1)
  Title$(2) = STRING$(80, "=")
  
  TitLines = 4
  MaxLines = 55
  LPTNo = 1
  Prn2Prn ReportFile$, Title$(), TitLines%, MaxLines%, LPTNo%, CancelFlag%
  
  RETURN
  
  
END SUB
  
SUB PackIFFile (IFFileName$)
  
  CLS
  QPrintRC "Please wait.  Packing Interface File...", 25, 1, 11
  
  DIM NewIFRec AS IFRecType
  IFRecLen = LEN(NewIFRec)
  NewIFFile = FREEFILE
  OPEN "IF.DAT" FOR RANDOM AS NewIFFile LEN = IFRecLen
  
  DIM IFRec AS IFRecType
  IFFile = FREEFILE
  OPEN IFFileName$ FOR RANDOM AS IFFile LEN = IFRecLen
  NumRecs = LOF(1) \ IFRecLen
  
  FOR Cnt = 1 TO NumRecs
    GET IFFile, Cnt, IFRec
    IF NOT IFFile.Posted THEN
      Keeper = Keeper + 1
      NewIFRec.TrDate = IFRec.TrDate
      NewIFRec.AcctNum = IFRec.AcctNum
      NewIFRec.Desc = IFRec.Desc
      NewIFRec.Ref = IFRec.Ref
      NewIFRec.DrAmt = IFRec.DrAmt
      NewIFRec.CrAmt = IFRec.CrAmt
      NewIFRec.Src = IFRec.Src
      NewIFRec.Posted = 0
      PUT NewIFFile, Keeper, NewIFRec
    END IF
  NEXT
  
  CLOSE
  
  KILL IFFileName$
  NAME "IF.DAT" AS IFFileName$
  
END SUB
  
SUB ViewIFFile
  
  SHARED Choice$()
  REDIM Choice$(0 TO 3, 1)      'for pick lists
  Choice$(0, 0) = "1"
  Choice$(1, 0) = "AR-Accts Receivable"
  Choice$(2, 0) = "AP-Accounts Payable"
  Choice$(3, 0) = "PR-Payroll"
  
  REDIM Form$(0, 0)             'Form array holds data while editing
  REDIM Fld(0) AS FieldInfo     'Field info Array
  DIM Frm AS FormInfo           'Form editing info
  
  LibName$ = "GL.QSL"
  FormName$ = "CPIFV"
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  
  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    SELECT CASE Frm.KeyCode     'Check the user's key presses
      
    CASE F10Key 'SaveButton
      'ExitFlag = True
      IF LEN(QPTrim$(Form$(1, 0))) = 2 THEN
        EXIT DO
      ELSE
        Frm.FldNo = 1
        Action = 1
      END IF
      
    CASE EscKey
      ExitFlag = True
      EXIT DO
      'Check to see if form was edited before ESC was pressed
      
    END SELECT
    
  LOOP          'UNTIL ExitFlag
  
  IF ExitFlag THEN EXIT SUB
  
  Module$ = LEFT$(Form$(1, 0), 2)
  ModuleFile$ = Module$ + "IF.DAT"
  
  IF LEN(QPTrim$(Form$(2, 0))) = 6 THEN
    EndDate = GetEndDate
  ELSE
    EndDate = Date2Num(Form$(2, 0))
  END IF
  
  DIM IFRec AS IFRecType
  RecLen = LEN(IFRec)
  
  OPEN ModuleFile$ FOR RANDOM AS #1 LEN = RecLen
  OPEN "IF.PRN" FOR OUTPUT AS #2
  
  NumOFRecs = LOF(1) \ RecLen
  TotDr# = 0
  TotCr# = 0
  
  QPrintRC "Reading interface file.    % complete.", 25, 1, -1
  Pct$ = SPACE$(3)
  
  FOR Cnt = 1 TO NumOFRecs
    Complete# = ((Cnt / NumOFRecs) * 100)
    Done$ = STR$(INT(Complete#))
    RSET Pct$ = Done$
    QPrintRC Done$, 25, 25, -1
    
    GET #1, Cnt, IFRec
    
    IF NOT IFRec.Posted THEN
      TransDate = Date2Num(IFRec.TrDate)
      IF TransDate <= EndDate THEN
        GOSUB ListIt
      END IF
    END IF
    
  NEXT
  
  PRINT #2, ""
  PRINT #2, USING "Total Debits   :#########,.##"; Round#(TotDr#)
  PRINT #2, USING "Total Credits  :#########,.##"; Round#(TotCr#)
  
  Diff# = TotDr# - TotCr#
  PRINT #2, USING "Difference   : :#########,.##"; Round#(Diff#)
  
  CLOSE
  
  CLS
  LOCATE 1, 1, 0
  
  RptTitle$ = "Viewing: " + ModuleFile$
  QPrintRC RptTitle$, 1, 1, 3
  QPrintRC "Esc to exit. F7 to Print.", 25, 1, 3
  
  REDIM Desc$(1)
  Desc$(1) = "Date     Account    Description         Ref            Debit     Credit  Record"
  
  Prn2Scrn2 Desc$(), "if.prn", 6, 21, 15, PrnFlag%
  IF PrnFlag = 1 THEN GOSUB PrintIF
  
  KILL "if.prn"
  
  EXIT SUB
  
  ListIt:
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = IFRec.TrDate
  MID$(ToPrint$, 10) = IFRec.AcctNum
  MID$(ToPrint$, 21) = IFRec.Desc
  MID$(ToPrint$, 41) = IFRec.Ref
  MID$(ToPrint$, 51) = FUsing$(STR$(IFRec.DrAmt), "######,.##")
  MID$(ToPrint$, 62) = FUsing$(STR$(IFRec.CrAmt), "######,.##")
  MID$(ToPrint$, 73) = STR$(Cnt)
  PRINT #2, ToPrint$
  
  'Add transaction totals to running totals
  TotDr# = TotDr# + IFRec.DrAmt
  TotCr# = TotCr# + IFRec.CrAmt
  
  RETURN
  
  
  PrintIF:
  REDIM Title$(1 TO 4)
  Title$(1) = RptTitle$
  Title$(2) = ""
  Title$(2) = Desc$(1)
  Title$(2) = STRING$(80, "=")
  
  TitLines = 4
  MaxLines = 55
  LPTNo = 1
  ReportFile$ = "IF.PRN"
  Prn2Prn ReportFile$, Title$(), TitLines%, MaxLines%, LPTNo%, CancelFlag%
  
  RETURN
  
  
END SUB
  
