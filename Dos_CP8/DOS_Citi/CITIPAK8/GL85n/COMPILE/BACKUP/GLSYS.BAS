DEFINT A-Z
DECLARE SUB GetTaxInfo ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
  
'$INCLUDE: 'Fileview.bi'
'$INCLUDE: 'DefCnf.bi'
'$INCLUDE: 'FieldInf.bi'
'$INCLUDE: 'FormEdit.bi'
'$INCLUDE: 'PRO7.bi'
'$INCLUDE: 'QSCR.BI'
'$INCLUDE: 'GLAUX.BI'
'$INCLUDE: 'GL.BI'
'$INCLUDE: 'ap.BI'
  
DECLARE SUB VertMenuT (Items() AS ANY, Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB UnderConst ()
DECLARE SUB OpenSetupFile (SetupFileNum%)
DECLARE SUB DefineAcctStruct ()
DECLARE SUB UserInfo ()
DECLARE SUB SetFYDates ()
DECLARE SUB DispLoad ()
DECLARE SUB LibLoadDisplayForm (LibName$, FormName$, Form$(), Fld() AS ANY)
DECLARE SUB MnuTitle (Title$, X%, y%, Wdth%, TitColor%)
DECLARE SUB ClrBG ()
DECLARE SUB ReLinkTrans ()
DECLARE SUB RelinkBgtTrans ()
DECLARE FUNCTION GetPassword$ (Row%, Col%, Max%, Abort%)
DECLARE SUB SaveScrn (Array())
DECLARE SUB RestScrn (Array())
  
  '$INCLUDE: 'SetCnf.BI'
  ''$INCLUDE: 'keycode.bi'
  
  TYPE FLen
    V AS STRING * 42
  END TYPE
  
  CONST False = 0, True = NOT False
  
  DIM SHARED GLSetup AS GLSetupRecType
  DIM SHARED Acct AS GLAcctRecType
  
  OpenSetupFile SetupFileNum
  GET SetupFileNum, 1, GLSetup
  User$ = GLSetup.UserName
  '   TotLen = GLSetup.TotAcctLen
  '   FundLen = GLSetup.FundLen
  '   AcctLen = GLSetup.AcctLen
  '   DetLen = GLSetup.DetLen
  CLOSE SetupFileNum
  
  
  '***Master Menu
  
  DIM mChoice$(1 TO 6)
  
  mChoice$(1) = "User File Maintenance"
  mChoice$(2) = "Define G/L Account Structure"
  mChoice$(3) = "Set Fiscal Periods"
  'mChoice$(4) = "Set Allowable Posting Date Range"
  'mChoice$(4) = "Re-Link G/L Transaction Database"
  mChoice$(4) = "Invoice Tax Setup"
  mChoice$(5) = "G/L Utilities"
  mChoice$(6) = "Quit"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 20   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  FOR Cnt = 1 TO UBOUND(mChoice$)               'Find max menu width
    TLen = LEN(mChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  'Center Menu within Screen
  Row = ((25 - (UBOUND(mChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  
  ShowCursor    'turned mouse off temporarily
  
  DO
    
    LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, Attribute, ErrorCode
    'TitleBox 3, Col, MaxLen + 3, "Configuration Menu ", Cnf
    'TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    Title$ = "Config/Utility Menu "
    MnuTitle Title$, Row, Col, MaxLen, 15
    PrintTitle User$
    PrintHelp "Control File Menu"
    
    LOCATE Row, Col, 0
    VertMenu mChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO
    
    SELECT CASE Choice
      'CASE 1
      '   UnderConst
    CASE 1
      ClrBG
      UserInfo
    CASE 2
      ClrBG
      DefineAcctStruct
    CASE 3
      ClrBG
      SetFYDates
      'CASE 4
      '   ClrBG
      '   'SetPostDates
      'CASE 4
      '   ClrBG
      '   ReLinkTrans
    CASE 4
      ClrBG
      GetTaxInfo
      
    CASE 5
      LibFile2Scrn "GL.QSL", "GETPW", MonoCode, Attribute, ErrorCode
      PW$ = GetPassword$(12, 42, 10, Abort%)
      IF PW$ = "SOSOFT" THEN
        RUN "GLUTIL"
      ELSE
        BEEP
      END IF
    CASE 6
      ClrBG
      END
    CASE ELSE
    END SELECT
    
  LOOP
  
Getout:
  ClrBG
  HideCursor
  DispLoad
  RUN "glmaint"

SUB DefineAcctStruct
  
  'SHARED Cnf AS Config                      'System config & colors
  REDIM Form$(0, 0)             'Holds data from form
  REDIM Fld(0) AS FieldInfo     'Field editing array
  
  DIM Frm AS FormInfo           'Dim the form info array
  LibLoadDisplayForm "GL.QSL", "SETUPGL", Form$(), Fld()
  PrintHelp "Define Account Structure"
  
  OpenSetupFile SetupFileNum
  
  GET SetupFileNum, 1, GLSetup
  
  Form$(1, 0) = STR$(GLSetup.FundLen)
  Form$(2, 0) = STR$(GLSetup.AcctLen)
  Form$(3, 0) = STR$(GLSetup.DetLen)
  Form$(4, 0) = STR$(GLSetup.TotAcctLen)
  
  CLOSE
  
  Action = 1    'Init form edit
  
  DO            'Edit the form
    
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    SELECT CASE Frm.KeyCode     'Respond to the key presses
    CASE F10Key 'F10=Save
      'IF AcctLen > 14 THEN
      '   Warn = MsgBox%("GL.QSL", "BADLEN")  'badlen doesn't exist
      '   RETURN
      'END IF
      
      OpenSetupFile SetupFileNum
      GLSetup.FundLen = VAL(Form$(1, 0))
      GLSetup.AcctLen = VAL(Form$(2, 0))
      GLSetup.DetLen = VAL(Form$(3, 0))
      GLSetup.TotAcctLen = VAL(Form$(4, 0))
      PUT SetupFileNum, 1, GLSetup
      
      CLOSE
      
      EXIT SUB
      
    END SELECT
    
  LOOP UNTIL Frm.KeyCode = ESC
  
  
END SUB

SUB GetTaxInfo
  
  SHARED Cnf AS Config          'System config & colors
  
  REDIM InvTaxRec(1)  AS InvoiceTaxInfoType
  TaxRecLen = LEN(InvTaxRec(1))
  
  OpenAcctIdx AcctIdxFileNum, NumIdxRecs
  CLOSE AcctIdxFileNum
  REDIM TempAcctIdx(1 TO NumIdxRecs) AS GLAcctIndexType
  ElSize = LEN(TempAcctIdx(1))
  FGetAH "GLACCT.IDX", TempAcctIdx(1), ElSize, NumIdxRecs
  
  REDIM Form$(0, 0)             'Holds data from form
  REDIM Fld(0) AS FieldInfo     'Field editing array
  REDIM Frm(1) AS FormInfo      'Dim the form info array
  
  TaxFile = FREEFILE
  OPEN "APINVTAX.DAT" FOR RANDOM AS TaxFile LEN = TaxRecLen
  IF LOF(TaxFile) > 0 THEN
    GET TaxFile, 1, InvTaxRec(1)
    GotTaxFile = True
  END IF
  CLOSE TaxFile
  
  FirstTime = True
  
  LibLoadDisplayForm "GL.QSL", "TAXINFO", Form$(), Fld()        'userinfo
  PrintHelp "Invoice Tax Accounts Setup"
  
  Action = 1    'Init form edit
  
  DO            'Edit the form
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    IF FirstTime THEN
      FirstTime = False
      IF GotTaxFile THEN
        LSET Form$(1, 0) = InvTaxRec(1).InvTax(1).AcctNo
        LSET Form$(2, 0) = STR$(InvTaxRec(1).InvTax(1).TaxAmt)
        LSET Form$(3, 0) = InvTaxRec(1).InvTax(2).AcctNo
        LSET Form$(4, 0) = STR$(InvTaxRec(1).InvTax(2).TaxAmt)
        LSET Form$(5, 0) = InvTaxRec(1).AUTODIST
      ELSE
        LSET Form$(2, 0) = "0.000"
        LSET Form$(4, 0) = "0.000"
      END IF
      Action = 1

    END IF
    
    IF Frm(1).FldNo <> Frm(1).PrevFld THEN
      SELECT CASE Frm(1).PrevFld
      CASE 1, 3
        GOSUB CheckForValidAcctNum
      END SELECT
    END IF
    SELECT CASE Frm(1).KeyCode  'Respond to the key presses
    CASE F10Key 'F10=Save
      GOSUB SaveInvTaxRec
      IF OKFlag THEN
        EXIT DO
      END IF
      
    END SELECT
    
  LOOP UNTIL Frm(1).KeyCode = ESC
  
  EXIT SUB
  
SaveInvTaxRec:
  
  OKFlag = True
  GOSUB CheckTaxInfo

  IF OKFlag THEN
    TaxFile = FREEFILE
    OPEN "APINVTAX.DAT" FOR RANDOM AS TaxFile LEN = TaxRecLen
    PUT TaxFile, 1, InvTaxRec(1)
    CLOSE TaxFile
  END IF

RETURN

CheckTaxInfo:

  REDIM TAcct$(1 TO 2)
  REDIM TaxAmt(1 TO 2) AS DOUBLE

  FOR Cnt = 2 TO 4 STEP 2      'test tax amt field for blank
    IF LEN(QPTrim$(Form$(Cnt, 0))) = 0 THEN
      OKFlag = False
      Ok = MsgBox("GL.QSL", "NOACCT")
      GOTO NoGood
    END IF
  NEXT

  TAcct$(1) = StripAcct(QPTrim$(Form$(1, 0)))
  TAcct$(2) = StripAcct(QPTrim$(Form$(3, 0)))
  TaxAmt(1) = Value#(Form$(2, 0), ErrCode%)
  TaxAmt(2) = Value#(Form$(4, 0), ErrCode%)
  
  FOR Cnt = 1 TO 2
    TAcct$(Cnt) = FmtAcct$(TAcct$(Cnt), GLSetup.FundLen, GLSetup.AcctLen, GLSetup.DetLen)
    IF LEN(QPTrim$(TAcct$(Cnt))) > 0 THEN
      IF FindAcct(TAcct$(Cnt)) <= 0 THEN
        OKFlag = False
        Ok = MsgBox("GL.QSL", "NOACCT")
        GOTO NoGood
      END IF
    END IF
  NEXT

  IF OKFlag THEN
    FOR Cnt = 1 TO 2
      InvTaxRec(1).InvTax(Cnt).AcctNo = TAcct$(Cnt)
      InvTaxRec(1).InvTax(Cnt).TaxAmt = TaxAmt#(Cnt)
    NEXT
    InvTaxRec(1).AUTODIST = QPTrim$(Form$(5, 0))
  END IF

NoGood:

RETURN
  
  
CheckForValidAcctNum:
  
  TotLen = GLSetup.FundLen + GLSetup.AcctLen + GLSetup.DetLen
  AcctFld = Frm(1).PrevFld
  MatchAccts = 0
  
  Lookup$ = QPTrim$(UCASE$(Form$(AcctFld, 0)))
  TAcct$ = StripAcct(Lookup$)
  
  Lookup$ = FmtAcct$(TAcct$, GLSetup.FundLen, GLSetup.AcctLen, GLSetup.DetLen)
  Lookuplen = LEN(Lookup$)
  
  REDIM ScrBuf(2000)
  
  SaveScrn ScrBuf()             'Save the existing screen
  
  'look here
  'The EDCheck4Acct function this gosub replaces, was loading the vendor index
  'file everytime it checked an account. Should help speed this up
  OpenAcctFile AcctFileNum, NumAccts
  REDIM AcctList(1 TO NumIdxRecs) AS FLen       'Dim array to hold trans list
  FOR Cnt = 1 TO NumIdxRecs     'Read the file
    IF LEFT$(TempAcctIdx(Cnt).AcctNum, Lookuplen) = Lookup$ THEN
      GET AcctFileNum, TempAcctIdx(Cnt).RecNum, Acct
      MatchAccts = MatchAccts + 1
      LSET AcctList(MatchAccts).V = List$
      AcctList(MatchAccts).V = TempAcctIdx(Cnt).AcctNum
      MID$(AcctList(MatchAccts).V, TotLen + 4) = Acct.Title
      MID$(AcctList(MatchAccts).V, 41) = MKI$(TempAcctIdx(Cnt).RecNum)
    END IF
  NEXT
  
  CLOSE AcctFileNum
  
  SELECT CASE MatchAccts
  CASE 0
    PLAY "L32O3EC"              'Acct NOT found
    Ok = MsgBox("GL.QSL", "NOACCT")
    EDCheck4Acct = False
  CASE 1
    Choice = 1
    GOSUB EGetAcctInfo
    EDCheck4Acct = True
  CASE ELSE
    Row = 14
    Col = 35
    LOCATE Row, Col
    Action = 0
    MaxLen = 39
    BoxBot = 23
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
        'AcctRecNum = CVI(RIGHT$(AcctList(Choice).V, 2))
        'LSET Form$(AcctFld, 0) = LEFT$(AcctList(Choice).V, TotLen + 4)
        GOSUB EGetAcctInfo
        EDCheck4Acct = True
        EXIT DO
      END IF
    LOOP
  END SELECT
  
  RestScrn ScrBuf()             'Restore Screen
  
  ERASE AcctList, ScrBuf
  
  SELECT CASE EDCheck4Acct
  CASE True
    SaveField AcctFld, Form$(), Fld(), BadFld   'Write the flag to the form$
    SaveField AcctFld + 1, Form$(), Fld(), BadFld               'Write the flag to the form$
    Action = 2
  CASE ELSE
    Frm(1).FldNo = AcctFld      'Put user back on acct field
    'LSET Form$(Frm(1).FldNo, 0) = ""
    'LSET Form$(Frm(1).FldNo + 1, 0) = ""
    Action = 2
  END SELECT
  
  RETURN
  
EGetAcctInfo:
  LSET Form$(AcctFld, 0) = LEFT$(AcctList(Choice).V, TotLen + 3)
  'OpenAcctFile AcctFileNum, NumAccts
  'GET AcctFileNum, AcctRecNum, Acct
  'CLOSE AcctFileNum
RETURN
  
  
  
  
END SUB

SUB OldReLink
  
  CLS
  LOCATE 1, 1, 0
  
  QPrintRC "Re-link Accounting Database", 12, 1, 15
  QPrintRC "Esc=Cancel, any other key to continue", 13, 1, 15
  k$ = INPUT$(1)
  IF k$ = CHR$(27) THEN EXIT SUB
  
  CLS
  'ON LOCAL ERROR GOTO GotErr
  
  DIM Trans AS GLTransRecType
  'DIM Acct AS GLAcctRecType
  
  OpenTransFile TransFileNum, NumTrans&
  OpenAcctFile AcctFileNum, NumAccts
  
  LOCK TransFileNum
  LOCK AcctFileNum
  
  LogFile = FREEFILE
  OPEN "LINK_ERR.LOG" FOR OUTPUT AS #LogFile
  
  FOR TCnt& = 1 TO NumTrans&
    Done# = (TCnt& / NumTrans&) * 100
    LOCATE 12, 1, 0
    PRINT USING "Resetting transaction file pointers. ###% complete."; Done#
    GET TransFileNum, TCnt&, Trans
    Trans.NextTran = 0
    PUT TransFileNum, TCnt&, Trans
  NEXT
  
  LOCATE 12, 1, 0
  PRINT STRING$(80, 32)
  
  FOR Cnt = 1 TO NumAccts
    Done# = (Cnt / NumAccts) * 100
    LOCATE 12, 1, 0
    PRINT USING "Resetting account file pointers. ###% complete."; Done#
    
    GET AcctFileNum, Cnt, Acct
    Acct.FrstTran = 0
    Acct.LastTran = 0
    PUT AcctFileNum, Cnt, Acct
  NEXT
  
  LOCATE 12, 1, 0
  PRINT STRING$(80, 32)
  
  FOR TCnt& = 1& TO NumTrans&
    
    LOCATE 12, 1, 0
    PRINT USING "Re-linking transaction: #####"; TCnt&
    
    GET TransFileNum, TCnt&, Trans
    
    RecordNum = FindAcct(Trans.AcctNum)
    
    IF RecordNum = 0 THEN
      BEEP
      BadTran = BadTran + 1
      LOCATE 13, 1
      PRINT USING "Orphaned transactions: #####"; BadTran
      GOSUB Logit
      GOTO JumpNext
    END IF
    
    GET AcctFileNum, RecordNum, Acct
    
    IF Acct.FrstTran = 0 THEN
      Acct.FrstTran = TCnt&
      Acct.LastTran = TCnt&
      PUT AcctFileNum, RecordNum, Acct
      GOTO JumpNext
    END IF
    
    Prev& = Acct.LastTran
    Acct.LastTran = TCnt&
    
    PUT AcctFileNum, RecordNum, Acct
    
    GET TransFileNum, Prev&, Trans
    Trans.NextTran = TCnt&
    PUT TransFileNum, Prev&, Trans
    
JumpNext:
  NEXT
  
  UNLOCK TransFileNum
  UNLOCK AcctFileNum
  CLOSE
  
  IF BadTran > 0 THEN
    QPrintRC "There were errors in transaction file.", 12, 1, 15
    QPrintRC "The text file LINKERR.LOG contains a list of orphaned transactions.", 13, 1, 15
    k$ = INPUT$(1)
    'SHELL "List link_err.log"
  ELSE
    KILL "LINK_ERR.LOG"
  END IF
  
  
  EXIT SUB
  
  
Logit:
  
  ToPrint$ = SPACE$(90)
  
  LSET ToPrint$ = Trans.AcctNum
  MID$(ToPrint$, 18) = Num2Date$(Trans.TrDate)
  MID$(ToPrint$, 30) = LEFT$(Trans.Desc, 15)
  MID$(ToPrint$, 50) = Trans.Ref
  MID$(ToPrint$, 60) = STR$(Trans.DrAmt)
  MID$(ToPrint$, 70) = STR$(Trans.CrAmt)
  MID$(ToPrint$, 80) = STR$(Cnt)
  
  PRINT #LogFile, ToPrint$
  
  RETURN
  
  
GotErr:
  SELECT CASE ERR
  CASE 70
    CLS
    QPrintRC "Unable to gain exclusive access to database.", 10, 1, 11
    QPrintRC "Wait until other users have exited system and", 11, 1, 11
    QPrintRC "try again.", 12, 1, 11
    EXIT SUB
  END SELECT
  
  RETURN
  
  
END SUB

SUB OpenSetupFile (SetupFileNum)
  
  '   SetupGJEdLen = LEN(GJEdit)
  '   GJEditFileNum = FREEFILE
  '   OPEN "gjed.dat" FOR RANDOM ACCESS READ WRITE SHARED AS GJEditFileNum LEN = GJEdLen
  '   NumEdTrans = LOF(GJEditFileNum) \ GJEdLen
  
  GLSetupRecLen = LEN(GLSetup)
  SetupFileNum = FREEFILE
  OPEN "GLSETUP.DAT" FOR RANDOM AS SetupFileNum LEN = GLSetupRecLen
  
END SUB

SUB oSetPostDates
  
  REDIM Form$(0, 0)             'Holds data from form
  REDIM Fld(0) AS FieldInfo     'Field editing array
  DIM Frm AS FormInfo           'Dim the form info array
  LibLoadDisplayForm "GL.QSL", "POSTDATE", Form$(), Fld()
  PrintHelp "Set Allowable Posting Dates"
  
  OpenSetupFile SetupFileNum
  
  GET SetupFileNum, 1, GLSetup
  
  Form$(1, 0) = Num2Date$(GLSetup.LPDate)
  Form$(2, 0) = Num2Date$(GLSetup.HPDate)
  CLOSE
  
  Action = 1    'Init form edit
  
  DO            'Edit the form
    
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    SELECT CASE Frm.KeyCode     'Respond to the key presses
    CASE F10Key 'F10=Save
      
      OpenSetupFile SetupFileNum
      GLSetup.LPDate = Date2Num(Form$(1, 0))
      GLSetup.HPDate = Date2Num(Form$(2, 0))
      PUT SetupFileNum, 1, GLSetup
      CLOSE
      ExitFlag = True
      
    CASE EscKey
      ExitFlag = True
      
    END SELECT
    
  LOOP UNTIL ExitFlag
  
  
END SUB

SUB SetFYDates
  
  SHARED Cnf AS Config          'System config & colors
  REDIM Form$(0, 0)             'Holds data from form
  REDIM Fld(0) AS FieldInfo     'Field editing array
  
  DIM Frm AS FormInfo           'Dim the form info array
  LibLoadDisplayForm "GL.QSL", "SETDATE", Form$(), Fld()
  PrintHelp "Set Accounting Periods"
  
  OpenSetupFile SetupFileNum
  
  GET SetupFileNum, 1, GLSetup
  
  Form$(1, 0) = Num2Date$(GLSetup.FYBeg)
  Form$(2, 0) = Num2Date$(GLSetup.FYEnd)
  Form$(3, 0) = Num2Date$(GLSetup.NYBeg)
  Form$(4, 0) = Num2Date$(GLSetup.NYEnd)
  'Form$(5, 0) = QPTrim$(STR$(GLSetup.ActiveYr))
  
  CLOSE
  
  Action = 1    'Init form edit
  
  DO            'Edit the form
    
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    SELECT CASE Frm.KeyCode     'Respond to the key presses
    CASE F10Key 'F10=Save
      
      OpenSetupFile SetupFileNum
      GLSetup.FYBeg = Date2Num(Form$(1, 0))
      GLSetup.FYEnd = Date2Num(Form$(2, 0))
      GLSetup.NYBeg = Date2Num(Form$(3, 0))
      GLSetup.NYEnd = Date2Num(Form$(4, 0))
      'GLSetup.ActiveYr = VAL(Form$(5, 0))
      PUT SetupFileNum, 1, GLSetup
      
      CLOSE
      
      EXIT SUB
      
    END SELECT
    
  LOOP UNTIL Frm.KeyCode = ESC
  
END SUB

SUB UnderConst
  
  LOCATE 1, 1, 0
  CLS
  QPrintRC "This is still under construction. Press any key to return to the menu.", 10, 1, 11
  Ok$ = INPUT$(1)
  
END SUB

SUB UserInfo
  
  SHARED Cnf AS Config          'System config & colors
  REDIM Form$(0, 0)             'Holds data from form
  REDIM Fld(0) AS FieldInfo     'Field editing array
  
  LibLoadDisplayForm "GL.QSL", "USERINF1", Form$(), Fld()       'userinfo
  
  DIM Frm AS FormInfo           'Dim the form info array
  
  OpenSetupFile SetupFileNum
  GET SetupFileNum, 1, GLSetup
  Form$(1, 0) = GLSetup.UserName
  Form$(2, 0) = GLSetup.CashAcct
  Form$(3, 0) = GLSetup.APAcct
  Form$(4, 0) = GLSetup.EncAcct
  Form$(5, 0) = GLSetup.FBAcct
  Form$(6, 0) = GLSetup.CDCash  'new
  Form$(7, 0) = GLSetup.CDDue
  Form$(8, 0) = GLSetup.CDActive
  Form$(9, 0) = GLSetup.CRCashAcct
  Form$(10, 0) = GLSetup.CDsbCash
  Form$(11, 0) = GLSetup.DeptCode
  
  
  CLOSE SetupFileNum
  
  Action = 1    'Init form edit
  
  DO            'Edit the form
    
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    SELECT CASE Frm.KeyCode     'Respond to the key presses
    CASE F10Key 'F10=Save
      OpenSetupFile SetupFileNum
      GLSetup.UserName = Form$(1, 0)
      GLSetup.CashAcct = Form$(2, 0)
      GLSetup.APAcct = Form$(3, 0)
      GLSetup.EncAcct = Form$(4, 0)
      GLSetup.FBAcct = Form$(5, 0)
      GLSetup.CDCash = Form$(6, 0)              'new
      GLSetup.CDDue = Form$(7, 0)
      GLSetup.CDActive = Form$(8, 0)
      GLSetup.CRCashAcct = Form$(9, 0)
      GLSetup.CDsbCash = Form$(10, 0)
      GLSetup.DeptCode = Form$(11, 0)
      PUT SetupFileNum, 1, GLSetup
      CLOSE
      EXIT SUB
      
    END SELECT
    
  LOOP UNTIL Frm.KeyCode = ESC
  
END SUB

