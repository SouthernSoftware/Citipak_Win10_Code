'GLRPT.BAS (c) Southern Software, Inc.
'by Nick Picerno, Jim Brown and Dale Watson

DEFINT A-Z
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
DECLARE SUB QPrint (X$, Colr%, Page%)
DECLARE SUB QPrintRC (t$, R%, c%, Clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
DECLARE SUB GetFundCodes (FirstFund$, LastFund$)
DECLARE SUB GLSumry ()
DECLARE SUB TrlBal ()
DECLARE SUB TranQuery ()
DECLARE SUB AcctHist ()
DECLARE SUB AQuery ()
DECLARE SUB OpenDeptFile (DeptFile%, NumDepts%)
DECLARE SUB BgtHist ()
DECLARE FUNCTION GetPct$ (N1#, N2#)
DECLARE FUNCTION GetDeptTitle$ (DeptRecNum%)
DECLARE FUNCTION FindDept% (DeptNum$)
DECLARE FUNCTION Num2Month% (Dt%)
DECLARE SUB BgtvsAct (User$, FundLen, AcctLen)
DECLARE SUB BgtvsActDept (User$, FundLen, AcctLen)
DECLARE SUB BalSheet ()
DECLARE SUB CashBalance ()
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
'DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB MnuTitle (Title$, Row, Col, Wdth, TxtColor)
DECLARE SUB ClrBG ()
DECLARE SUB GetFundList (FundList$(), NumFunds%)
DECLARE SUB GetFYDates (FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate)
DECLARE FUNCTION GetDeptOffsets$ ()
DECLARE FUNCTION InQtr% (TDate%, RDate%)
DECLARE FUNCTION Unique$ (Path$)
DECLARE SUB OpenFundIdx (FundIdxFileNum%, NumFIdxRecs%)
DECLARE SUB OpenFundFile (FundFileNum%, NumFunds%)

CONST FALSE = 0, True = NOT FALSE
  
'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'GL.bi'                        'GL File Types
'$INCLUDE: 'SCRAUX.BI'
'$INCLUDE: 'QSCR.BI'          'QuickScreen Declarations

'DIM SHARED TotLen%, FundLen%, AcctLen%, DetLen%
DIM SHARED Acct AS GLAcctRecType
DIM SHARED AcctIdx AS GLAcctIndexType
DIM SHARED Trans AS GLTransRecType
DIM SHARED FundIdx AS GLFundIndexType
DIM SHARED DeptIdx AS GLDeptIndexType


TYPE BgtVsActRecType
  Fund AS STRING * 3
  Bgt AS DOUBLE
  MTD AS DOUBLE
  YTD AS DOUBLE
END TYPE

'REDIM SHARED Choice$(0, 0)    'Used to hold multiple choices
  
'$INCLUDE: 'SetCnf.bi'        'Set monitor and color information
''$INCLUDE: 'init.bi'

SUB BgtvsAct (User$, FundLen, AcctLen)

'  ON LOCAL ERROR GOTO GotErr

  REDIM GLSetUpRec(1) AS GLSetupRecType
  RecLen = LEN(GLSetUpRec(1))
  GLFile = FREEFILE
  OPEN "glsetup.dat" FOR RANDOM AS GLFile LEN = RecLen
  GET #GLFile, , GLSetUpRec(1)
  CLOSE GLFile

  Town$ = QPTrim$(UCASE$(GLSetUpRec(1).UserName))
  IF INSTR(Town$, "HOLLY SPR") THEN
    HollyFlag = True
    Pitch12$ = CHR$(27) + CHR$(38) + CHR$(107) + CHR$(52) + CHR$(83)
  END IF

  SHARED Choice$()
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  REDIM Choice$(0 TO 3, 0 TO 2)
  DIM Frm AS FormInfo

  LibName$ = "GL.QSL"           'screen library
  FormName$ = "PRNBGACT"        'form to edit
  DevFld = 8   'Set to output device field number

  Choice$(0, 0) = STR$(DevFld)  'Choices for Output Device
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"
  Choice$(0, 1) = "4"
  Choice$(1, 1) = "Detail"
  Choice$(2, 1) = "Summary"
  Choice$(0, 2) = "5"
  Choice$(1, 2) = "1 - Bgt, Mo, YTD, Var :Req. Wide Carriage"
  Choice$(2, 2) = "2 - Bgt, Enc, YTD, Var :Req. Wide Carriage"
  Choice$(3, 2) = "3 - Bgt, QTD, YTD, Var :Req. Wide Carriage"
  
  GetFundCodes FirstFund$, LastFund$
  
  CommaFmt$ = "#########,.##"    'format takes 13 chars
  BgtFmt$ = "########,"         'format takes 11 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  BSumLine$ = STRING$(10, "-")  'summary line for budget columns
  PSumLine$ = " ----"            'summary line for Pct columns
  DivLine$ = STRING$(79, "-")   'dashed line
  DivLine2$ = STRING$(79, "=")  'Double Line
  FF$ = CHR$(12)
  REDIM Desc$(1)
  RptTitle$ = "Budget vs. Actual "
  MaxLines = 55

  '--Column offsets for printing amounts
  Col1 = 38    'Budget
  Col2 = 49    'Month or Actual YTD
  Col3 = 63    'YTD or Var
  Col4 = 76    'Pct
  Col5 = 89    'Enc
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()

  'GOTO Here
  
  Action = 1
  
  DO            'Allow user to input report params
    
    RunFlag = 0
    
    EditForm Form$(), Fld(), Frm, Cnf, Action
   ' GOTO Here
    
    IF Frm.PrevFld <> Frm.FldNo THEN            'User is moving around in form
      

      SELECT CASE Frm.PrevFld
        
      CASE 1    'Ending Date
        EndDate = Date2Num(Form$(1, 0))
        IF EndDate < -32000 THEN
          Ok = MsgBox("GL.QSL", "BADDATE")
          Frm.FldNo = 1
        END IF
        
      CASE 2, 3 'Fund Field
        Fld = Frm.PrevFld
        IF Frm.FldEdited THEN   'If the field was edited
          Form$(Fld, 0) = QPTrim$(Form$(Fld, 0))
          SaveField Fld, Form$(), Fld(), BadFld
          
          IF LEN(Form$(Fld, 0)) <> FundLen THEN
            Form$(Fld, 0) = ""
            Frm.FldNo = Fld
            
          ELSE
            GoodFund = FindFund(Form$(Fld, 0))
            IF GoodFund = 0 THEN
              Ok = MsgBox("GL.QSL", "NOFUND")
              Form$(Fld, 0) = ""
              Frm.FldNo = Fld
              BadFld = -1
            END IF
          END IF
        END IF
      CASE ELSE
        'what else?
      END SELECT
      
      Action = 1
      
    END IF
    
    SELECT CASE Frm.KeyCode     'Check the user's key presses
      
    CASE 42  'Asterisk
      Form$(1, 0) = DATE$
      Form$(2, 0) = FirstFund$
      Form$(3, 0) = LastFund$
      Form$(4, 0) = "D"
      Form$(5, 0) = "1"
      Form$(6, 0) = "N"
      Form$(7, 0) = "N"
      Form$(8, 0) = "Screen"
      Frm.FldNo = 9
      Action = 1
      
    CASE F10Key 'SaveButton

      'check date
      EndDate = Date2Num(Form$(1, 0))
      IF EndDate < -32000 THEN
        Ok = MsgBox("GL.QSL", "BADDATE")
        BadFld = 1
      END IF
      IF BadFld THEN
        Frm.FldNo = 1
        Action = 1
        GOTO DoAgain
      ELSE
        RunFlag = -1
      END IF
      
      'check for valid funds
      FOR FieldNum = 2 TO 3
        IF LEN(QPTrim$(Form$(FieldNum, 0))) = 0 THEN
          Ok = MsgBox("GL.QSL", "NOBLANK")
          Frm.FldNo = FieldNum
          BadFld = FieldNum
          EXIT FOR
        END IF
      NEXT
      IF BadFld THEN
        Action = 1
        RunFlag = 0
        GOTO DoAgain
      ELSE
        RunFlag = -1
      END IF
      
      'make sure funds are in ascending order
      StartFund$ = QPTrim$(Form$(2, 0))
      EndFund$ = QPTrim$(Form$(3, 0))
      IF StartFund$ > EndFund$ THEN
        Ok = MsgBox("GL.QSL", "RANGERR")
        BadFld = 1
      END IF
      IF BadFld THEN
        RunFlag = 0
        Frm.FldNo = 2
        Action = 1
        GOTO DoAgain
      ELSE
        RunFlag = -1
      END IF
  
      IF LEN(LTRIM$(RTRIM$(Form$(DevFld, 0)))) = 0 THEN
        Dev$ = "S"
        LPTNo = 1
      ELSE
        Dev$ = LEFT$(Form$(DevFld, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(DevFld, 0)), 1))
      END IF
  
    CASE ESC
      EXIT SUB
      
    END SELECT
    
DoAgain:
    BadFld = 0
  LOOP UNTIL RunFlag

Here:
  'Form$(1, 0) = DATE$
  'Form$(2, 0) = FirstFund$
  'Form$(3, 0) = LastFund$
  'Form$(4, 0) = "D"
  'Form$(5, 0) = "2"
  'Form$(6, 0) = "N"
  'Form$(7, 0) = "N"
  'Form$(8, 0) = "Screen"
  'RunFlag = -1

  LOCATE , , 0
  
  IF Form$(4, 0) = "D" THEN DoingDetail = True

  IF Form$(7, 0) = "Y" THEN SubTotalRevenues = True

  '  SubTotalRevenues = True
  'ELSE
  '  SubTotalRevenues = False
  'END IF

  IF DoingDetail THEN
    '--if break on depts
    IF Form$(6, 0) = "Y" THEN
      DeptOnNewPage = True
    END IF
  END IF

  WhichReport = VAL(Form$(5, 0))
  GetMonth = FALSE
  GetQtr = FALSE
  Desc$(1) = "Description                             Budget        Encumb           YTD      Variance  Pct"

  MTDBal# = 0

  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate

  '--Determine fiscal year
  IF EndDate >= FY2BegDate THEN
    ActiveYear = 2
    FYStartDate = FY2BegDate
  ELSE
    ActiveYear = 1
    FYStartDate = FY1BegDate
  END IF

  ReportFile$ = "BGTACT.PRN"

  'ReportFile$ = Unique$(Path$)
  'P17$ = CHR$(27) + "(s17H"

  PrnFile = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile

  'PRINT #PrnFile, P17$;

  IF HollyFlag THEN
    PRINT #PrnFile, Pitch12$;
  END IF

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFile, NumFunds
  
  REDIM RevAccts%(1 TO NumGLAccts)              'Holds all rev acct record nums
  REDIM ExpAccts%(1 TO NumGLAccts)              'Holds all exp acct record nums
  REDIM FundList$(1)                            'List of all active Funds
  GetFundList FundList$(), NumFunds

  PrintHelp "Processing:"

    PRINT #PrnFile, " Acct Num"; TAB(20); "Budget"; TAB(40); "BAL"; TAB(60); " Encumb"
    FOR A = 1 TO NumGLAccts
      GET AcctIdxFileNum, A, AcctIdx
      GET AcctFileNum, AcctIdx.RecNum, Acct

      IF Acct.Deleted = 0 THEN
        PCnt = PCnt + 1
        PRINT #PrnFile, Acct.Num;
        PRINT #PrnFile, USING "########.##"; TAB(20); Acct.Bgt; TAB(40); Acct.Bal; TAB(60); Acct.Encumb
      'ELSE
      '  STOP
      END IF
    NEXT
    PRINT #PrnFile, "Active Acccounts:"; PCnt
    CLOSE
    GOTO ShowMe

  '--Build a list of revenue and exp accounts
  FOR GLAcct = 1 TO NumGLAccts
    '--Initialize
    BGTBal# = 0
    YTDBal# = 0
    MTDBal# = 0
    
    GET AcctIdxFileNum, GLAcct, AcctIdx
    GET AcctFileNum, AcctIdx.RecNum, Acct
    
    '--Find what fund this account is in
    FundCode$ = LEFT$(Acct.Num, FundLen)
    '--See if the account is in a fund we want to see
    IF FundCode$ >= StartFund$ AND FundCode$ <= EndFund$ THEN
      '--Account is in fund, check to see if its proper type
      '--We want only revenue or expenditure accounts
      IF Acct.Typ = "R" OR Acct.Typ = "E" THEN
        
        '--Print the account on screen for user to see what's going on
        QPrintRC Acct.Num, 25, 14, -1
        
        '--Assign the Account Record Number to proper list
        SELECT CASE Acct.Typ
        CASE "E"
          ECnt = ECnt + 1
          ExpAccts%(ECnt) = AcctIdx.RecNum
          
        CASE "R"
          RCnt = RCnt + 1
          RevAccts%(RCnt) = AcctIdx.RecNum
        END SELECT
        
        '--Get account balances
        '--There should be no beginning balances in rev & exp accts
        'YTDBal# = Round#(Acct.BegBal)           'get the beginning balance

        NextTr& = Acct.FrstTran 'get the first trans for this acct
        
        DO UNTIL NextTr& = 0    'keep going 'til we run out
          
          GET TransFileNum, NextTr&, Trans
          
          '--Get MTD Account Balance if necessary

          IF Trans.TrDate >= FYStartDate AND Trans.TrDate <= EndDate THEN
          
          IF GetMonth THEN
            'Lookhere change num2month to reflect year & month
            D$ = Num2Date(Trans.TrDate)
            TransMonth = Num2Month%(Trans.TrDate)
            IF TransMonth = RptMonth THEN
              SELECT CASE Acct.Typ
              CASE "E"
                MTDBal# = MTDBal# + Round#(Trans.DrAmt) - Round#(Trans.CrAmt)
              CASE "R"
                MTDBal# = MTDBal# + Round#(Trans.CrAmt) - Round#(Trans.DrAmt)
              END SELECT
            END IF
          END IF
         
          IF GetQtr THEN
            InThisQtr = InQtr(Trans.TrDate, EndDate)
            IF InThisQtr THEN
              SELECT CASE Acct.Typ
                CASE "E"
                  MTDBal# = MTDBal# + Round#(Trans.DrAmt) - Round#(Trans.CrAmt)
                CASE "R"
                  MTDBal# = MTDBal# + Round#(Trans.CrAmt) - Round#(Trans.DrAmt)
              END SELECT
            END IF
          END IF
          
          '--Get YTD Account Balance
          '--original
          '--Carthage 9/17/96
          'IF Trans.TrDate <= EndDate THEN
          '--08/07/96 keeping funds open --
          '--Does'nt work when 2 years are open...what to do?
          'IF Trans.TrDate <= EndDate THEN
            SELECT CASE Acct.Typ
            CASE "E"
              YTDBal# = YTDBal# + Round#(Trans.DrAmt) - Round#(Trans.CrAmt)
            CASE "R"
              YTDBal# = YTDBal# + Round#(Trans.CrAmt) - Round#(Trans.DrAmt)
            END SELECT
          END IF
          
          NextTr& = Trans.NextTran              'Get the next transaction
          
        LOOP
        
        '--Put the new totals in the file
        Acct.MTD = Round#(MTDBal#)
        Acct.YTD = Round#(YTDBal#)
        PUT AcctFileNum, AcctIdx.RecNum, Acct
        
      END IF    '--test for rev or exp accts
    END IF      '--End of acct in fund range test
  NEXT          'Process next account
  
  '--Now write the report to file.
  PrintHelp "Generating Report..."
  
  FirstTime = True 'flag not to print page feed first time thru

  '--Process each fund
  FOR Fund = 1 TO NumFunds
    ThisFund$ = FundList$(Fund)
    IF ThisFund$ >= StartFund$ AND ThisFund$ <= EndFund$ THEN
      UsingFund = True
      FundRec = FindFund(ThisFund$)             'Get the fund name
      FundName$ = QPTrim$(GetFundTitle(FundRec))
      Dept$ = ""
      LastDept$ = ""
      LastDeptName$ = ""
    ELSE
      UsingFund = FALSE
    END IF
    
    IF UsingFund THEN
      '--print a form feed for each new fund
      IF FirstTime THEN
        FirstTime = FALSE
      ELSE
        PRINT #PrnFile, FF$
      END IF

      'RE$ = "Revenues"
      IF NOT SubTotalRevenues OR NOT DoingDetail THEN
        GOSUB PrintBVAPageHeader
      END IF

      '--Search thru list of revenue accounts
      FOR Cnt = 1 TO RCnt
        Rec = RevAccts%(Cnt)
        GET AcctFileNum, Rec, Acct
        FundCode$ = LEFT$(Acct.Num, FundLen)
        
        IF FundCode$ = ThisFund$ THEN
          Account$ = QPTrim$(Acct.Num) + "  " + QPTrim$(Acct.Title)
          SELECT CASE ActiveYear
          CASE 1
            BudgetAmt# = Acct.Bgt
          CASE 2
            BudgetAmt# = Acct.NYApp
          END SELECT

          '----------
          'IF SubTotalRevenues THEN
          '  '--Extract the Dept$ from the G/L Acct
          '  Dept$ = MID$(Acct.Num, FundLen + 2, AcctLen)
          '
          '  '--if a new dept, get its name from the Dept name file
          '  IF Dept$ <> LastDept$ THEN
          '    DeptRecNum = FindDept(Dept$)
          '    IF DeptRecNum > 0 THEN
          '      DeptName$ = QPTrim$(GetDeptTitle$(DeptRecNum))
          '    ELSE
          '      'DeptName$ = "Department " + Dept$
          '      DeptName$ = " "
          '    END IF
          '  END IF
          '
          '  '--Print Department Header first time thru
          '  IF LEN(LastDeptName$) = 0 THEN
          '    '--if we're not printing departments on separate pages then
          '    '--print a new page header
          '    IF DeptOnNewPage = FALSE THEN
          '      IF DoingDetail = True THEN
          '        PRINT #PrnFile, FF$
          '        GOSUB PrintBVAPageHeader
          '      ELSE
          '        PRINT #PrnFile,
          '        PRINT #PrnFile, "Revenues"
          '      END IF
          '    END IF
          '    LastDeptName$ = DeptName$
          '    LastDept$ = Dept$
          '    IF DoingDetail THEN
          '      GOSUB PrintDeptHeader
          '    END IF
          '  END IF
          '
          '  '--see if we need to subtotal dept
          '  IF LEN(LastDept$) > 0 THEN
          '    IF Dept$ <> LastDept$ THEN
          '      GOSUB PrintDeptTotals
          '      IF DoingDetail THEN
          '        GOSUB PrintDeptHeader
          '      END IF
          '    END IF
          '
          '  END IF
          '
          'END IF
          '===========
          
          Pct$ = GetPct$(Acct.YTD, BudgetAmt#)
          
          IF DoingDetail THEN
            Variance# = Acct.YTD - BudgetAmt# 'Acct.Bgt
            ToPrint$ = SPACE$(96)
            LSET ToPrint$ = LEFT$(Account$, 36)
            MID$(ToPrint$, Col1) = FUsing$(STR$(BudgetAmt#), BgtFmt$)   'changed
            'SELECT CASE WhichReport
            'CASE 1, 3
            '  MID$(ToPrint$, Col2) = FUsing$(STR$(Acct.MTD), CommaFmt$)
            '  MID$(ToPrint$, Col3) = FUsing$(STR$(Acct.YTD), CommaFmt$)
            '  MID$(ToPrint$, Col4) = FUsing$(STR$(Variance#), CommaFmt$)
            '  MID$(ToPrint$, Col5) = Pct$
            'CASE 2
              MID$(ToPrint$, Col2) = FUsing$(STR$(Acct.Encumb), CommaFmt$)
              MID$(ToPrint$, Col3) = FUsing$(STR$(Acct.YTD), CommaFmt$)
              MID$(ToPrint$, Col4) = FUsing$(STR$(Variance#), CommaFmt$)
              MID$(ToPrint$, Col5) = Pct$
            'END SELECT

            PRINT #PrnFile, RTRIM$(ToPrint$)

            LineCnt = LineCnt + 1
            IF LineCnt > MaxLines THEN
              PRINT #PrnFile, FF$
              GOSUB PrintBVAPageHeader
            END IF
          END IF
          
          'changed
          IF GetMonth OR GetQtr THEN
            MTDSum# = MTDSum# + Acct.MTD
          END IF
          
          BgtSum# = BgtSum# + BudgetAmt# 'Acct.Bgt
          YTDSum# = YTDSum# + Acct.YTD
          
          DeptBgtSum# = DeptBgtSum# + BudgetAmt#
          DeptYTDSum# = DeptYTDSum# + Acct.YTD
          DeptEncSum# = DeptEncSum# + Acct.Encumb
          DeptMTDSum# = DeptMTDSum# + Acct.MTD

          LastDept$ = Dept$
          LastDeptName$ = DeptName$
          
        END IF
        
      NEXT      'Revenue Acct
      
      '--Summarize Revenues
      'IF DoingDetail THEN
        GOSUB PrintSummaryLines
      'END IF
      
      ToPrint$ = SPACE$(96)
      LSET ToPrint$ = "Total Revenues"
      Pct$ = GetPct$(YTDSum#, BgtSum#)
      Variance# = YTDSum# - BgtSum#
      'SELECT CASE WhichReport
      'CASE 1, 3
      '  Variance# = YTDSum# - BgtSum#
      '  MID$(ToPrint$, Col1) = FUsing$(STR$(BgtSum#), BgtFmt$)
      '  MID$(ToPrint$, Col2) = FUsing$(STR$(MTDSum#), TotalFmt$)
      '  MID$(ToPrint$, Col3) = FUsing$(STR$(YTDSum#), TotalFmt$)
      '  MID$(ToPrint$, Col4) = FUsing$(STR$(Variance#), CommaFmt$)
      '  MID$(ToPrint$, Col5) = Pct$
      '  '--Reset vars
      '  FundRevMTD# = MTDSum#
      '  MTDSum# = 0
      'CASE 2
        Variance# = YTDSum# - BgtSum#
        MID$(ToPrint$, Col1) = FUsing$(STR$(BgtSum#), BgtFmt$)
        MID$(ToPrint$, Col2) = FUsing$(" 0", TotalFmt$)
        MID$(ToPrint$, Col3) = FUsing$(STR$(YTDSum#), TotalFmt$)
        MID$(ToPrint$, Col4) = FUsing$(STR$(Variance#), CommaFmt$)
        MID$(ToPrint$, Col5) = Pct$
      'END SELECT
      PRINT #PrnFile, RTRIM$(ToPrint$)
      LineCnt = LineCnt + 1
      IF LineCnt > MaxLines THEN
        PRINT #PrnFile, FF$
        GOSUB PrintBVAPageHeader
      END IF
      
      FundRevBgt# = BgtSum#
      FundRevYTD# = YTDSum#
      EncSum# = 0
      BgtSum# = 0
      YTDSum# = 0
      
      'initialize dept variables
      DeptBgtSum# = 0
      DeptMTDSum# = 0
      DeptYTDSum# = 0
      DeptEncSum# = 0
      LastDept$ = ""
      LastDeptName$ = ""
      
      '--Search exp accounts list for accounts in this fund
      'RE$ = "Expenditures"
      FOR Cnt = 1 TO ECnt
        Rec = ExpAccts%(Cnt)
        GET AcctFileNum, Rec, Acct
        
        FundCode$ = LEFT$(Acct.Num, FundLen)
        IF FundCode$ = ThisFund$ THEN
          Account$ = QPTrim$(Acct.Num) + "  " + QPTrim$(Acct.Title)
          
          SELECT CASE ActiveYear
          CASE 1
            BudgetAmt# = Acct.Bgt
          CASE 2
            BudgetAmt# = Acct.NYApp
          END SELECT
         
          '--Extract the Dept$ from the G/L Acct
          Dept$ = MID$(Acct.Num, FundLen + 2, AcctLen)
          
          '--if a new dept, get its name from the Dept name file
          IF Dept$ <> LastDept$ THEN
            DeptRecNum = FindDept(Dept$)
            IF DeptRecNum > 0 THEN
              DeptName$ = QPTrim$(GetDeptTitle$(DeptRecNum))
            ELSE
              'DeptName$ = "Department " + Dept$
              DeptName$ = " "
            END IF
          END IF

          '--Print Department Header first time thru
          IF LEN(LastDeptName$) = 0 THEN
            '--if we're not printing departments on separate pages then
            '--print a new page header
            IF DeptOnNewPage = FALSE THEN
              IF DoingDetail = True THEN
                PRINT #PrnFile, FF$
                GOSUB PrintBVAPageHeader
              ELSE
                PRINT #PrnFile,
                PRINT #PrnFile, "Expenditures"
              END IF
            END IF
            LastDeptName$ = DeptName$
            LastDept$ = Dept$
            IF DoingDetail THEN
              GOSUB PrintDeptHeader
            END IF
          END IF
          
          '--see if we need to subtotal dept
          IF LEN(LastDept$) > 0 THEN
            IF Dept$ <> LastDept$ THEN
              GOSUB PrintDeptTotals
              IF DoingDetail THEN
                GOSUB PrintDeptHeader
              END IF
            END IF
          END IF
          
          IF DoingDetail THEN   'Print Account Detail
            ToPrint$ = SPACE$(96)

            LSET ToPrint$ = LEFT$(Account$, 36)

            'SELECT CASE WhichReport
            'CASE 1, 3
            '  Pct$ = GetPct$(Acct.YTD, BudgetAmt#) 'Acct.Bgt
            '  Variance# = BudgetAmt# - Acct.YTD
            '  MID$(ToPrint$, Col1) = FUsing$(STR$(BudgetAmt#), BgtFmt$)
            '  MID$(ToPrint$, Col2) = FUsing$(STR$(Acct.MTD), CommaFmt$)
            '  MID$(ToPrint$, Col3) = FUsing$(STR$(Acct.YTD), CommaFmt$)
            '  MID$(ToPrint$, Col4) = FUsing$(STR$(Variance#), CommaFmt$)
            '  MID$(ToPrint$, Col5) = Pct$
            'CASE 2
              Pct$ = GetPct$(Acct.Encumb + Acct.YTD, BudgetAmt#) 'Acct.Bgt
              Variance# = BudgetAmt# - Acct.Encumb - Acct.YTD
              MID$(ToPrint$, Col1) = FUsing$(STR$(BudgetAmt#), BgtFmt$)
              MID$(ToPrint$, Col2) = FUsing$(STR$(Acct.Encumb), CommaFmt$)
              MID$(ToPrint$, Col3) = FUsing$(STR$(Acct.YTD), CommaFmt$)
              MID$(ToPrint$, Col4) = FUsing$(STR$(Variance#), CommaFmt$)
              MID$(ToPrint$, Col5) = Pct$
            'END SELECT
            PRINT #PrnFile, RTRIM$(ToPrint$)
            LineCnt = LineCnt + 1
            IF LineCnt > MaxLines THEN
              PRINT #PrnFile, FF$
              GOSUB PrintBVAPageHeader
            END IF
          END IF
          
          IF GetMonth OR GetQtr THEN
            MTDSum# = MTDSum# + Acct.MTD
            DeptMTDSum# = DeptMTDSum# + Acct.MTD
          END IF
          
          BgtSum# = BgtSum# + BudgetAmt#
          YTDSum# = YTDSum# + Acct.YTD
          EncSum# = EncSum# + Acct.Encumb

          DeptBgtSum# = DeptBgtSum# + BudgetAmt#
          DeptYTDSum# = DeptYTDSum# + Acct.YTD
          DeptEncSum# = DeptEncSum# + Acct.Encumb

          LastDept$ = Dept$
          LastDeptName$ = DeptName$
        
        END IF
      NEXT      'Exp Acct
      
      '--Summarize last Dept after loop
      GOSUB PrintDeptTotals
      
      '--Now summarize all expenditures
      GOSUB PrintSummaryLines   'Print dashed line after last

      '--print total exp for fund
      IF DeptOnNewPage THEN
        PRINT #PrnFile, FF$
        GOSUB PrintBVAPageHeader
      END IF

      Pct$ = GetPct$(YTDSum#, BgtSum#)
      ToPrint$ = SPACE$(96)
      LSET ToPrint$ = "Total Expenditures for Fund:"
      'SELECT CASE WhichReport
      'CASE 1, 3
      '  Variance# = BgtSum# - YTDSum#
      '  MID$(ToPrint$, Col1) = FUsing$(STR$(BgtSum#), BgtFmt$)
      '  MID$(ToPrint$, Col2) = FUsing$(STR$(MTDSum#), CommaFmt$)
      '  MID$(ToPrint$, Col3) = FUsing$(STR$(YTDSum#), CommaFmt$)
      '  MID$(ToPrint$, Col4) = FUsing$(STR$(Variance#), CommaFmt$)
      '  MID$(ToPrint$, Col5) = Pct$
      '  FundExpMTD# = MTDSum#
      '  MTDSum# = 0
      'CASE 2
        Variance# = BgtSum# - EncSum# - YTDSum#
        MID$(ToPrint$, Col1) = FUsing$(STR$(BgtSum#), BgtFmt$)
        MID$(ToPrint$, Col2) = FUsing$(STR$(EncSum#), CommaFmt$)
        MID$(ToPrint$, Col3) = FUsing$(STR$(YTDSum#), CommaFmt$)
        MID$(ToPrint$, Col4) = FUsing$(STR$(Variance#), CommaFmt$)
        MID$(ToPrint$, Col5) = Pct$
      'END SELECT
      PRINT #PrnFile, RTRIM$(ToPrint$)
      
      '--Summarize Exp
      FundExpBgt# = BgtSum#
      FundExpYTD# = YTDSum#
      FundEncYTD# = EncSum#

      BgtSum# = 0
      YTDSum# = 0
      EncSum# = 0

      '--Summarize fund
      IF GetMonth OR GetQtr THEN 'changed
        MTDBal# = Round#(FundRevMTD# - FundExpMTD#)
      END IF
      
      BGTBal# = Round#(FundRevBgt# - FundExpBgt#)
      YTDBal# = Round#(FundRevYTD# - FundExpYTD#)
      EncBal# = Round#(FundEncYTD#)

      PRINT #PrnFile,
      
      '--print the net
      ToPrint$ = SPACE$(96)
      LSET ToPrint$ = "Revenues Over/(Under) Expenditures"
      MID$(ToPrint$, Col1) = FUsing$(STR$(BGTBal#), BgtFmt$)
      'SELECT CASE WhichReport
      'CASE 1, 3
      '  MID$(ToPrint$, Col2) = FUsing$(STR$(MTDBal#), CommaFmt$)
      '  MID$(ToPrint$, Col3) = FUsing$(STR$(YTDBal#), CommaFmt$)
      '  '--Reset MTD Variables
      '  FundRevMTD# = 0
      '  FundExpMTD# = 0
      '  DeptMTDSum# = 0
      'CASE 2
        MID$(ToPrint$, Col3) = FUsing$(STR$(YTDBal#), CommaFmt$)
      'END SELECT

      'FOR zz = 1 TO 96

      PRINT #PrnFile, RTRIM$(ToPrint$)
      
      '--Blank lines between funds
      PRINT #PrnFile,
      PRINT #PrnFile,
      LineCnt = LineCnt + 2

      '--Reset variables for next fund
      FundRevBgt# = 0
      FundRevYTD# = 0
      
      FundExpBgt# = 0
      FundExpYTD# = 0
      FundEncYTD# = 0

      DeptBgtSum# = 0
      DeptYTDSum# = 0
      
    END IF      'Using fund test
  NEXT          'process next fund
  
  PRINT #PrnFile, FF$
  CLOSE
  
  '====End Report Processing
  
  SELECT CASE Dev$
    CASE "S"
      EntryPoint = 2
    CASE "P"
      EntryPoint = 5
  END SELECT

ShowMe:
  PrintRptFile RptTitle$, ReportFile$, LPTNo, RetCode%, EntryPoint
  CLOSE

  'KILL ReportFile$
  'End Report Printing========================================================
  
EXIT SUB
  

PrintBVAPageHeader:
  PRINT #PrnFile, User$; TAB(43); "Run Date: " + DATE$
  PRINT #PrnFile, FundName$ + " " + RptTitle$
  PRINT #PrnFile, "Period Ending: " + Num2Date$(EndDate)
  PRINT #PrnFile,
  PRINT #PrnFile, Desc$(1)
  PRINT #PrnFile, STRING$(LEN(Desc$(1)), "-")
  LineCnt = 6
RETURN

  
PrintSummaryLines:
    '--Print summary lines
    ToPrint$ = SPACE$(96)
    MID$(ToPrint$, Col1) = BSumLine$
    MID$(ToPrint$, Col2) = SumLine$
    MID$(ToPrint$, Col3) = SumLine$
    MID$(ToPrint$, Col4) = SumLine$
    MID$(ToPrint$, Col5) = PSumLine$
    PRINT #PrnFile, RTRIM$(ToPrint$)
    LineCnt = LineCnt + 1
RETURN
  
  
PrintDeptTotals:
  IF DoingDetail THEN
    GOSUB PrintSummaryLines
    DeptSummary$ = LastDeptName$ + " Totals"
  ELSE
    DeptSummary$ = LastDept$ + " " + LastDeptName$
  END IF

  'IF INSTR(DeptSummary$, "4910") > 0 THEN STOP

  'IF INSTR(DeptSummary$, "ZONING") > 0 THEN STOP

  ToPrint$ = SPACE$(96)
  LSET ToPrint$ = DeptSummary$
'  SELECT CASE WhichReport
'     CASE 1, 3
'       Pct$ = GetPct$(DeptYTDSum#, DeptBgtSum#)
'       Variance# = DeptBgtSum# - DeptYTDSum#
'       MID$(ToPrint$, Col1) = FUsing$(STR$(DeptBgtSum#), BgtFmt$)
'       MID$(ToPrint$, Col2) = FUsing$(STR$(DeptMTDSum#), CommaFmt$)
'       MID$(ToPrint$, Col3) = FUsing$(STR$(DeptYTDSum#), CommaFmt$)
'       MID$(ToPrint$, Col4) = FUsing$(STR$(Variance#), TotalFmt$)
'       MID$(ToPrint$, Col5) = Pct$
'       DeptMTDSum# = 0
'     CASE 2
       Pct$ = GetPct$(DeptYTDSum# + DeptEncSum#, DeptBgtSum#)
       Variance# = DeptBgtSum# - DeptEncSum# - DeptYTDSum#
       MID$(ToPrint$, Col1) = FUsing$(STR$(DeptBgtSum#), BgtFmt$)
       MID$(ToPrint$, Col2) = FUsing$(STR$(DeptEncSum#), CommaFmt$)
       MID$(ToPrint$, Col3) = FUsing$(STR$(DeptYTDSum#), CommaFmt$)
       MID$(ToPrint$, Col4) = FUsing$(STR$(Variance#), TotalFmt$)
       MID$(ToPrint$, Col5) = Pct$
'     CASE ELSE
'  END SELECT
  PRINT #PrnFile, RTRIM$(ToPrint$)
  LineCnt = LineCnt + 1
  IF DoingDetail THEN
    '--formfeed if were at max
    IF LineCnt >= MaxLines THEN
      PRINT #PrnFile, FF$
      GOSUB PrintBVAPageHeader
    ELSE
      '--print a blank line after totals
      PRINT #PrnFile,
      LineCnt = LineCnt + 1
    END IF
  END IF
  DeptBgtSum# = 0
  DeptYTDSum# = 0
  DeptEncSum# = 0
RETURN
  
  
PrintDeptHeader:
  IF DeptOnNewPage THEN
    PRINT #PrnFile, FF$
    GOSUB PrintBVAPageHeader
  END IF
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = DeptName$
  PRINT #PrnFile, RTRIM$(ToPrint$)
  LineCnt = LineCnt + 1
RETURN


GotErr:
  ErrorCode$ = STR$(ERR)
  SELECT CASE ERR
    CASE 70
      CLS
      QPrintRC "Access Denied. Try again later.", 12, 1, 12
      QPrintRC "Press any key to continue.", 14, 1, 11
    CASE ELSE
      CLS
      QPrintRC "An Error has halted the system, Error Code: " + ErrorCode$, 12, 1, 14
      QPrintRC "Press any key exit.", 13, 1, 14
   END SELECT

   K$ = INPUT$(1)
   EXIT SUB
   
RETURN
  
  
END SUB

FUNCTION FindFund (FundNum$)
'*****************************************************************************
'Searches the acct index for a matching fund number and returns the record
'number of the fund
'
'    Input: AcctNum$ as a formatted G/L account number string
'  Returns: Record number of the account
'*****************************************************************************
'
   OpenFundIdx FundIdxFileNum, NumFIdxRecs

   IF NumFIdxRecs = 0 THEN
      RecordNum = 0
      CLOSE FundIdxFileNum
      EXIT FUNCTION
   END IF

   Match = FALSE
   FirstRec = 1
   LastRec = NumFIdxRecs
   Lookfor$ = FundNum$ 'LTRIM$(RTRIM$(FundNum$))

   DO UNTIL LastRec < FirstRec

      MiddleRec = (LastRec + FirstRec) \ 2

      GET FundIdxFileNum, MiddleRec, FundIdx

      IF Lookfor$ = RTRIM$(FundIdx.FundNum) THEN
         Match = -1
         EXIT DO
      ELSEIF Lookfor$ < FundIdx.FundNum THEN
         LastRec = MiddleRec - 1
      ELSE
         FirstRec = MiddleRec + 1
      END IF

   LOOP

   IF Match THEN
      RecNum = FundIdx.RecNum
   ELSE
      RecNum = 0
   END IF

   FindFund = RecNum

   CLOSE FundIdxFileNum


END FUNCTION

FUNCTION GetDeptOffsets$
'---------------------------------------------------
'Gets the Dept Code from the setup file.  The
'code determines where to look for the dept code in
'the acct code
'---------------------------------------------------
   DIM GLSetup AS GLSetupRecType
   SetUpRecLen = LEN(GLSetup)

   SetupFile = FREEFILE
   OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetupFile LEN = SetUpRecLen

   GET SetupFile, 1, GLSetup

   GetDeptOffsets$ = GLSetup.DeptCode

   CLOSE SetupFile


END FUNCTION

SUB GetFundCodes (FirstFund$, LastFund$)
   
   OpenFundIdx FundIdxFileNum, NumFIdxRecs

   IF NumFIdxRecs = 0 THEN
      FirstFund = 0
      LastFund = 0
      EXIT SUB
   END IF
   
   GET FundIdxFileNum, 1, FundIdx
   FirstFund$ = QPTrim$(FundIdx.FundNum)

   GET FundIdxFileNum, NumFIdxRecs, FundIdx
   LastFund$ = QPTrim$(FundIdx.FundNum)
   
   CLOSE FundIdxFileNum
   
END SUB

'****************************************************************************
'Retrieves the fund title from the fund data file.
'****************************************************************************
'
FUNCTION GetFundTitle$ (FundRecNum)

   DIM FundRec AS GLFundRecType
   OpenFundFile FundFileNum, NumFunds
   GET FundFileNum, FundRecNum, FundRec
   GetFundTitle$ = FundRec.Title
   CLOSE FundFileNum

END FUNCTION

FUNCTION GetPct$ (N1#, N2#)
  
  Pct$ = SPACE$(5)
  IF N1# > 0 AND N2# > 0 THEN
    P# = Round#((N1# / N2#) * 100)
    P$ = STR$(INT(P#)) + "%"
    RSET Pct$ = P$
  END IF
  GetPct$ = Pct$
  
END FUNCTION

FUNCTION InQtr (TDate, RDate)
    
  '--Get the Report Quarter
  R$ = Num2Date(RDate)
  RM$ = LEFT$(R$, 2)

  '--Get the Transaction Quarter
  t$ = Num2Date(TDate)
  TM$ = LEFT$(t$, 2)

  SELECT CASE RM$
    CASE "01", "02", "03"
      RQ = 1
    CASE "04", "05", "06"
      RQ = 2
    CASE "07", "08", "09"
      RQ = 3
    CASE "10", "11", "12"
      RQ = 4
  END SELECT
  
  SELECT CASE TM$
    CASE "01", "02", "03"
      TQ = 1
    CASE "04", "05", "06"
      TQ = 2
    CASE "07", "08", "09"
      TQ = 3
    CASE "10", "11", "12"
      TQ = 4
  END SELECT

  IF TQ = RQ THEN
    InQtr = True
  ELSE
    InQtr = FALSE
  END IF

END FUNCTION

FUNCTION Num2Month% (Dt%)
  
  D$ = Num2Date(Dt%)
  m$ = RIGHT$(D$, 2) + LEFT$(D$, 2)
  Num2Month% = VAL(m$)
  
END FUNCTION

