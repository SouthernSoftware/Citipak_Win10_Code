DECLARE SUB BAcctacctbal ()
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
DECLARE SUB QPRintRC (T$, R%, c%, Clr%)
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
DECLARE SUB Export ()
DECLARE SUB FindAcctBal ()
DECLARE SUB BgtTrList ()
DECLARE FUNCTION Unique$ (Path$)
  
  CONST False = 0, True = NOT False
  
  '$INCLUDE: 'DefCnf.bi'                    'monitor and color info
  '$INCLUDE: 'FieldInf.bi'                  'field information
  '$INCLUDE: 'FormEdit.bi'                  'form editing information
  '$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
  '$INCLUDE: 'GL.bi'                        'GL File Types
  '$INCLUDE: 'SCRAUX.BI'
  '$INCLUDE: 'QSCR.BI'          'QuickScreen Declarations
  
  DIM SHARED TotLen%, FundLen%, AcctLen%, DetLen%
  DIM SHARED Acct AS GLAcctRecType
  DIM SHARED AcctIdx AS GLAcctIndexType
  DIM SHARED Trans AS GLTransRecType
  
  'REDIM SHARED Choice$(0, 0)    'Used to hold multiple choices
  
  '$INCLUDE: 'SetCnf.bi'        'Set monitor and color information
  '$INCLUDE: 'init.bi'
  
  'STOP
  'BgtTrList
  'BgtvsActFund
  'END
  
  LOCATE 1, 1, 0
  TextCursor -2, -2             'Define the Mouse cursor
  ShowCursor    'Show it
  
  REDIM MChoice$(1 TO 11)
  MChoice$(1) = "Trial Balance"
  MChoice$(2) = "Cash Balance"
  MChoice$(3) = "Account Balance Summary"
  MChoice$(4) = "Account History"
  MChoice$(5) = "Budget History"
  MChoice$(6) = "Balance Sheet"
  MChoice$(7) = "Budget vs Actual"
  MChoice$(8) = "Department Budget vs Actual"
  MChoice$(9) = "Query G/L Transactions"
  MChoice$(10) = "Export Files"
  MChoice$(11) = "Quit"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 20   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  FOR Cnt = 1 TO UBOUND(MChoice$)               'Find max menu width
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  'Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  Help$ = "General Ledger Reports Menu"
  
  DO
    
    LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, Attribute, ErrorCode
    'TitleBox 3, Col, MaxLen + 3, "Reports Menu  ", Cnf
    'TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    'CALL MTitle("Reports Menu", Row - 1, Col, MaxLen, 15)
    
    Title$ = "Reports Menu "
    MnuTitle Title$, Row, Col, MaxLen, 15
    PrintTitle User$
    PrintHelp Help$
    
    LOCATE Row, Col, 0
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO
    
    SELECT CASE Choice
    CASE 1
      ClrBG
      TrlBal
    CASE 2
      ClrBG
      'CashBalance
      BAcctacctbal              'unrem
    CASE 3
      ClrBG
      FindAcctBal
    CASE 4
      ClrBG
      AcctHist
    CASE 5
      ClrBG
      BgtHist
    CASE 6
      ClrBG
      BalSheet
    CASE 7
      ClrBG
      BgtvsAct User$, FundLen, AcctLen          'unrem
    CASE 8
      ClrBG
      BgtvsActDept User$, FundLen, AcctLen      'unrem
    CASE 9
      ClrBG
      AQuery
    CASE 10
      ClrBG
      Export    'unrem
    CASE 11
      ClrBG
      HideCursor                'Be nice and turn the mouse cursor Off
      CLS
      END
    END SELECT
    
  LOOP
  
  IF Exist("glmenu.exe") THEN
    ClrBG
    HideCursor  'Be nice and turn the mouse cursor Off
    DispLoad
    RUN "GLMENU"
  ELSE
    CLS
    END
  END IF

SUB AcctHist
  
  SHARED User$
  Null$ = CHR$(0)
  
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
  
  'REDIM Desc$(1)
  
  LibName$ = "GL.QSL"           'screen library
  FormName$ = "PRNHIST"         'form to edit
  
  DevFld = 6    'Set to output device field number
  SHARED Choice$()
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  REDIM Choice$(0 TO 3, 1)
  REDIM Frm(1) AS FormInfo
  
  Choice$(0, 0) = STR$(DevFld)  'Choices for Output Device
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"
  
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  
  Action = 1
  
  DO
    
    EditForm Form$(), Fld(), Frm(1), Cnf, Action
    
    IF Frm(1).PrevFld <> Frm(1).FldNo THEN
      
      '--strip out dashes for editing
      IF Frm(1).FldNo = 1 OR Frm(1).FldNo = 2 THEN
        
        IF LEN(Form$(Frm(1).FldNo, 0)) > 0 THEN
          Form$(Frm(1).FldNo, 0) = StripAcct$(Form$(Frm(1).FldNo, 0))
          SaveField Frm(1).FldNo, Form$(), Fld(), 0
          Action = 1
        END IF
      END IF
      
      SELECT CASE Frm(1).PrevFld
      CASE 1, 2
        FldLen = LEN(QPTrim$(Form$(Frm(1).PrevFld, 0)))
        IF FldLen > 0 THEN
          Form$(Frm(1).PrevFld, 0) = FmtAcct$(Form$(Frm(1).PrevFld, 0), FundLen%, AcctLen%, DetLen%)
        END IF
      CASE ELSE
      END SELECT
      Action = 1
      
    END IF
    
    SELECT CASE Frm(1).KeyCode  'Check the user's key presses
      
    CASE 42     '*=wildcard
      OpenAcctIdx AcctIdxFileNum, NumIdxRecs
      GET AcctIdxFileNum, 1, AcctIdx
      Form$(1, 0) = AcctIdx.AcctNum
      GET AcctIdxFileNum, NumIdxRecs, AcctIdx
      Form$(2, 0) = AcctIdx.AcctNum
      CLOSE AcctIdxFileNum
      
      D = Date2Num(DATE$)
      IF D > FY1EndDate THEN
        Form$(3, 0) = Num2Date(FY2BegDate)
        Form$(4, 0) = Num2Date(FY2EndDate)
      ELSE
        Form$(3, 0) = Num2Date(FY1BegDate)
        Form$(4, 0) = Num2Date(FY1EndDate)
      END IF
      Form$(5, 0) = "S"
      Form$(6, 0) = "Screen"
      Frm(1).FldNo = DevFld
      Action = 1
      
    CASE F10Key 'SaveButton
      'First Acct
      IF UCASE$(LEFT$(QPTrim$(Form$(1, 0)), 1)) = "F" THEN
        FirstAcct$ = "0"
      ELSE
        FirstAcct$ = QPTrim$(Form$(1, 0))
      END IF
      
      'Last Acct
      IF UCASE$(LEFT$(QPTrim$(Form$(2, 0)), 1)) = "L" THEN
        LastAcct$ = "999999999999"
      ELSE
        LastAcct$ = QPTrim$(Form$(2, 0))
      END IF
      
      BegDate = Date2Num(Form$(3, 0))
      EndDate = Date2Num(Form$(4, 0))
      
      IF LEN(LTRIM$(RTRIM$(Form$(DevFld, 0)))) = 0 THEN
        Dev$ = "S"
        LPTNo = 1
      ELSE
        Dev$ = LEFT$(Form$(DevFld, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(DevFld, 0)), 1))
      END IF
      
      IF BegDate > EndDate OR Form$(1, 0) > Form$(2, 0) THEN
        OK = MsgBox("GL", "BADPARMS")
        Ok2Go = 0
        Frm(1).FldNo = 5
        Action = 1
      ELSE
        Ok2Go = -1
      END IF
      
    CASE ESC
      EXIT SUB
      
    END SELECT
    
  LOOP UNTIL Ok2Go
  
  LOCATE , , 0
  
  SELECT CASE Dev$
  CASE "S"
    EntryPoint = 2
  CASE "P"
    EntryPoint = 5
  END SELECT
  SrcBalOpt$ = Form$(5, 0)
  IF SrcBalOpt$ = " " THEN SrcBalOpt$ = "S"
  RunThru$ = Form$(3, 0) + " thru " + Form$(4, 0)
  ERASE Form$, Fld, Choice$, Frm
  
  'End of Input
  '=====================================================
  'Start Report Processing

  
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  FF$ = CHR$(12)
  MaxLines = 53
  LineCnt = 0
  TotDr# = 0
  TotCr# = 0
  
  T$ = "Account History " + RunThru$ + "                     Page:"
  ToPrint$ = SPACE$(80)
  
  IF EndDate >= FY2BegDate THEN
    ActiveYear = 2
    OpenDate$ = ConvDateStr1(Num2Date$(FY2BegDate))
  ELSE
    ActiveYear = 1
    OpenDate$ = ConvDateStr1(Num2Date$(FY1BegDate))
  END IF
  
  ReportFile$ = Unique$(Path$)
  PrnFile = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  IF HollyFlag THEN
    PRINT #PrnFile, Pitch12$;
  END IF
  
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  
  IF NumTrans& = 0 THEN
    CLOSE
    EXIT SUB
  END IF
  
  '--trying to set TrSort array to max needed

  REDIM TrSort(1 TO 10000) AS TrSortType1
  'SHELL
  
  GrTotDr# = 0
  GrTotCr# = 0
  
  NewAcct = True
  GOSUB PrintAHPageHeader
  
  FOR Cnt = 1 TO NumGLAccts
    
    DrFwd# = 0
    CrFwd# = 0
    TotAcctDr# = 0
    TotAcctCr# = 0
    
    GET AcctIdxFileNum, Cnt, AcctIdx
    GET AcctFileNum, AcctIdx.RecNum, Acct
    
    QPRintRC SPACE$(80), 25, 1, -1
    AcctNum$ = QPTrim$(Acct.Num)
    QPRintRC "Processing: " + AcctNum$, 25, 2, -1
    
    IF AcctNum$ >= FirstAcct$ AND AcctNum$ <= LastAcct$ THEN
      
      BalFwd# = Round#(Acct.BegBal)             'get the beginning balance
      NextTr& = Acct.FrstTran   'get the first trans for this acct
      
      AcctNumber$ = "Account " + AcctNum$ + " - " + QPTrim$(Acct.Title)
      
      LSET ToPrint$ = ""
      LSET ToPrint$ = AcctNumber$
      PRINT #PrnFile, ToPrint$
      LineCnt = LineCnt + 1
      
      '--Deal with the bal fwd field
      IF BalFwd# <> 0 THEN
        '--Convert the balance fwd to a debit or a credit amt
        SELECT CASE Acct.Typ
        CASE "A", "E"
          IF BalFwd# >= 0 THEN
            Debit$ = FUsing$(STR$(BalFwd#), CommaFmt$)
            Credit$ = ""
            DrFwd# = BalFwd#
          ELSE
            Credit$ = FUsing$(STR$(ABS(BalFwd#)), CommaFmt$)
            Debit$ = ""
            CrFwd# = ABS(BalFwd#)
          END IF
        CASE "L", "R"
          IF BalFwd# >= 0 THEN
            Credit$ = FUsing$(STR$(BalFwd#), CommaFmt$)
            Debit$ = ""
            CrFwd# = BalFwd#
          ELSE
            Debit$ = FUsing$(STR$(ABS(BalFwd#)), CommaFmt$)
            Credit$ = ""
            DrFwd# = ABS(BalFwd#)
          END IF
        END SELECT
        
        '--If we're running a history with a Report Begin Date that is prior
        '  to the FY Beginning date then print as opening balance.
        IF BegDate <= FYBeg THEN
          LSET ToPrint$ = ""
          MID$(ToPrint$, 1) = OpenDate$
          MID$(ToPrint$, 12) = "Balance Foward"
          MID$(ToPrint$, 45) = Debit$
          MID$(ToPrint$, 59) = Credit$
          PRINT #PrnFile, ToPrint$
          LineCnt = LineCnt + 1
          IF LineCnt > MaxLines THEN
            PRINT #PrnFile, FF$
            NewAcct = False
            GOSUB PrintAHPageHeader
          END IF
        ELSE
          '--otherwise set a flag so we'll know we have a bal fwd
          '--to print
          FwdFlag = -1
        END IF
      END IF
      
      '--search thru acct transactions
      DO UNTIL NextTr& = 0
        GET TransFileNum, NextTr&, Trans
        
        '--see if transaction is in date range
        IF Trans.TrDate >= BegDate AND Trans.TrDate <= EndDate THEN
          
          '--within range - assign to array for sorting
          NumAcctTrans = NumAcctTrans + 1
          TrSort(NumAcctTrans).TrDate = STR$(Trans.TrDate) + QPTrim$(Trans.Ref)
          TrSort(NumAcctTrans).Record = NextTr&
          
        ELSE
          '--check the transaction to see if we need to carry it in
          '  the balance fwd
          IF Trans.TrDate < BegDate THEN
            
            '--Check here for dual year to bring asset & liab acct
            '  balances fwd but not rev & exp acct balances
            SELECT CASE ActiveYear
              
            CASE 1              '--Year 1
              DrFwd# = DrFwd# + Trans.DrAmt
              CrFwd# = CrFwd# + Trans.CrAmt
              FwdFlag = -1
              
            CASE 2              '--Year 2
              SELECT CASE Acct.Typ
                
                '--Get Asset & Liab acct balances as bal fwd
              CASE "A", "L"
                DrFwd# = DrFwd# + Trans.DrAmt
                CrFwd# = CrFwd# + Trans.CrAmt
                FwdFlag = -1
                
                '--Get Rev & Liab acct balances
              CASE "R", "E"
                IF Trans.TrDate > FY2BegDate THEN               'STOP
                  DrFwd# = DrFwd# + Trans.DrAmt
                  CrFwd# = CrFwd# + Trans.CrAmt
                  FwdFlag = -1
                END IF
                
              END SELECT
              
            END SELECT
          END IF                'IF Trans.TrDate < BegDate THEN
        END IF  '
        
        NextTr& = Trans.NextTran                'Get the next transaction
        
      LOOP
      
      IF FwdFlag THEN
        
        '--Print the balance forward
        LSET ToPrint$ = ""
        MID$(ToPrint$, 1) = "Balance Forward"
        SELECT CASE Acct.Typ
        CASE "A", "E"
          BalFwd# = DrFwd# - CrFwd#
          IF BalFwd# >= 0 THEN
            Debit$ = FUsing$(STR$(BalFwd#), CommaFmt$)
            Credit$ = ""
          ELSE
            Credit$ = FUsing$(STR$(ABS(BalFwd#)), CommaFmt$)
            Debit$ = ""
          END IF
        CASE "L", "R"
          BalFwd# = CrFwd# - DrFwd#
          IF BalFwd# >= 0 THEN
            Credit$ = FUsing$(STR$(BalFwd#), CommaFmt$)
            Debit$ = ""
          ELSE
            Debit$ = FUsing$(STR$(ABS(BalFwd#)), CommaFmt$)
            Credit$ = ""
          END IF
        END SELECT
        
        MID$(ToPrint$, 43) = Debit$
        MID$(ToPrint$, 57) = Credit$
        
        PRINT #PrnFile, ToPrint$
        
        LineCnt = LineCnt + 1
        IF LineCnt > MaxLines THEN
          NewAcct = False
          PRINT #PrnFile, FF$
          GOSUB PrintAHPageHeader
        END IF
        
      END IF
      
      IF NumAcctTrans > 0 OR BalFwd# <> 0 THEN
        SortT TrSort(1), NumAcctTrans, 0, 16, 0, 0
        AcctRunBal# = BalFwd#
        FOR Trn = 1 TO NumAcctTrans
          GET TransFileNum, TrSort(Trn).Record, Trans
          
          LSET ToPrint$ = ""
          MID$(ToPrint$, 1) = ConvDateStr1(Num2Date(Trans.TrDate))
          MID$(ToPrint$, 10) = QPTrim$(Trans.Desc)
          MID$(ToPrint$, 34) = QPTrim$(Trans.Ref)
          IF Trans.DrAmt <> 0 THEN
            MID$(ToPrint$, 43) = FUsing$(STR$(Trans.DrAmt), CommaFmt$)
          END IF
          IF Trans.CrAmt <> 0 THEN
            MID$(ToPrint$, 57) = FUsing$(STR$(Trans.CrAmt), CommaFmt$)
          END IF
          SELECT CASE SrcBalOpt$
          CASE "R"
            SELECT CASE Acct.Typ
            CASE "A", "E"
              AcctRunBal# = AcctRunBal# + Trans.DrAmt - Trans.CrAmt
            CASE "L", "R"
              AcctRunBal# = AcctRunBal# + Trans.CrAmt - Trans.DrAmt
            END SELECT
            MID$(ToPrint$, 70) = FUsing$(STR$(AcctRunBal#), "#######.##")
          CASE "S"
            MID$(ToPrint$, 74) = LEFT$(Trans.Src, 6)
          END SELECT
          
          PRINT #PrnFile, ToPrint$
          LineCnt = LineCnt + 1
          IF LineCnt > MaxLines THEN
            NewAcct = False
            PRINT #PrnFile, FF$
            GOSUB PrintAHPageHeader
          END IF
          
          TotAcctDr# = TotAcctDr# + Trans.DrAmt
          TotAcctCr# = TotAcctCr# + Trans.CrAmt
          
          GrTotDr# = GrTotDr# + Trans.DrAmt
          GrTotCr# = GrTotCr# + Trans.CrAmt
          
        NEXT
        
        '--Print summary lines
        LSET ToPrint$ = ""
        MID$(ToPrint$, 42) = SumLine$
        MID$(ToPrint$, 56) = SumLine$
        PRINT #PrnFile, ToPrint$
        LineCnt = LineCnt + 1
        
        '--Print transaction totals
        IF NumAcctTrans > 0 THEN
          'LSET ToPrint$ = ""
          LSET ToPrint$ = "Transaction Totals"
          MID$(ToPrint$, 42) = FUsing$(STR$(TotAcctDr#), TotalFmt$)
          MID$(ToPrint$, 56) = FUsing$(STR$(TotAcctCr#), TotalFmt$)
          PRINT #PrnFile, ToPrint$
          LineCnt = LineCnt + 1
        END IF
        
        '--Print ending balance
        PRINT #PrnFile,
        LSET ToPrint$ = ""
        MID$(ToPrint$, 1) = "Ending Balance"
        SELECT CASE Acct.Typ
        CASE "A", "E"
          AcctBal# = BalFwd# + TotAcctDr# - TotAcctCr#
          IF AcctBal# >= 0 THEN
            Debit$ = FUsing$(STR$(AcctBal#), TotalFmt$)
            Credit$ = ""
          ELSE
            Credit$ = FUsing$(STR$(ABS(AcctBal#)), TotalFmt$)
            Debit$ = ""
          END IF
        CASE "L", "R"
          AcctBal# = BalFwd# + TotAcctCr# - TotAcctDr#
          IF AcctBal# >= 0 THEN
            Credit$ = FUsing$(STR$(AcctBal#), TotalFmt$)
            Debit$ = ""
          ELSE
            Debit$ = FUsing$(STR$(ABS(AcctBal#)), TotalFmt$)
            Credit$ = ""
          END IF
        END SELECT
        MID$(ToPrint$, 42) = Debit$
        MID$(ToPrint$, 56) = Credit$
        PRINT #PrnFile, ToPrint$
        LineCnt = LineCnt + 1
        
        IF Acct.Typ = "R" OR Acct.Typ = "E" THEN
          IF Acct.Typ = "R" THEN
            BgtCol = 56
          ELSE
            BgtCol = 42
          END IF
          
          SELECT CASE ActiveYear
          CASE 1
            BudgetAmt# = Acct.Bgt
          CASE 2
            BudgetAmt# = Acct.NYApp
          END SELECT
          
          LSET ToPrint$ = ""
          LSET ToPrint$ = "Budget"
          MID$(ToPrint$, BgtCol) = FUsing$(STR$(ABS(BudgetAmt#)), TotalFmt$)
          PRINT #PrnFile, ToPrint$
          LineCnt = LineCnt + 1
          
          Var# = BudgetAmt# - AcctBal#
          SELECT CASE Acct.Typ
          CASE "R"
            IF Var# > 0 THEN
              VarCol = 56
              VarText$ = "Uncollected Balance"
            ELSE
              VarCol = 56
              VarText$ = "Revenues Exceeding Budget"
            END IF
          CASE "E"
            IF Var# > 0 THEN
              VarCol = 42
              VarText$ = "Appropriation Remaining"
            ELSE
              VarCol = 42
              VarText$ = "Over Spent"
            END IF
          END SELECT
          LSET ToPrint$ = ""
          LSET ToPrint$ = VarText$
          MID$(ToPrint$, VarCol) = FUsing$(STR$(ABS(Var#)), TotalFmt$)
          PRINT #PrnFile, ToPrint$
          LineCnt = LineCnt + 1
        END IF
        
        ToPrint$ = STRING$(80, "=")
        PRINT #PrnFile, ToPrint$
        LineCnt = LineCnt + 1
        
        '--Don't break up summary section
        IF LineCnt > MaxLines THEN
          NewAcct = True
          PRINT #PrnFile, FF$
          GOSUB PrintAHPageHeader
        END IF
        
        
      ELSE
        LSET ToPrint$ = ""
        MID$(ToPrint$, 5) = " -- No Activity --"
        PRINT #PrnFile, ToPrint$
        LineCnt = LineCnt + 1

        'Don't orphan the next line
        ToPrint$ = STRING$(80, "=")
        PRINT #PrnFile, ToPrint$
        LineCnt = LineCnt + 1
        IF LineCnt > MaxLines THEN
          NewAcct = True
          PRINT #PrnFile, FF$
          GOSUB PrintAHPageHeader
        END IF
        
      END IF
    END IF      'Account is not of this fund
    
    NumAcctTrans = 0            'reset for next account
    
  NEXT
  
  LSET ToPrint$ = ""
  LSET ToPrint$ = "Grand Total Debits"
  MID$(ToPrint$, 25) = FUsing$(STR$(GrTotDr#), "##########,.##")
  PRINT #PrnFile, ToPrint$
  
  LSET ToPrint$ = ""
  LSET ToPrint$ = "Grand Total Credits"
  MID$(ToPrint$, 25) = FUsing$(STR$(GrTotCr#), "##########,.##")
  PRINT #PrnFile, ToPrint$
  
  IF ActiveYear = 2 THEN
    LSET ToPrint$ = ""
    LSET ToPrint$ = "Note: Unaudited balances - prior year has not been closed."
    PRINT #PrnFile, ToPrint$
  END IF
  
  PRINT #PrnFile, FF$
  CLOSE
  
  ERASE TrSort
  
  'End Report Processing
  '===========================================================================

  PrintRptFile "Account History", ReportFile$, LPTNo, RetCode%, EntryPoint
  EXIT SUB

  
PrintAHPageHeader:
  PageNum = PageNum + 1
  PRINT #PrnFile, User$; TAB(43); "Run Date: " + DATE$
  PRINT #PrnFile, T$; PageNum
  PRINT #PrnFile,
  SELECT CASE SrcBalOpt$
  CASE "S", " "
    PRINT #PrnFile, "Date     Description             Reference       Debit        Credit    Post Ref"
  CASE "R"
    PRINT #PrnFile, "Date     Description             Reference       Debit        Credit     Run Bal"
  END SELECT
  PRINT #PrnFile, "--------------------------------------------------------------------------------"
  IF NewAcct THEN
    LineCnt = 5
  ELSE
    PRINT #PrnFile, AcctNumber$ + " continued."
    LineCnt = 5
    NewAcct = True
  END IF
  RETURN
  
END SUB

SUB AQuery
  
  LibName$ = "GL.QSL"           'screen library
  FormName$ = "QRYTR"           'form to edit
  
  SHARED Choice$()
  
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  REDIM Choice$(0 TO 11, 0 TO 2)
  DIM Frm AS FormInfo
  
  DIM FundIdx AS GLFundIndexType
  
  Choice$(0, 0) = "5"           'Choices for Acct Type
  Choice$(1, 0) = "All"
  Choice$(2, 0) = "Asset"
  Choice$(3, 0) = "Liability"
  Choice$(4, 0) = "Revenue"
  Choice$(5, 0) = "Expenditure"
  
  Choice$(0, 1) = "10"          'Choices for Module
  Choice$(1, 1) = "Any"
  Choice$(2, 1) = "AR"
  Choice$(3, 1) = "AP"
  Choice$(4, 1) = "CD"
  Choice$(5, 1) = "CK"
  Choice$(6, 1) = "CR"
  Choice$(7, 1) = "GJ"
  Choice$(8, 1) = "IF"
  Choice$(9, 1) = "PR"
  Choice$(10, 1) = "VC"
  Choice$(11, 1) = "VI"
  
  Choice$(0, 2) = "12"
  Choice$(1, 2) = "Screen"
  Choice$(2, 2) = "Prn#1"
  Choice$(3, 2) = "Prn#2"
  
  MaxLines = 53
  FF$ = CHR$(12)
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  RptTitle$ = "Transaction Query"
  TotDr# = 0
  TotCr# = 0
  
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  
  Form$(1, 0) = Num2Date(GetBegDate)
  Form$(2, 0) = Num2Date(GetEndDate)
  GetFundCodes Form$(3, 0), Form$(4, 0)
  Form$(5, 0) = "All"
  
  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    SELECT CASE Frm.KeyCode     'Check the user's key presses
      
    CASE F10Key 'SaveButton
      LowDate = Date2Num(Form$(1, 0))
      HighDate = Date2Num(Form$(2, 0))
      StartFund$ = QPTrim$(Form$(3, 0))
      EndFund$ = QPTrim$(Form$(4, 0))
      AType$ = QPTrim$(Form$(5, 0))
      AcctCode$ = QPTrim$(Form$(6, 0))
      ObjCode$ = QPTrim$(Form$(7, 0))
      TSrch$ = QPTrim$(Form$(8, 0))
      TRef$ = QPTrim$(Form$(9, 0))
      Src$ = QPTrim$(Form$(10, 0))
      PostDate$ = QPTrim$(Form$(11, 0))
      
      IF LEN(LTRIM$(RTRIM$(Form$(12, 0)))) = 0 THEN
        Dev$ = "S"
        LPTNo = 1
      ELSE
        Dev$ = LEFT$(Form$(12, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(12, 0)), 1))
      END IF
      
    CASE ESC
      EXIT SUB
      
    END SELECT
    
  LOOP UNTIL Frm.KeyCode = F10Key
  
  IF AType$ = "All" THEN
    AType$ = ""
    AT$ = "All"
  ELSE
    AType$ = LEFT$(AType$, 1)
    AT$ = AType$
  END IF
  
  IF Src$ = "Any" THEN
    Src$ = ""
    Sc$ = "Any "
  ELSE
    Sc$ = Src$
  END IF
  
  IF AcctCode$ = "" THEN
    AC$ = "All "
  ELSE
    AC$ = AcctCode$
  END IF
  
  IF ObjCode$ = "" THEN
    Obj$ = "All "
  ELSE
    Obj$ = ObjCode$
  END IF
  
  'End of Input
  '=====================================================
  'Start Report Processing
  
  LOCATE 1, 1, 0
  
  'ReportFile$ = "QUERY.PRN"
  ReportFile$ = Unique$(Path$)
  
  PrnFile = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFileNum, NumFIdxRecs
  
  QPRintRC SPACE$(80), 25, 1, 112
  QPRintRC " Searching:", 25, 1, -1
  
  GOSUB PrintQueryPageHeader
  
  FOR Cnt = 1 TO NumGLAccts     'NumGLAccts
    GET AcctIdxFileNum, Cnt, AcctIdx
    
    AcctNumber$ = QPTrim$(AcctIdx.AcctNum)
    QPRintRC AcctNumber$, 25, 13, -1
    
    Fund$ = LEFT$(AcctNumber$, FundLen)
    Dept$ = MID$(AcctNumber$, FundLen + 2, AcctLen)
    Det$ = RIGHT$(AcctNumber$, DetLen)
    
    IF Fund$ >= StartFund$ AND Fund$ <= EndFund$ THEN
      IF INSTR(Dept$, AcctCode$) AND INSTR(Det$, ObjCode$) THEN
        
        GET AcctFileNum, AcctIdx.RecNum, Acct
        IF INSTR(Acct.Typ, AType$) THEN
          
          NextTr& = Acct.FrstTran               'get the first trans for this acct
          
          DO UNTIL NextTr& = 0  'keep going 'til we run out of trans
            GET TransFileNum, NextTr&, Trans
            
            IF Trans.TrDate >= LowDate AND Trans.TrDate <= HighDate THEN
              TrDateStamp$ = MID$(Trans.Src, 3, 6)
              IF INSTR(TrDateStamp$, PostDate$) THEN
                IF INSTR(Trans.Src, Src$) THEN
                  IF INSTR(UCASE$(Trans.Desc), UCASE$(TSrch$)) THEN             'kill me
                    IF INSTR(1, Trans.Ref, TRef$) THEN
                      Match = Match + 1
                      TotDr# = TotDr# + Round#(Trans.DrAmt)
                      TotCr# = TotCr# + Round#(Trans.CrAmt)
                      
                      ToPrint$ = SPACE$(100)
                      'LSET ToPrint$ = ConvDateStr1(Num2Date(Trans.TrDate))
                      LSET ToPrint$ = Num2Date(Trans.TrDate)
                      MID$(ToPrint$, 12) = Trans.AcctNum
                      MID$(ToPrint$, 28) = LEFT$(Trans.Desc, 15)
                      MID$(ToPrint$, 45) = Trans.Ref
                      IF Trans.DrAmt <> 0 THEN
                        MID$(ToPrint$, 55) = FUsing$(STR$(Trans.DrAmt), CommaFmt$)
                      ELSE
                        MID$(ToPrint$, 55) = ""
                      END IF
                      IF Trans.CrAmt <> 0 THEN
                        MID$(ToPrint$, 69) = FUsing$(STR$(Trans.CrAmt), CommaFmt$)
                      ELSE
                        MID$(ToPrint$, 69) = ""
                      END IF
                      MID$(ToPrint$, 82) = RecNum$
                      MID$(ToPrint$, 90) = Trans.Src
                      PRINT #PrnFile, ToPrint$
                      LineCnt = LineCnt + 1
                      IF LineCnt > MaxLines THEN
                        PRINT #PrnFile, FF$
                        GOSUB PrintQueryPageHeader
                      END IF
                    END IF
                  END IF        'kill me
                END IF
              END IF
            END IF
            
            NextTr& = Trans.NextTran            'Get the next transaction
            
          LOOP
        END IF
      END IF
    END IF
  NEXT
  
  Diff# = Round#(TotDr#) - Round#(TotCr#)
  Bal# = ABS(Diff#)
  
  IF Diff# = 0 THEN
    DrCr$ = ""
  ELSEIF Diff# > 0 THEN
    DrCr$ = " Dr"
  ELSE
    DrCr$ = " Cr"
  END IF
  
  PRINT #PrnFile, STRING$(78, "-")
  PRINT #PrnFile, "Totals";
  PRINT #PrnFile, TAB(54); USING "#########,.##"; Round#(TotDr#);
  PRINT #PrnFile, TAB(68); USING "#########,.##"; Round#(TotCr#)
  PRINT #PrnFile,
  PRINT #PrnFile, "Report Summary"
  PRINT #PrnFile, USING "Balance:                   #########,.##"; Round#(Bal#); : PRINT #PrnFile, DrCr$
  PRINT #PrnFile, USING "Number of transactions:        #####,"; Match
  PRINT #PrnFile, FF$
  
  CLOSE
  
  SELECT CASE Dev$
  CASE "S"
    EntryPoint = 2
  CASE "P"
    EntryPoint = 5
  CASE ELSE
  END SELECT
  
  PrintRptFile Header$, ReportFile$, LPTNo, RetCode%, EntryPoint
  KILL ReportFile$
  
  EXIT SUB
  
PrintQueryPageHeader:
  
  PRINT #PrnFile, RptTitle$
  PRINT #PrnFile,
  PRINT #PrnFile, "Date Range: " + Form$(1, 0) + " thru " + Form$(2, 0) + SPACE$(2) + "Module: " + Sc$
  PRINT #PrnFile, "Funds " + StartFund$ + " thru " + EndFund$ + "  Account: " + AC$ + "  Object: " + Obj$ + "  Account Type: " + AT$
  PRINT #PrnFile, "Date        Acct No        Desc             Ref             Debit         Credit          Source"
  PRINT #PrnFile, "------------------------------------------------------------------------------------------------"
  LineCnt = 6
  RETURN
  
  
END SUB

SUB BalSheet
  
  SHARED User$
  LibName$ = "GL.QSL"           'screen library
  FormName$ = "PRNBALSH"        'form to edit
  
  DevFld = 4    'Set to output device field number
  
  SHARED Choice$()
  
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  REDIM Choice$(0 TO 3, 1)
  DIM Frm AS FormInfo
  
  DIM FundIdx AS GLFundIndexType
  
  Choice$(0, 0) = STR$(DevFld)  'Choices for Output Device
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"
  
  GetFundCodes FirstFund$, LastFund$
  
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  
  Action = 1
  
  DO            'Allow user to input report params
    RunFlag = 0
    
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    IF Frm.PrevFld <> Frm.FldNo THEN            'User is moving around in form
      
      SELECT CASE Frm.PrevFld
        
      CASE 1    'Ending Date
        EndDate = Date2Num(Form$(1, 0))
        IF EndDate < -32000 THEN
          OK = MsgBox("GL.QSL", "BADDATE")
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
              OK = MsgBox("GL.QSL", "NOFUND")
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
      
    CASE 42
      Form$(2, 0) = FirstFund$
      Form$(3, 0) = LastFund$
      Form$(1, 0) = Num2Date(GetEndDate)
      Form$(4, 0) = "Screen"
      Frm.FldNo = DevFld
      Action = 1
      
    CASE F10Key 'SaveButton
      'check date
      EndDate = Date2Num(Form$(1, 0))
      IF EndDate < -32000 THEN
        OK = MsgBox("GL.QSL", "BADDATE")
        BadFld = 1
      END IF
      IF BadFld THEN
        Frm.FldNo = 1
        Action = 1
        GOTO Again
      ELSE
        RunFlag = -1
      END IF
      
      'check for valid funds
      FOR FieldNum = 2 TO 3
        IF LEN(QPTrim$(Form$(FieldNum, 0))) = 0 THEN
          OK = MsgBox("GL.QSL", "NOBLANK")
          Frm.FldNo = FieldNum
          BadFld = FieldNum
          EXIT FOR
        END IF
      NEXT
      IF BadFld THEN
        Action = 1
        RunFlag = 0
        GOTO Again
      ELSE
        RunFlag = -1
      END IF
      
      'make sure funds are in ascending order
      StartFund$ = QPTrim$(Form$(2, 0))
      EndFund$ = QPTrim$(Form$(3, 0))
      IF StartFund$ > EndFund$ THEN
        OK = MsgBox("GL.QSL", "RANGERR")
        BadFld = 1
      END IF
      IF BadFld THEN
        RunFlag = 0
        Frm.FldNo = 2
        Action = 1
        GOTO Again
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
    
Again:
    BadFld = 0
  LOOP UNTIL RunFlag
  
  LOCATE , , 0
  
  IF Dev$ = "F" THEN Dev$ = "S"
  
  
  'End of Input
  '=====================================================
  'Start Report Processing
  
  REDIM Desc$(1)
  
  'ReportFile$ = "BALSHEET.PRN"  'Report File Name
  
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(79, "-")   'dashed line
  DivLine2$ = STRING$(79, "=")  'Double Line
  CrLf$ = CHR$(13) + CHR$(10)
  PageBreak$ = CHR$(12)
  
  RptTitle$ = "Balance Sheet"
  Desc$(1) = "Acct Number     Title                                                Balance"
  'Desc$ = "Acct Number     Title                                                   Balance"
  TotDr# = 0
  TotCr# = 0
  MaxLines = 55
  
  ReportFile$ = Unique$(Path$)
  PrnFile = FREEFILE
  
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFile, NumFunds
  
  REDIM Asset%(1 TO NumGLAccts) 'Holds all Asset acct record nums
  REDIM Liab%(1 TO NumGLAccts)  'Holds all Liab acct record nums
  REDIM FundList$(1 TO NumFunds)                'List of all active Funds
  REDIM FundTotAssets#(1 TO NumFunds)           'List of total assets by fund
  REDIM FundTotLiab#(1 TO NumFunds)             'list of total liab by fund
  REDIM FundTotRev#(1 TO NumFunds)              'List of total revenues by fund
  REDIM FundTotExp#(1 TO NumFunds)              'list of tot exp by fund
  
  'Initialize FundList
  IF NumFunds = 0 THEN
    OK = MsgBox("GL", "NOFUNDS")
    CLOSE
    EXIT SUB
  END IF
  FOR Fund = 1 TO NumFunds
    GET FundIdxFile, Fund, FundIdx
    FundList$(Fund) = QPTrim$(FundIdx.FundNum)
  NEXT
  
  PrintHelp "Processing:"
  
  FOR GLAcct = 1 TO NumGLAccts
    
    GET AcctIdxFileNum, GLAcct, AcctIdx
    GET AcctFileNum, AcctIdx.RecNum, Acct
    
    FundCode$ = LEFT$(Acct.Num, FundLen)
    IF FundCode$ >= StartFund$ AND FundCode$ <= EndFund$ THEN
      
      '--Show the account we're processing
      QPRintRC Acct.Num, 25, 14, -1
      
      '--get the beginning balance and the first transaction rec num
      CalcBal# = Round#(Acct.BegBal)
      NextTr& = Acct.FrstTran
      
      '--Loop until we run out of transactions for this account
      DO UNTIL NextTr& = 0
        
        GET TransFileNum, NextTr&, Trans
        
        IF Trans.TrDate <= EndDate THEN
          
          '--Calculate balance depending on account type
          SELECT CASE Acct.Typ
          CASE "A", "E"
            CalcBal# = CalcBal# + Round#(Trans.DrAmt) - Round#(Trans.CrAmt)
          CASE "L", "R"
            CalcBal# = CalcBal# + Round#(Trans.CrAmt) - Round#(Trans.DrAmt)
          END SELECT
          
        END IF
        
        '--Get the next transaction
        NextTr& = Trans.NextTran
        
      LOOP
      
      '--Update account balance
      Acct.Bal = Round#(CalcBal#)
      PUT AcctFileNum, AcctIdx.RecNum, Acct
      
      '--Apply balance to proper fund
      FOR Fund = 1 TO NumFunds
        IF FundCode$ = FundList$(Fund) THEN
          SELECT CASE Acct.Typ
          CASE "A"
            TotAsset# = TotAsset# + CalcBal#
            ACnt = ACnt + 1
            Asset%(ACnt) = AcctIdx.RecNum
            FundTotAssets#(Fund) = FundTotAssets#(Fund) + CalcBal#
            
          CASE "L"
            TotLiab# = TotLiab# + CalcBal#
            LCnt = LCnt + 1
            Liab%(LCnt) = AcctIdx.RecNum
            FundTotLiab#(Fund) = FundTotLiab#(Fund) + CalcBal#
            
          CASE "R"
            TotRev# = TotRev# + CalcBal#
            FundTotRev#(Fund) = FundTotRev#(Fund) + CalcBal#
            
          CASE "E"
            TotExp# = TotExp# + CalcBal#
            FundTotExp#(Fund) = FundTotExp#(Fund) + CalcBal#
            
          END SELECT
          
          EXIT FOR              'acct was a match to this fund so exit loop
        END IF  'test for fund in list
      NEXT      'Summarize next fund's totals
    END IF      'End of fund in range test
  NEXT          'Process next account
  
  
  '--Now write the report to file.
  PrintHelp "Writing report..."
  
  '--Go thru fund list to see if its one we're reporting on
  FOR Fund = 1 TO NumFunds
    ThisFund$ = FundList$(Fund)
    IF ThisFund$ >= StartFund$ AND ThisFund$ <= EndFund$ THEN
      UsingFund = True
      FundRec = FindFund(ThisFund$)
      FundName$ = GetFundTitle(FundRec)
    ELSE
      UsingFund = False
    END IF
    
    IF UsingFund THEN
      '--List Assets for this fund
      AcctType$ = "Assets"
      GOSUB PrintBSHeader
      
      FOR Cnt = 1 TO ACnt
        Rec = Asset%(Cnt)
        GET AcctFileNum, Rec, Acct
        FundCode$ = LEFT$(Acct.Num, FundLen)
        
        IF FundCode$ = ThisFund$ THEN
          ToPrint$ = SPACE$(80)
          LSET ToPrint$ = Acct.Num
          MID$(ToPrint$, 17) = Acct.Title
          IF Acct.Bal >= 0 THEN
            Debit$ = FUsing$(STR$(Acct.Bal), CommaFmt$)
            Credit$ = " "
          ELSE
            Debit$ = " "
            Credit$ = "(" + FUsing$(STR$(ABS(Acct.Bal)), CommaFmt$) + ")"
          END IF
          MID$(ToPrint$, 68) = Debit$
          MID$(ToPrint$, 67) = Credit$
          PRINT #PrnFile, ToPrint$
          
          LineCnt = LineCnt + 1
          IF LineCnt >= MaxLines THEN
            GOSUB BreakBSPage
            GOSUB PrintBSHeader
          END IF
          
        END IF
        
      NEXT      'AssetAcct
      
      ToPrint$ = SPACE$(80)
      IF FundTotAssets#(Fund) >= 0 THEN
        DrCr$ = "Dr"
        FundAsset# = FundTotAssets#(Fund)
      ELSE
        DrCr$ = "Cr"
        FundAsset# = ABS(FundTotAssets#(Fund))
      END IF
      
      PRINT #PrnFile, DivLine$
      LSET ToPrint$ = "Total Assets"
      MID$(ToPrint$, 67) = FUsing$(STR$(FundAsset#), TotalFmt$) '+ " " + DrCr$
      PRINT #PrnFile, ToPrint$
      
      '--List Liabilities
      AcctType$ = "Liabilities"
      GOSUB BreakBSPage
      GOSUB PrintBSHeader
      
      FOR Cnt = 1 TO LCnt
        Rec = Liab%(Cnt)
        GET AcctFileNum, Rec, Acct
        
        FundCode$ = LEFT$(Acct.Num, FundLen)
        IF FundCode$ = ThisFund$ THEN
          ToPrint$ = SPACE$(80)
          LSET ToPrint$ = Acct.Num
          MID$(ToPrint$, 17) = Acct.Title
          IF Acct.Bal >= 0 THEN
            Credit$ = FUsing$(STR$(Acct.Bal), CommaFmt$)
            Debit$ = ""
          ELSE
            Credit$ = ""
            Debit$ = "(" + FUsing$(STR$(ABS(Acct.Bal)), CommaFmt$) + ")"
          END IF
          
          MID$(ToPrint$, 67) = Debit$
          MID$(ToPrint$, 68) = Credit$
          PRINT #PrnFile, ToPrint$
          LineCnt = LineCnt + 1
          IF LineCnt >= MaxLines THEN
            GOSUB BreakBSPage
            GOSUB PrintBSHeader
          END IF
          
        END IF
        
      NEXT      'LiabAcct
      
      FundBalAdj# = Round#(FundTotRev#(Fund) - FundTotExp#(Fund))
      LSET ToPrint$ = "Current Fund Balance Adjustment"
      MID$(ToPrint$, 68) = FUsing$(STR$(FundBalAdj#), CommaFmt$)                '+ " " + DrCr$
      PRINT #PrnFile, ToPrint$
      
      ToPrint$ = SPACE$(80)
      IF FundTotLiab#(Fund) >= 0 THEN
        DrCr$ = "Cr"
        FundLiab# = FundTotLiab#(Fund)
      ELSE
        DrCr$ = "Dr"
        FundLiab# = ABS(FundTotLiab#(Fund))
      END IF
      PRINT #PrnFile, DivLine$
      
      TotLiabnCap# = FundLiab# + FundBalAdj#
      
      LSET ToPrint$ = "Total Liabilities & Fund Balance"
      MID$(ToPrint$, 67) = FUsing$(STR$(TotLiabnCap#), TotalFmt$)               ' + " " + DrCr$
      PRINT #PrnFile, ToPrint$
      
      PRINT #PrnFile, PageBreak$
      
      '--Now summarize the fund
      'ToPrint$ = SPACE$(80)
      'LSET ToPrint$ = "Fund Summary:"
      'PRINT #PrnFile, ToPrint$
      'ToPrint$ = SPACE$(80)
      
      'LSET ToPrint$ = "Revenues"
      'MID$(ToPrint$, 17) = FUsing$(STR$(FundTotRev#(Fund)), CommaFmt$)
      'PRINT #PrnFile, ToPrint$
      'ToPrint$ = SPACE$(80)
      
      'LSET ToPrint$ = "Expenditures"
      'MID$(ToPrint$, 17) = FUsing$(STR$(FundTotExp#(Fund)), CommaFmt$)
      'PRINT #PrnFile, ToPrint$
      
      'FundBalAdj# = Round#(FundTotRev#(Fund) - FundTotExp#(Fund))
      'IF FundBalAdj# < 0 THEN
      '   DrCr$ = " Debit"
      'ELSE
      '   DrCr$ = " Credit"
      'END IF
      
      'ToPrint$ = SPACE$(80)
      'LSET ToPrint$ = "Fund Bal Adj"
      'MID$(ToPrint$, 17) = FUsing$(STR$(FundBalAdj#), CommaFmt$) + DrCr$
      'PRINT #PrnFile, ToPrint$
      
    END IF      'using fund test
    
  NEXT          'FundList$()
  
  '--This summarizes all funds--
  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Assets"
  'MID$(ToPrint$, 17) = STR$(TotAsset#)
  'MID$(ToPrint$, 37) = STR$(ACnt)
  'PRINT #PrnFile, ToPrint$
  
  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Liabilities"
  'MID$(ToPrint$, 17) = STR$(TotLiab#)
  'MID$(ToPrint$, 37) = STR$(LCnt)
  'PRINT #PrnFile, ToPrint$
  '
  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Revenues"
  'MID$(ToPrint$, 17) = STR$(TotRev#)
  'PRINT #PrnFile, ToPrint$
  '
  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Expenditures"
  'MID$(ToPrint$, 17) = STR$(TotExp#)
  'PRINT #PrnFile, ToPrint$
  '
  'FundBalAdj# = TotRev# - TotExp#
  'ToPrint$ = SPACE$(80)
  'LSET ToPrint$ = "Fund Bal Adj"
  'MID$(ToPrint$, 17) = STR$(FundBalAdj#)
  'PRINT #PrnFile, ToPrint$
  '^--All funds summary --
  
  'PRINT #PrnFile, PageBreak$
  
  CLOSE
  
  'End Report Processing
  '===========================================================================
  'Start Report Printing
  
  SELECT CASE Dev$
  CASE "S"
    EntryPoint = 2
  CASE "P"
    EntryPoint = 5
  END SELECT
  
  PrintRptFile Header$, ReportFile$, LPTNo, RetCode%, EntryPoint
  KILL ReportFile$
  
  EXIT SUB
  
PrintBSHeader:
  PRINT #PrnFile, User$; TAB(43); "Run Date: " + DATE$
  PRINT #PrnFile, QPTrim$(FundName$) + " " + RptTitle$
  PRINT #PrnFile, "Period Ending: " + Form$(1, 0)
  PRINT #PrnFile,
  PRINT #PrnFile, AcctType$
  PRINT #PrnFile, Desc$(1)
  PRINT #PrnFile, DivLine$
  LineCnt = 7
  RETURN
  
BreakBSPage:
  PRINT #PrnFile, PageBreak$
  LineCnt = 0
  RETURN
  
  
END SUB

SUB BgtHist
  
  REDIM Desc$(1)
  
  LibName$ = "GL.QSL"           'screen library
  FormName$ = "PRNBHIST"        'form to edit
  DevFld = 5    'Set to output device field number
  
  'SHARED Cnf AS Config
  SHARED Choice$()
  
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  REDIM Choice$(0 TO 3, 1)
  DIM Frm AS FormInfo
  
  Choice$(0, 0) = STR$(DevFld)  'Choices for Output Device
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"
  'Choice$(4, 0) = "Prn#3"
  'Choice$(5, 0) = "File"
  
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  
  Action = 1
  
  DO            'Allow user to input report params
    
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    IF Frm.PrevFld <> Frm.FldNo THEN            'User is moving around in form
      
      IF Frm.FldNo = 1 OR Frm.FldNo = 2 THEN
        Form$(Frm.FldNo, 0) = ""
        Action = 1
      END IF
      
      '--strip out dashes for editing
      IF Frm.FldNo = 1 OR Frm.FldNo = 2 THEN
        IF LEN(Form$(Frm.FldNo, 0)) > 0 THEN
          Form$(Frm.FldNo, 0) = StripAcct$(Form$(Frm.FldNo, 0))
          SaveField Frm.FldNo, Form$(), Fld(), 0
          Action = 1
        END IF
      END IF
      
      SELECT CASE Frm.PrevFld
      CASE 1, 2
        Form$(Frm.PrevFld, 0) = FmtAcct$(Form$(Frm.PrevFld, 0), FundLen%, AcctLen%, DetLen%)
      CASE ELSE
      END SELECT
      Action = 1
      
    END IF
    
    SELECT CASE Frm.KeyCode     'Check the user's key presses
      
    CASE 42
      Form$(1, 0) = "First Account"
      Form$(2, 0) = "Last Account"
      Form$(3, 0) = Num2Date(GetBegDate)
      Form$(4, 0) = Num2Date(GetEndDate)
      Frm.FldNo = DevFld
      Action = 1
      
    CASE F10Key 'SaveButton
      'First Acct
      IF UCASE$(LEFT$(QPTrim$(Form$(1, 0)), 1)) = "F" THEN
        FirstAcct$ = "000000000000"
      ELSE
        FirstAcct$ = QPTrim$(Form$(1, 0))
      END IF
      
      'Last Acct
      IF UCASE$(LEFT$(QPTrim$(Form$(2, 0)), 1)) = "L" THEN
        LastAcct$ = "999999999999"
      ELSE
        LastAcct$ = QPTrim$(Form$(2, 0))
      END IF
      
      'Report Beginning and ending date range
      BegDate = Date2Num(Form$(3, 0))
      EndDate = Date2Num(Form$(4, 0))
      
      IF LEN(LTRIM$(RTRIM$(Form$(DevFld, 0)))) = 0 THEN
        Dev$ = "S"
        LPTNo = 1
      ELSE
        Dev$ = LEFT$(Form$(DevFld, 0), 1)
        LPTNo = VAL(RIGHT$(RTRIM$(Form$(DevFld, 0)), 1))
      END IF
      
      IF BegDate > EndDate OR Form$(1, 0) > Form$(2, 0) THEN
        OK = MsgBox("GL", "BADPARMS")
        Ok2Go = 0
        Frm.FldNo = 5
        Action = 1
      ELSE
        Ok2Go = -1
      END IF
      
    CASE ESC
      EXIT SUB
      
    END SELECT
    
  LOOP UNTIL Ok2Go
  
  LOCATE , , 0
  
  IF Dev$ = "F" THEN Dev$ = "S"
  
  
  'End of Input
  '=====================================================
  'Start Report Processing
  
  'ReportFile$ = "BGTHIST.PRN"   'Report File Name
  
  DescColor = 3 'Color of description text
  RptColor = 15 'Color of report text
  
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(77, "-")   'dashed line
  DivLine2$ = STRING$(77, "=")  'Double Line
  CrLf$ = CHR$(13) + CHR$(10)
  
  Header$ = "Budget History"
  Desc$(1) = "Date       Description             Reference       Debit        Credit  Post Ref"
  
  TotDr# = 0
  TotCr# = 0
  OpenDate$ = ConvDateStr1(Num2Date$(GetBegDate))
  OpenDesc$ = "Balance Foward"
  
  ReportFile$ = Unique$(Path$)
  PrnFile = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  
  PrintHelp "   Processing:"
  
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  
  DIM BgtTrans AS GLTransRecType
  BgtTransLen = LEN(BgtTrans)
  BgtTransFile = FREEFILE
  OPEN "BGTTRANS.dat" FOR RANDOM ACCESS READ WRITE SHARED AS BgtTransFile LEN = BgtTransLen
  NumTrans = LOF(BgtTransFile) \ BgtTransLen
  
  IF NumTrans = 0 THEN
    CLOSE
    EXIT SUB
  END IF
  
  'trying to set TrSort array to max needed
  IGuess = 50

  DIM TrSort(1 TO IGuess) AS TrSortType
  
  GrTotDr# = 0
  GrTotCr# = 0
  
  FOR Cnt = 1 TO NumGLAccts
    
    DrFwd# = 0
    CrFwd# = 0
    TotAcctDr# = 0
    TotAcctCr# = 0
    
    GET AcctIdxFileNum, Cnt, AcctIdx
    GET AcctFileNum, AcctIdx.RecNum, Acct
    
    AcctNum$ = QPTrim$(Acct.Num)
    PrintHelp "Processing account: " + AcctNum$
    
    IF AcctNum$ >= FirstAcct$ AND AcctNum$ <= LastAcct$ THEN
      IF Acct.Typ = "R" OR Acct.Typ = "E" THEN
        
        NextTr = Acct.FrstBTran 'get the first trans for this acct
        AcctNumber$ = "Account " + AcctNum$ + " - " + QPTrim$(Acct.Title)
        
        ToPrint$ = SPACE$(80)
        LSET ToPrint$ = AcctNumber$
        PRINT #PrnFile, ToPrint$
        
        DO UNTIL NextTr = 0     'keep going 'til we run out of trans
          
          GET BgtTransFile, NextTr, BgtTrans
          
          IF BgtTrans.TrDate >= BegDate AND BgtTrans.TrDate <= EndDate THEN
            '--within range - assign to array for sorting
            NumBgtTrans = NumBgtTrans + 1
            TrSort(NumBgtTrans).TrDate = Trans.TrDate
            TrSort(NumBgtTrans).Record = NextTr
            
          ELSE
            '--check the transaction to see if we need to carry it in
            '  the balance fwd
            'QPRintRC Num2Date(BgtTrans.TrDate), 24, 25, -1
            
            IF Trans.TrDate < BegDate THEN
              DrFwd# = DrFwd# + BgtTrans.DrAmt
              CrFwd# = CrFwd# + BgtTrans.CrAmt
              FwdFlag = -1
            END IF
          END IF
          
          NextTr = BgtTrans.NextTran            'Get the next transaction
          
        LOOP
        
        IF FwdFlag THEN
          '--
          ToPrint$ = SPACE$(80)
          MID$(ToPrint$, 1) = "Balance Forward"
          
          SELECT CASE Acct.Typ
          CASE "E"
            BalFwd# = DrFwd# - CrFwd#
            IF BalFwd# >= 0 THEN
              Debit$ = FUsing$(STR$(BalFwd#), CommaFmt$)
              Credit$ = ""
            ELSE
              Credit$ = FUsing$(STR$(ABS(BalFwd#)), CommaFmt$)
              Debit$ = ""
            END IF
            
          CASE "R"
            BalFwd# = CrFwd# - DrFwd#
            IF BalFwd# >= 0 THEN
              Credit$ = FUsing$(STR$(BalFwd#), CommaFmt$)
              Debit$ = ""
            ELSE
              Debit$ = FUsing$(STR$(ABS(BalFwd#)), CommaFmt$)
              Credit$ = ""
            END IF
            
          END SELECT
          
          MID$(ToPrint$, 45) = Debit$
          MID$(ToPrint$, 59) = Credit$
          PRINT #PrnFile, ToPrint$
          
        END IF
        
        IF NumBgtTrans > 0 OR BalFwd# <> 0 THEN
          
          PrintHelp "Sorting Transactions..."
          
          SortT TrSort(1), NumAcctTrans, 0, 6, 0, -1
          
          PrintHelp "Writing to report file..."
          FOR Trn = 1 TO NumBgtTrans
            GET BgtTransFile, TrSort(Trn).Record, BgtTrans
            
            ToPrint$ = SPACE$(80)
            
            MID$(ToPrint$, 1) = ConvDateStr1(Num2Date(BgtTrans.TrDate))
            MID$(ToPrint$, 12) = BgtTrans.Desc
            MID$(ToPrint$, 36) = BgtTrans.Ref
            
            IF BgtTrans.DrAmt <> 0 THEN
              MID$(ToPrint$, 45) = FUsing$(STR$(BgtTrans.DrAmt), CommaFmt$)
            END IF
            
            IF BgtTrans.CrAmt <> 0 THEN
              MID$(ToPrint$, 59) = FUsing$(STR$(BgtTrans.CrAmt), CommaFmt$)
            END IF
            
            MID$(ToPrint$, 74) = LEFT$(BgtTrans.Src, 6)
            
            PRINT #PrnFile, ToPrint$
            
            TotAcctDr# = TotAcctDr# + BgtTrans.DrAmt
            TotAcctCr# = TotAcctCr# + BgtTrans.CrAmt
            
            GrTotDr# = GrTotDr# + BgtTrans.DrAmt
            GrTotCr# = GrTotCr# + BgtTrans.CrAmt
            
          NEXT
          
          '--Print summary lines
          ToPrint$ = SPACE$(80)
          MID$(ToPrint$, 44) = SumLine$
          MID$(ToPrint$, 58) = SumLine$
          PRINT #PrnFile, ToPrint$
          
          '--Print transaction totals
          IF NumAcctTrans > 0 THEN
            ToPrint$ = SPACE$(80)
            MID$(ToPrint$, 1) = "Transaction Totals"
            MID$(ToPrint$, 44) = FUsing$(STR$(TotAcctDr#), TotalFmt$)
            MID$(ToPrint$, 58) = FUsing$(STR$(TotAcctCr#), TotalFmt$)
            PRINT #PrnFile, ToPrint$
          END IF
          
          '--Print ending balance
          ToPrint$ = SPACE$(80)
          MID$(ToPrint$, 1) = "Budget Balance"
          SELECT CASE Acct.Typ
          CASE "E"
            AcctBal# = BalFwd# + TotAcctDr# - TotAcctCr#
            IF AcctBal# >= 0 THEN
              Debit$ = FUsing$(STR$(AcctBal#), TotalFmt$)
              Credit$ = ""
            ELSE
              Credit$ = FUsing$(STR$(ABS(AcctBal#)), TotalFmt$)
              Debit$ = ""
            END IF
            
          CASE "R"
            AcctBal# = BalFwd# + TotAcctCr# - TotAcctDr#
            IF AcctBal# >= 0 THEN
              Credit$ = FUsing$(STR$(AcctBal#), TotalFmt$)
              Debit$ = ""
            ELSE
              Debit$ = FUsing$(STR$(ABS(AcctBal#)), TotalFmt$)
              Credit$ = ""
            END IF
            
          END SELECT
          
          MID$(ToPrint$, 44) = Debit$
          MID$(ToPrint$, 58) = Credit$
          PRINT #PrnFile, ToPrint$
          
          ToPrint$ = STRING$(80, "*")
          PRINT #PrnFile, ToPrint$
          
        ELSE
          ToPrint$ = SPACE$(80)
          MID$(ToPrint$, 5) = " -- No Activity --"
          PRINT #PrnFile, ToPrint$
          
          ToPrint$ = STRING$(80, "*")
          PRINT #PrnFile, ToPrint$
          
        END IF
        
      END IF
    END IF      'Account is not of this fund
    
    NumBgtTrans = 0             'reset for next account
    BalFwd# = 0
    FwdFlag = 0
    TotAcctDr# = 0
    TotAcctCr# = 0
    
  NEXT
  
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "Grand Total Debits"
  MID$(ToPrint$, 25) = FUsing$(STR$(GrTotDr#), "##########,.##")
  PRINT #PrnFile, ToPrint$
  
  ToPrint$ = SPACE$(80)
  LSET ToPrint$ = "Grand Total Credits"
  MID$(ToPrint$, 25) = FUsing$(STR$(GrTotCr#), "##########,.##")
  PRINT #PrnFile, ToPrint$
  
  CLOSE
  
  SELECT CASE Dev$
  CASE "S"
    EntryPoint = 2
  CASE "P"
    EntryPoint = 5
  END SELECT

  PrintRptFile Header$, ReportFile$, LPTNo, RetCode%, EntryPoint

  'KILL ReportFile$
  
END SUB

SUB BgtTrList
  
  'prints the budget transaction file
  
  DIM BgtTrans AS GLTransRecType
  BgtTransLen = LEN(BgtTrans)
  BgtTransFile = FREEFILE
  OPEN "BGTTRANS.dat" FOR RANDOM ACCESS READ WRITE SHARED AS BgtTransFile LEN = BgtTransLen
  NumTrans = LOF(BgtTransFile) \ BgtTransLen
  
  'ReportFile$ = "BGTTR.PRN"
  ReportFile$ = Unique$(Path$)
  PrnFile = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  
  CommaFmt$ = "########,.##"
  
  FOR T = 1 TO NumTrans
    GET BgtTransFile, T, BgtTrans
    
    ToPrint$ = SPACE$(80)
    
    MID$(ToPrint$, 1) = ConvDateStr1(Num2Date(BgtTrans.TrDate))
    MID$(ToPrint$, 12) = BgtTrans.AcctNum
    MID$(ToPrint$, 28) = BgtTrans.Desc
    MID$(ToPrint$, 36) = BgtTrans.Ref
    IF BgtTrans.DrAmt <> 0 THEN
      MID$(ToPrint$, 45) = FUsing$(STR$(BgtTrans.DrAmt), CommaFmt$)
    END IF
    IF BgtTrans.CrAmt <> 0 THEN
      MID$(ToPrint$, 59) = FUsing$(STR$(BgtTrans.CrAmt), CommaFmt$)
    END IF
    MID$(ToPrint$, 74) = LEFT$(BgtTrans.Src, 6)
    PRINT #PrnFile, ToPrint$
    
    TotDr# = TotDr# + BgtTrans.DrAmt
    TotCr# = TotCr# + BgtTrans.CrAmt
    
  NEXT
  
  PRINT #PrnFile, "Debits  :" + FUsing$(STR$(TotDr#), CommaFmt$)
  PRINT #PrnFile, "Credits :" + FUsing$(STR$(TotCr#), CommaFmt$)
  
  CLOSE
  
END SUB

SUB CashBalance
  
  LOCATE , , 0
  SHARED User$
  LibName$ = "GL.QSL"           'screen library
  FormName$ = "PRNCASH"         'form to edit
  DevFld = 2    'Set to output device field number
  SHARED CashAcct$
  
  SHARED Choice$()
  REDIM Choice$(0 TO 3, 1)
  '--Choices for Output Device
  Choice$(0, 0) = STR$(DevFld)
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"
  
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  DIM FundIdx AS GLFundIndexType
  
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  RunFlag = 0
  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    SELECT CASE Frm.KeyCode
    CASE F10Key
      EndDate = Date2Num(Form$(1, 0))
      IF EndDate < -32000 THEN
        OK = MsgBox("GL.QSL", "BADDATE")
        RunFlag = 0
        Frm.FldNo = 1
        Action = 1
      ELSE
        RunFlag = -1
      END IF
    CASE ESC
      EXIT SUB
    END SELECT
    
  LOOP UNTIL RunFlag
  
  IF LEN(LTRIM$(RTRIM$(Form$(DevFld, 0)))) = 0 THEN
    Dev$ = "S"
    LPTNo = 1
  ELSE
    Dev$ = LEFT$(Form$(DevFld, 0), 1)
    LPTNo = VAL(RIGHT$(RTRIM$(Form$(DevFld, 0)), 1))
  END IF
  
  '=====================================================
  
  '--Clear help row
  QPRintRC SPACE$(80), 25, 1, -1
  
  PrnFileName$ = "CASHBAL.PRN"
  PrnFileNum = FREEFILE
  OPEN PrnFileName$ FOR OUTPUT AS #PrnFileNum
  
  OpenFundIdx FundIdxFile, NumFunds
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  
  PRINT #PrnFileNum, User$; TAB(43); "Run Date: " + DATE$
  PRINT #PrnFileNum, "Cash Balance Summary"
  PRINT #PrnFileNum, "Ending Date: " + Form$(1, 0)
  PRINT #PrnFileNum,
  PRINT #PrnFileNum, "Account                 Balance"
  PRINT #PrnFileNum, "-------------------------------"
  
  FOR Cnt = 1 TO NumFunds
    GET FundIdxFile, Cnt, FundIdx
    
    FundCashAcct$ = QPTrim$(FundIdx.FundNum) + CashAcct$
    AcctRec = FindAcct(FundCashAcct$)
    IF AcctRec > 0 THEN
      
      QPRintRC "Summarizing: " + FundCashAcct$, 25, 2, -1
      
      GET AcctFileNum, AcctRec, Acct
      CalcBal# = Round#(Acct.BegBal)
      NextTr& = Acct.FrstTran
      DO UNTIL NextTr& = 0
        GET TransFileNum, NextTr&, Trans
        IF Trans.TrDate <= EndDate THEN
          CalcBal# = Round#(CalcBal#) + Round#(Trans.DrAmt) - Round#(Trans.CrAmt)
        END IF
        NextTr& = Trans.NextTran
      LOOP
      
      PRINT #PrnFileNum, FundCashAcct$; TAB(20); FUsing(STR$(CalcBal#), "########,.##")
      TotalCashBal# = TotalCashBal# + CalcBal#
      
    END IF
  NEXT
  
  PRINT #PrnFileNum, TAB(18); "--------------"
  PRINT #PrnFileNum, "Total Cash"; TAB(17); FUsing(STR$(TotalCashBal#), "$$#########,.##")
  PRINT #PrnFileNum, CHR$(12)
  CLOSE
  
  SELECT CASE Dev$
  CASE "S"
    EntryPoint = 2
  CASE "P"
    EntryPoint = 5
  END SELECT
  
  PrintRptFile "Cash Balance Report", PrnFileName$, LPTNo, RetCode%, EntryPoint
  'KILL PrnFileName$
  
  EXIT SUB
  
END SUB

SUB FindAcctBal
  
  LOCATE , , 0
  SHARED User$
  LibName$ = "GL.QSL"           'screen library
  FormName$ = "PRNABAL"         'form to edit
  DevFld = 4    'Set to output device field number
  
  SHARED Choice$()
  REDIM Choice$(0 TO 3, 1)
  '--Choices for Output Device
  Choice$(0, 0) = STR$(DevFld)
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"
  
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  DIM FundIdx AS GLFundIndexType
  
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  RunFlag = 0
  Action = 1
  
  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    SELECT CASE Frm.KeyCode
    CASE F10Key
      EndDate = Date2Num(Form$(3, 0))
      IF EndDate < -32000 THEN
        OK = MsgBox("GL.QSL", "BADDATE")
        RunFlag = 0
        Frm.FldNo = 1
        Action = 1
      ELSE
        RunFlag = -1
      END IF
    CASE ESC
      EXIT SUB
    END SELECT
    
  LOOP UNTIL RunFlag
  
  IF LEN(LTRIM$(RTRIM$(Form$(DevFld, 0)))) = 0 THEN
    Dev$ = "S"
    LPTNo = 1
  ELSE
    Dev$ = LEFT$(Form$(DevFld, 0), 1)
    LPTNo = VAL(RIGHT$(RTRIM$(Form$(DevFld, 0)), 1))
  END IF
  
  '=====================================================
  
  '--Clear help row
  QPRintRC SPACE$(80), 25, 1, -1
  
  
  Acct2Bal$ = "-" + QPTrim$(Form$(1, 0)) + "-" + QPTrim$(Form$(2, 0))
  
  PrnFileName$ = "ACCTBAL.PRN"
  PrnFileNum = FREEFILE
  OPEN PrnFileName$ FOR OUTPUT AS #PrnFileNum
  
  OpenFundIdx FundIdxFile, NumFunds
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  
  PRINT #PrnFileNum, User$; TAB(43); "Run Date: " + DATE$
  PRINT #PrnFileNum, "Account Summary"
  PRINT #PrnFileNum, "Ending Date: " + Form$(3, 0)
  PRINT #PrnFileNum,
  PRINT #PrnFileNum, "Account                 Balance"
  PRINT #PrnFileNum, "-------------------------------"
  
  FOR Cnt = 1 TO NumFunds
    GET FundIdxFile, Cnt, FundIdx
    
    AcctNum$ = QPTrim$(FundIdx.FundNum) + Acct2Bal$
    AcctRec = FindAcct(AcctNum$)
    IF AcctRec > 0 THEN
      
      QPRintRC "Summarizing: " + AcctNum$, 25, 2, -1
      
      GET AcctFileNum, AcctRec, Acct
      CalcBal# = Round#(Acct.BegBal)
      NextTr& = Acct.FrstTran
      DO UNTIL NextTr& = 0
        GET TransFileNum, NextTr&, Trans
        IF Trans.TrDate <= EndDate THEN
          CalcBal# = Round#(CalcBal#) + Round#(Trans.DrAmt) - Round#(Trans.CrAmt)
        END IF
        NextTr& = Trans.NextTran
      LOOP
      
      PRINT #PrnFileNum, AcctNum$; TAB(20); FUsing(STR$(CalcBal#), "########,.##")
      TotalCashBal# = TotalCashBal# + CalcBal#
      
    END IF
  NEXT
  
  PRINT #PrnFileNum, TAB(18); "--------------"
  PRINT #PrnFileNum, "Total"; TAB(17); FUsing(STR$(TotalCashBal#), "$$#########,.##")
  PRINT #PrnFileNum, CHR$(12)
  CLOSE
  
  SELECT CASE Dev$
  CASE "S"
    EntryPoint = 2
  CASE "P"
    EntryPoint = 5
  END SELECT
  
  PrintRptFile "Cash Balance Report", PrnFileName$, LPTNo, RetCode%, EntryPoint
  'KILL PrnFileName$
  
  
  
END SUB

FUNCTION Num2Month% (Dt%)
  
  D$ = Num2Date(Dt%)
  m$ = RIGHT$(D$, 2) + LEFT$(D$, 2)
  Num2Month% = VAL(m$)
  
END FUNCTION

SUB TrlBal
  
  SHARED User$
  LibName$ = "GL.QSL"           'screen library
  FormName$ = "PRNTRBAL"        'form to edit
  DevFld = 4    'Set to output device field number
  
  SHARED Choice$()
  REDIM Choice$(0 TO 3, 1)
  Choice$(0, 0) = STR$(DevFld)  'Choices for Output Device
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"
  
  DIM FundIdx AS GLFundIndexType
  GetFundCodes FirstFund$, LastFund$
  
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  DIM Frm AS FormInfo
  
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  Action = 1
  
  DO            'Allow user to input report params
    RunFlag = 0
    
    EditForm Form$(), Fld(), Frm, Cnf, Action
    
    IF Frm.PrevFld <> Frm.FldNo THEN            'User is moving around in form
      
      SELECT CASE Frm.PrevFld
        
      CASE 1    'Ending Date
        EndDate = Date2Num(Form$(1, 0))
        IF EndDate < -32000 THEN
          OK = MsgBox("GL.QSL", "BADDATE")
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
              OK = MsgBox("GL.QSL", "NOFUND")
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
      
    CASE 42     'Asterisk  'I was using F7KEY
      Form$(2, 0) = FirstFund$
      Form$(3, 0) = LastFund$
      Form$(1, 0) = DATE$       'Num2Date(GetEndDate)
      Form$(4, 0) = "Screen"
      Frm.FldNo = DevFld
      Action = 1
      
    CASE F10Key 'SaveButton
      'check date
      EndDate = Date2Num(Form$(1, 0))
      IF EndDate < -32000 THEN
        OK = MsgBox("GL.QSL", "BADDATE")
        BadFld = 1
      END IF
      IF BadFld THEN
        Frm.FldNo = 1
        Action = 1
        GOTO ThruAgain
      ELSE
        RunFlag = -1
      END IF
      
      'check for valid funds
      FOR FieldNum = 2 TO 3
        IF LEN(QPTrim$(Form$(FieldNum, 0))) = 0 THEN
          OK = MsgBox("GL.QSL", "NOBLANK")
          Frm.FldNo = FieldNum
          BadFld = FieldNum
          EXIT FOR
        END IF
      NEXT
      IF BadFld THEN
        Action = 1
        RunFlag = 0
        GOTO ThruAgain
      ELSE
        RunFlag = -1
      END IF
      
      'make sure funds are in ascending order
      StartFund$ = QPTrim$(Form$(2, 0))
      EndFund$ = QPTrim$(Form$(3, 0))
      IF StartFund$ > EndFund$ THEN
        OK = MsgBox("GL.QSL", "RANGERR")
        BadFld = 1
      END IF
      IF BadFld THEN
        RunFlag = 0
        Frm.FldNo = 2
        Action = 1
        GOTO ThruAgain
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
    
ThruAgain:
    BadFld = 0
  LOOP UNTIL RunFlag
  
  LOCATE , , 0
  
  IF Dev$ = "F" THEN Dev$ = "S"
  
  'End of Input
  '=====================================================
  'Start Report Processing
  
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  
  IF EndDate >= FY2BegDate THEN
    ActiveYear = 2
    REDIM FundList$(1)
    GetFundList FundList$(), NumFunds
    REDIM PYFundRev#(NumFunds + 2)
    REDIM PYFundExp#(NumFunds + 2)
  ELSE
    ActiveYear = 1
  END IF
  
  ReportFile$ = "TRIALBAL.PRN"                'Report File Name
  
  Header$ = "Trial Balance"
  REDIM Desc$(1)
  Desc$(1) = "Acct Number     Title                                      Debit        Credit"
  CommaFmt$ = "########,.##"    'format takes 13 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(13, "-")   'column summary line
  DivLine$ = STRING$(79, "-")   'dashed line
  DivLine2$ = STRING$(79, "=")  'Double Line
  CrLf$ = CHR$(13) + CHR$(10)
  FF$ = CHR$(12)
  
  MaxLines = 55
  
  TotDr# = 0
  TotCr# = 0
  
  'ReportFile$ = Unique$(Path$)
  PrnFile = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  
  QPRintRC SPACE$(80), 25, 1, -1
  QPRintRC "Processing:", 25, 2, -1
  
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFileNum, NumFunds
  
  FOR Cnt = 1 TO NumGLAccts
    GET AcctIdxFileNum, Cnt, AcctIdx
    GET AcctFileNum, AcctIdx.RecNum, Acct
    
    FundCode$ = LEFT$(Acct.Num, FundLen)
    IF FundCode$ >= StartFund$ AND FundNumber$ <= EndFund$ THEN
      
      IF ActiveYear = 2 THEN
        '--Find the fund so we can calc prior year fund balance if necessary
        FOR FundCnt = 1 TO NumFunds
          IF FundCode$ = FundList$(FundCnt) THEN
            Found = True
            EXIT FOR
          END IF
        NEXT
      END IF
      
      CalcBal# = Round#(Acct.BegBal)            'get the beginning balance
      
      NextTr& = Acct.FrstTran   'get the first trans for this acct
      QPRintRC Acct.Num, 25, 14, -1
      
      DO UNTIL NextTr& = 0      'keep going 'til we run out of trans
        
        GET TransFileNum, NextTr&, Trans
        
        IF Trans.TrDate <= EndDate THEN
          
          SELECT CASE Acct.Typ
          CASE "A"              ', "E"
            CalcBal# = CalcBal# + Round#(Trans.DrAmt) - Round#(Trans.CrAmt)
            
          CASE "L"              ', "R"
            CalcBal# = CalcBal# + Round#(Trans.CrAmt) - Round#(Trans.DrAmt)
            
          CASE "R"
            'IF Trans.TrDate >= FY2BegDate THEN
            '  CalcBal# = CalcBal# + Round#(Trans.CrAmt) - Round#(Trans.DrAmt)
            'ELSE
            '  PYFundRev#(FundCnt) = PYFundRev#(FundCnt) + Round#(Trans.CrAmt) - Round#(Trans.DrAmt)
            'END IF
            SELECT CASE ActiveYear
            CASE 1
              CalcBal# = CalcBal# + Round#(Trans.CrAmt) - Round#(Trans.DrAmt)
            CASE 2
              IF Trans.TrDate < FY2BegDate THEN
                PYFundRev#(FundCnt) = PYFundRev#(FundCnt) + Round#(Trans.CrAmt) - Round#(Trans.DrAmt)
              ELSE
                CalcBal# = CalcBal# + Round#(Trans.CrAmt) - Round#(Trans.DrAmt)
              END IF
            END SELECT
            
          CASE "E"
            'IF Trans.TrDate >= FY2BegDate THEN
            '  CalcBal# = CalcBal# + Round#(Trans.DrAmt) - Round#(Trans.CrAmt)
            'ELSE
            '  PYFundExp#(FundCnt) = PYFundExp#(FundCnt) + Round#(Trans.DrAmt) - Round#(Trans.CrAmt)
            'END IF
            SELECT CASE ActiveYear
            CASE 1
              CalcBal# = CalcBal# + Round#(Trans.DrAmt) - Round#(Trans.CrAmt)
            CASE 2
              IF Trans.TrDate < FY2BegDate THEN
                PYFundExp#(FundCnt) = PYFundExp#(FundCnt) + Round#(Trans.DrAmt) - Round#(Trans.CrAmt)
              ELSE
                CalcBal# = CalcBal# + Round#(Trans.DrAmt) - Round#(Trans.CrAmt)
              END IF
            END SELECT
            
          END SELECT
          
        END IF
        
        NextTr& = Trans.NextTran                'Get the next transaction
        
      LOOP
      
      Acct.Bal = CalcBal#
      PUT AcctFileNum, AcctIdx.RecNum, Acct
      
    END IF      'test for account in fund range
    
  NEXT          'next account
  
  PrintHelp "Please wait..."
  
  TotDr# = 0: TotCr# = 0
  
  FOR Cnt = 1 TO NumFunds
    FundDr# = 0: FundCr# = 0
    GET FundIdxFileNum, Cnt, FundIdx
    FundNumber$ = QPTrim$(FundIdx.FundNum)
    IF FundNumber$ >= StartFund$ AND FundNumber$ <= EndFund$ THEN
      
      FundRecNum = FindFund(FundNumber$)
      FundName$ = QPTrim$(GetFundTitle$(FundRecNum))
      
      IF ActiveYear = 2 THEN
        'Get the fund count for printing prior year fund balance
        Found = False
        FOR FundCnt = 1 TO NumFunds
          IF FundNumber$ = FundList$(FundCnt) THEN
            Found = True
            EXIT FOR
          END IF
        NEXT
      END IF
      
      GOSUB PrintPageHeader
      
      FOR RecNo = 1 TO NumGLAccts               'Active Accts
        
        GET AcctIdxFileNum, RecNo, AcctIdx
        
        FundCode$ = LEFT$(AcctIdx.AcctNum, FundLen)
        IF FundCode$ = FundNumber$ THEN
          GET AcctFileNum, AcctIdx.RecNum, Acct
          
          LineCnt = LineCnt + 1
          IF LineCnt >= MaxLines THEN
            PRINT #PrnFile, FF$
            GOSUB PrintPageHeader
          END IF
          
          ToPrint$ = SPACE$(80)
          LSET ToPrint$ = Acct.Num
          MID$(ToPrint$, 17) = Acct.Title
          
          SELECT CASE Acct.Typ
          CASE "A", "E"
            IF Acct.Bal >= 0 THEN
              Debit$ = FUsing$(STR$(Acct.Bal), CommaFmt$)
              TotDr# = TotDr# + Acct.Bal
              FundDr# = FundDr# + Acct.Bal
              Credit$ = ""
            ELSE
              Credit$ = FUsing$(STR$(ABS(Acct.Bal)), CommaFmt$)
              TotCr# = TotCr# + ABS(Acct.Bal)
              FundCr# = FundCr# + ABS(Acct.Bal)
              Debit$ = ""
            END IF
            
          CASE "L", "R"
            IF Acct.Bal >= 0 THEN
              Credit$ = FUsing$(STR$(Acct.Bal), CommaFmt$)
              TotCr# = TotCr# + Acct.Bal
              FundCr# = FundCr# + Acct.Bal
              Debit$ = ""
            ELSE
              Debit$ = FUsing$(STR$(ABS(Acct.Bal)), CommaFmt$)
              TotDr# = TotDr# + ABS(Acct.Bal)
              FundDr# = FundDr# + ABS(Acct.Bal)
              Credit$ = ""
            END IF
            
          END SELECT
          
          MID$(ToPrint$, 53) = Debit$
          MID$(ToPrint$, 67) = Credit$
          PRINT #PrnFile, ToPrint$
        END IF
        
      NEXT
      
      ToPrint$ = SPACE$(80)
      MID$(ToPrint$, 52) = SumLine$
      MID$(ToPrint$, 66) = SumLine$
      PRINT #PrnFile, ToPrint$
      
      ToPrint$ = SPACE$(80)
      LSET ToPrint$ = FundName$ + " " + "Totals"
      MID$(ToPrint$, 52) = FUsing$(STR$(FundDr#), TotalFmt$)
      MID$(ToPrint$, 66) = FUsing$(STR$(FundCr#), TotalFmt$)
      PRINT #PrnFile, ToPrint$
      
      Diff# = Round#(FundDr#) - Round#(FundCr#)
      IF Diff# <> 0 THEN
        ToPrint$ = SPACE$(80)
        LSET ToPrint$ = "Fund is out of balance :"
        MID$(ToPrint$, 25) = FUsing$(STR$(Diff#), CommaFmt$)
        PRINT #PrnFile, ToPrint$
      END IF
      
      IF ActiveYear = 2 THEN
        PYFundBal# = 0
        PYFundBal# = PYFundRev#(FundCnt) - PYFundExp#(FundCnt)
        ToPrint$ = SPACE$(80)
        LSET ToPrint$ = "Prior Year unaudited fund balance adjustment:"
        MID$(ToPrint$, 41) = FUsing$(STR$(PYFundBal#), CommaFmt$)
        PRINT #PrnFile, ToPrint$
      END IF
      
      PRINT #PrnFile,
      PRINT #PrnFile, FF$
      
    END IF      'if fund is in range test
  NEXT          'next fund
  
  ToPrint$ = SPACE$(80)
  
  '--Print a grand total if more than one fund
  IF StartFund$ <> EndFund$ THEN
    PRINT #PrnFile, "Combined totals - All Funds"
    PRINT #PrnFile, "Total Debits  : " + FUsing$(STR$(TotDr#), TotalFmt$)
    PRINT #PrnFile, "Total Credits : " + FUsing$(STR$(TotCr#), TotalFmt$)
    PRINT #PrnFile, CHR$(12)
  END IF
  
  CLOSE
  
  
  'End Report Processing
  
  SELECT CASE Dev$
  CASE "S"
    EntryPoint = 2
  CASE "P"
    EntryPoint = 5
  END SELECT
  
  PrintRptFile Header$, ReportFile$, LPTNo, RetCode%, EntryPoint
  'KILL ReportFile$
  
  EXIT SUB
  
PrintPageHeader:
  PRINT #PrnFile, User$; TAB(43); "Run Date: " + DATE$
  PRINT #PrnFile, FundName$ + " " + Header$
  PRINT #PrnFile, "Period Ending: " + Form$(1, 0)
  PRINT #PrnFile,
  PRINT #PrnFile, Desc$(1)
  PRINT #PrnFile, DivLine$
  LineCnt = 6
  RETURN
  
  
END SUB

