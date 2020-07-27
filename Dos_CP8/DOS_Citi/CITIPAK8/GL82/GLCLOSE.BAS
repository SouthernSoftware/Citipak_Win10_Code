'GLClose.BAS by Jim Brown
'see comments sub for more info

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
DECLARE SUB QPrintRC (T$, R%, c%, Clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
DECLARE SUB KillFile (FileName$)
DECLARE SUB PickList (Items() AS ANY, Picked%(), NPicked%, Cnf AS ANY)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)



DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
'DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB MnuTitle (Title$, Row, Col, Wdth, TxtColor)
DECLARE SUB ClrBG ()
DECLARE SUB GetFundCodes (FirstFund$, LastFund$)
DECLARE SUB GetFundCloseList (FundCloseList$(), NumFundstoClose%)
DECLARE SUB OpenDeptFile (DeptFile%, NumDepts%)
DECLARE SUB BgtHist ()
DECLARE FUNCTION GetPct$ (N1#, N2#)
DECLARE FUNCTION GetDeptTitle$ (DeptRecNum%)
DECLARE FUNCTION FindDept% (DeptNum$)
DECLARE FUNCTION Num2Month% (Dt%)
DECLARE SUB GetFundList (FundList$(), NumFunds%)
DECLARE SUB GetFYDates (FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate)
DECLARE FUNCTION GetDeptOffsets$ ()
DECLARE SUB CloseTrialBal ()
DECLARE SUB CheckClosingEntry ()
DECLARE SUB BalSheet ()
DECLARE SUB SplitTransFile ()
DECLARE SUB GetFundsToClose ()
DECLARE SUB PostAdj (Unposted)

DECLARE SUB ReLinkTrans ()
DECLARE SUB RelinkBgtTrans ()

DECLARE SUB UpDateAcctFile ()
DECLARE SUB MakeOpenEntries ()
DECLARE SUB UpDateNewYear ()
DECLARE SUB RepostNewYearTrans ()
DECLARE SUB ResetYears ()


CONST False = 0, True = NOT False
  
'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'GL.bi'                        'GL File Types
'$INCLUDE: 'SCRAUX.BI'
'$INCLUDE: 'QSCR.BI'                      'QuickScreen Declarations

TYPE PickListArray
  Item AS STRING * 64
END TYPE

DIM SHARED TotLen%, FundLen%, AcctLen%, DetLen%
DIM SHARED Acct AS GLAcctRecType
DIM SHARED AcctIdx AS GLAcctIndexType
DIM SHARED Trans AS GLTransRecType
DIM SHARED User$, FBAcct$
  
'$INCLUDE: 'SetCnf.bi'        'Set monitor and color information
'$INCLUDE: 'init.bi'

LOCATE 1, 1, 0
TextCursor -2, -2             'Define the Mouse cursor
ShowCursor    'Show it

REDIM MChoice$(1 TO 9)
MChoice$(1) = "Select Funds to Close"
MChoice$(2) = "Pre-Closing Operations"
MChoice$(3) = "Close Fiscal Year"
MChoice$(4) = "Quit"

MaxLen = 0  'Set menu width to zero
BoxBot = 17   'limit the box length to go no lower than line 20
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

'GOTO Adios

DO

  LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, Attribute, ErrorCode

  Title$ = "Close F/Y "
  MnuTitle Title$, Row, Col, MaxLen, 15
  PrintTitle User$
  PrintHelp Help$

  LOCATE Row, Col, 0
  VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

  IF Ky$ = CHR$(27) THEN EXIT DO

  SELECT CASE Choice
  CASE 1
    ClrBG
    GetFundsToClose
  CASE 2
    ClrBG
    IF Exist("FCLOSE.LST") THEN
      CheckClosingEntry
    ELSE
      CLS
      BEEP
      QPrintRC "Select the funds to close first!", 10, 24, 11
      QPrintRC "   Press any key to continue.", 11, 24, 15
      K$ = INPUT$(1)
    END IF
  CASE 3
    LOCATE 1, 1, 0
    ClrBG
    IF Exist("FBADJ.DAT") THEN
      OK = MsgBox%("GL.QSL", "OKCLOSE")
      IF OK = 2 THEN
        PostAdj (Unposted)
        QPrintRC SPACE$(80), 1, 1, 112
        MakeOpenEntries

        QPrintRC SPACE$(80), 1, 1, 112
        SplitTransFile

        QPrintRC SPACE$(80), 1, 1, 112
        UpDateNewYear

        LibFile2Scrn "GL.QSL", "OPS", MonoCode, Attribute, ErrorCode

        RelinkBgtTrans

        RepostNewYearTrans

        ResetYears

        LibFile2Scrn "GL.QSL", "OPS", MonoCode, Attribute, ErrorCode
        QPrintRC "Closing Operations Complete", 5, 10, 15
        IF Unposted > 0 THEN
          QPrintRC "Error - Unposted Adjusting Entries: " + STR$(Unposted), 6, 10, 15
        END IF
        QPrintRC "Press any key to exit", 10, 10, 14
        K$ = INPUT$(1)
        EXIT DO
      END IF
    ELSE
      CLS
      BEEP
      QPrintRC "Run Pre-Closing Operations First!", 10, 24, 11
      QPrintRC "   Press any key to continue.", 11, 24, 15
      K$ = INPUT$(1)
    END IF
  CASE 4
    ClrBG
    HideCursor                'Be nice and turn the mouse cursor Off
    CLS
    END
  END SELECT
LOOP

ClrBG
HideCursor
DispLoad
RUN "glmaint"

SUB BalSheet

  'SHARED User$
  'LibName$ = "GL.QSL"           'screen library
  'FormName$ = "PRNBALSH"        'form to edit
  '
  'DevFld = 4    'Set to output device field number
  '
  'SHARED Choice$()
  '
  'REDIM Form$(0, 0)
  'REDIM Fld(0) AS FieldInfo
  'REDIM Choice$(0 TO 3, 1)
  'DIM Frm AS FormInfo
  'Choice$(0, 0) = STR$(DevFld)  'Choices for Output Device
  'Choice$(1, 0) = "Screen"
  'Choice$(2, 0) = "Prn#1"
  'Choice$(3, 0) = "Prn#2"

  DIM FundIdx AS GLFundIndexType
  GetFundCodes FirstFund$, LastFund$
  
  'LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  '
  'Action = 1
  '
  'DO            'Allow user to input report params
  '  RunFlag = 0
  '
  '  EditForm Form$(), Fld(), Frm, Cnf, Action
  '
  '  IF Frm.PrevFld <> Frm.FldNo THEN            'User is moving around in form
  '
  '    SELECT CASE Frm.PrevFld
  '
  '    CASE 1    'Ending Date
  '      EndDate = Date2Num(Form$(1, 0))
  '      IF EndDate < -32000 THEN
  '        OK = MsgBox("GL.QSL", "BADDATE")
  '        Frm.FldNo = 1
  '      END IF
  '
  '    CASE 2, 3 'Fund Field
  '      Fld = Frm.PrevFld
  '      IF Frm.FldEdited THEN   'If the field was edited
  '        Form$(Fld, 0) = QPTrim$(Form$(Fld, 0))
  '        SaveField Fld, Form$(), Fld(), BadFld
  '
  '        IF LEN(Form$(Fld, 0)) <> FundLen THEN
  '          Form$(Fld, 0) = ""
  '          Frm.FldNo = Fld
  '
  '        ELSE
  '          GoodFund = FindFund(Form$(Fld, 0))
  '          IF GoodFund = 0 THEN
  '            OK = MsgBox("GL.QSL", "NOFUND")
  '            Form$(Fld, 0) = ""
  '            Frm.FldNo = Fld
  '            BadFld = -1
  '          END IF
  '
  '        END IF
  '
  '      END IF
  '
  '    CASE ELSE
  '      'what else?
  '    END SELECT
  '
  '    Action = 1
  '
  '  END IF
  '
  '  SELECT CASE Frm.KeyCode     'Check the user's key presses
  '
  '  CASE 42
  '    Form$(2, 0) = FirstFund$
  '    Form$(3, 0) = LastFund$
  '    Form$(1, 0) = Num2Date(GetEndDate)
  '    Form$(4, 0) = "Screen"
  '    Frm.FldNo = DevFld
  '    Action = 1
  '
  '  CASE F10Key 'SaveButton
  '    'check date
  '    EndDate = Date2Num(Form$(1, 0))
  '    IF EndDate < -32000 THEN
  '      OK = MsgBox("GL.QSL", "BADDATE")
  '      BadFld = 1
  '    END IF
  '    IF BadFld THEN
  '      Frm.FldNo = 1
  '      Action = 1
  '      GOTO Again
  '    ELSE
  '      RunFlag = -1
  '    END IF
  '
  '    'check for valid funds
  '    FOR FieldNum = 2 TO 3
  '      IF LEN(QPTrim$(Form$(FieldNum, 0))) = 0 THEN
  '        OK = MsgBox("GL.QSL", "NOBLANK")
  '        Frm.FldNo = FieldNum
  '        BadFld = FieldNum
  '        EXIT FOR
  '      END IF
  '    NEXT
  '    IF BadFld THEN
  '      Action = 1
  '      RunFlag = 0
  '      GOTO Again
  '    ELSE
  '      RunFlag = -1
  '    END IF
  '
  '    'make sure funds are in ascending order
  '    StartFund$ = QPTrim$(Form$(2, 0))
  '    EndFund$ = QPTrim$(Form$(3, 0))
  '    IF StartFund$ > EndFund$ THEN
  '      OK = MsgBox("GL.QSL", "RANGERR")
  '      BadFld = 1
  '    END IF
  '    IF BadFld THEN
  '      RunFlag = 0
  '      Frm.FldNo = 2
  '      Action = 1
  '      GOTO Again
  '    ELSE
  '      RunFlag = -1
  '    END IF
  '
  '    IF LEN(LTRIM$(RTRIM$(Form$(DevFld, 0)))) = 0 THEN
  '      Dev$ = "S"
  '      LPTNo = 1
  '    ELSE
  '      Dev$ = LEFT$(Form$(DevFld, 0), 1)
  '      LPTNo = VAL(RIGHT$(RTRIM$(Form$(DevFld, 0)), 1))
  '    END IF
  '
  '
  '  CASE ESC
  '    EXIT SUB
  '
  '  END SELECT
  '
'Again:
'    BadFld = 0
'  LOOP UNTIL RunFlag
'
'  LOCATE , , 0
'
'  IF Dev$ = "F" THEN Dev$ = "S"
'
'
'  'End of Input
'  '=====================================================
  'Start Report Processing
  
  'REDIM Desc$(1)

  'ReportFile$ = "BALSHEET.PRN"  'Report File Name
  'CommaFmt$ = "########,.##"    'format takes 13 chars
  'TotalFmt$ = "#########,.##"   'format takes 14 chars
  'SumLine$ = STRING$(13, "-")   'column summary line
  'DivLine$ = STRING$(79, "-")   'dashed line
  'DivLine2$ = STRING$(79, "=")  'Double Line
  'CrLf$ = CHR$(13) + CHR$(10)
  'PageBreak$ = CHR$(12)

  'RptTitle$ = "Balance Sheet"
  'Desc$(1) = "Acct Number     Title                                                Balance"
  'Desc$ = "Acct Number     Title                                                   Balance"

  TotDr# = 0
  TotCr# = 0
  'MaxLines = 55

  'PrnFile = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFile, NumFunds
  
  REDIM Asset%(1 TO NumGLAccts)          'Holds all Asset acct record nums
  REDIM Liab%(1 TO NumGLAccts)           'Holds all Liab acct record nums
  REDIM FundList$(1 TO NumFunds)         'List of all active Funds
  REDIM FundTotAssets#(1 TO NumFunds)    'List of total assets by fund
  REDIM FundTotLiab#(1 TO NumFunds)      'list of total liab by fund
  REDIM FundTotRev#(1 TO NumFunds)       'List of total revenues by fund
  REDIM FundTotExp#(1 TO NumFunds)       'list of tot exp by fund
  
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
      QPrintRC Acct.Num, 25, 14, -1

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
          
          EXIT FOR               'acct was a match to this fund so exit loop
        END IF                   'test for fund in list
      NEXT                       'Summarize next fund's totals
    END IF                       'End of fund in range test
  NEXT                           'Process next account
  
  
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
        
      NEXT  'AssetAcct
      
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
      MID$(ToPrint$, 68) = FUsing$(STR$(FundAsset#), CommaFmt$) '+ " " + DrCr$
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
        
      NEXT  'LiabAcct
      
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
      MID$(ToPrint$, 68) = FUsing$(STR$(TotLiabnCap#), CommaFmt$)               ' + " " + DrCr$
      PRINT #PrnFile, ToPrint$

      PRINT #PrnFile, PageBreak$
      

   END IF      'using fund test
    
  NEXT          'FundList$()
  

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
  'KILL ReportFile$
  
EXIT SUB

PrintBSHeader:
   PRINT #PrnFile, User$
   PRINT #PrnFile, QPTrim$(FundName$) + " " + RptTitle$
   PRINT #PrnFile, AcctType$
   PRINT #PrnFile,
   PRINT #PrnFile, 'Desc$(1)
   PRINT #PrnFile, DivLine$
   LineCnt = 6
RETURN
    
BreakBSPage:
   PRINT #PrnFile, PageBreak$
   LineCnt = 0
RETURN
    

END SUB

SUB CheckClosingEntry

  '--call close trial bal to calculate ending account balances
  'also copies ending acct balances to Prior Year Actual field
  CloseTrialBal

  SHARED User$, FBAcct$
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  Fmt$ = "########,.##"
  
  REDIM Accts(1) AS GLAcctRecType
  REDIM AcctIndex(1) AS GLAcctIndexType
  
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFileNum, NumFunds

  REDIM FundList$(1)                     'List of all active Funds
  GetFundList FundList$(), NumFunds

  REDIM FundCloseList$(1)                'List of Funds to close
  GetFundCloseList FundCloseList$(), NumFundstoClose
  
  FundBalAdjFileName$ = "FBADJ.DAT"
  IF Exist(FundBalAdjFileName$) THEN
     KILL FundBalAdjFileName$
  END IF

  DIM GLFBAdj AS GLFBAdjRecType
  FBARecLen = LEN(GLFBAdj)
  FundBalAdjFile = FREEFILE
  OPEN FundBalAdjFileName$ FOR RANDOM AS FundBalAdjFile LEN = FBARecLen
  
  REDIM FundTotRev#(1 TO NumFunds)       'List of total revenues by fund
  REDIM FundTotExp#(1 TO NumFunds)       'list of tot exp by fund
  REDIM FundBalAdj#(1 TO NumFunds)
  
  PrnFile = FREEFILE
  ReportFile$ = "GLCLOSE.PRN"
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  
  FOR Cnt = 1 TO NumGLAccts   'NumGLAccts
    GET AcctIdxFileNum, Cnt, AcctIndex(1)
    GET AcctFileNum, AcctIndex(1).RecNum, Accts(1)

    FundCode$ = LEFT$(Accts(1).Num, FundLen)

    ClosingThisFund = False
    FOR F = 1 TO NumFundstoClose
      IF FundCode$ = QPTrim$(FundCloseList$(F)) THEN
        ClosingThisFund = True
        EXIT FOR
      END IF
    NEXT

    IF ClosingThisFund THEN
       
      '--find position in list
      FOR F = 1 TO NumFunds
        IF FundCode$ = QPTrim$(FundList$(F)) THEN
          EXIT FOR
        END IF
      NEXT

      IF F = 0 THEN
        CLS
        QPrintRC "Fatal Error Code: TIHSHO", 1, 1, 12
        CLOSE : END
      END IF

      '--Apply balance to proper fund
      SELECT CASE Accts(1).Typ
        'CASE "A"
        '  IF Accts(1).Bal >= 0 THEN
        '    'Debit
        '    DebitAmt# = Accts(1).Bal
        '    CreditAmt# = 0
        '  ELSE
        '    DebitAmt# = 0
        '    CreditAmt# = ABS(Accts(1).Bal)
        '  END IF
        '  GOSUB WriteToInterfaceFile
        'CASE "L"
        '  IF Accts(1).Bal >= 0 THEN
        '    DebitAmt# = 0
        '    CreditAmt# = Accts(1).Bal
        '  ELSE
        '    DebitAmt# = ABS(Accts(1).Bal)
        '    CreditAmt# = 0
        '  END IF
        '  GOSUB WriteToInterfaceFile
        CASE "R"
          FundTotRev#(F) = FundTotRev#(F) + Accts(1).Bal
        CASE "E"
          FundTotExp#(F) = FundTotExp#(F) + Accts(1).Bal
       END SELECT

    END IF
  NEXT

  GOSUB PrnHeader
  FOR F = 1 TO NumFunds
    'IF FundTotRev#(F) <> 0 AND FundTotExp#(F) <> 0 THEN
    IF FundTotRev#(F) <> 0 OR FundTotExp#(F) <> 0 THEN
      FundBalAdj#(F) = FundTotRev#(F) - FundTotExp#(F)
      IF FundBalAdj#(F) < 0 THEN
        DC$ = " Debit"
      ELSE
        DC$ = " Credit"
      END IF
      AdjAmt# = ABS(FundBalAdj#(F))
      FundBalAcct$ = QPTrim$(FundList$(F)) + FBAcct$
      GoodAcct = 0
      GoodAcct = FindAcct(FundBalAcct$)
      IF GoodAcct = 0 THEN
        B$ = " *"
        MsgFlag = True
      ELSE
        B$ = "  "
      END IF
      PRINT #PrnFile, FundBalAcct$ + B$;
      PRINT #PrnFile, TAB(15); USING Fmt$; AdjAmt#;
      PRINT #PrnFile, DC$
      '--Write out adjustments
      RecCnt = RecCnt + 1
      GLFBAdj.AcctNum = FundBalAcct$
      GLFBAdj.AdjAmt = FundBalAdj#(F)
      PUT FundBalAdjFile, RecCnt, GLFBAdj
    END IF
  NEXT

  IF MsgFlag THEN
    PRINT #PrnFile,
    PRINT #PrnFile, "* Account Does not Exist!"
    PRINT #PrnFile, "Setup account before closing!"
  END IF
  PRINT #PrnFile, CHR$(12)

  CLOSE

  PrintRptFile "G/L Closing Entries", ReportFile$, LPTNo, RetCode%, EntryPoint

EXIT SUB

PrnHeader:
  PRINT #PrnFile, User$
  PRINT #PrnFile, "Fund Balance Closing Entries"
  PRINT #PrnFile, "Fiscal Year Ending "; Num2Date(FY1EndDate)
  PRINT #PrnFile,
  PRINT #PrnFile, "Fund                Amount  Entry"
  PRINT #PrnFile, "---------------------------------"
RETURN

END SUB

SUB CloseTrialBal
'*********************************************************************
'calculates account balances for each fund in the fund close out list
'and sets the p/y actual field to the ending balance
'prints closing trial balance report to disk file
'***********************************************************************

  SHARED User$
  LibFile2Scrn "GL.QSL", "OPS", MonoCode, Attribute, ErrorCode
  QPrintRC "Generating Closing Trial Balance Report.  Please wait.", 5, 10, 15

  DIM FundIdx AS GLFundIndexType
  GetFundCodes FirstFund$, LastFund$
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  ReportFile$ = "CLOSETB.PRN"                'Report File Name
  Header$ = "Closing Trial Balance"
  REDIM Desc$(1)
  Desc$(1) = "Acct Number     Title                                      Debit        Credit"
  CommaFmt$ = "########,.##"  'format takes 13 chars
  TotalFmt$ = "##########,.##" 'format takes 14 chars
  SumLine$ = STRING$(13, "-") 'column summary line
  DivLine$ = STRING$(79, "-") 'dashed line
  DivLine2$ = STRING$(79, "=")                'Double Line
  CrLf$ = CHR$(13) + CHR$(10)
  FF$ = CHR$(12)
  MaxLines = 55
  TotDr# = 0
  TotCr# = 0

  REDIM FundCloseList$(1)
  GetFundCloseList FundCloseList$(), NumFundstoClose

  PrnFile = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile

  QPrintRC SPACE$(80), 25, 1, -1
  QPrintRC "Processing:", 25, 2, -1

  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFileNum, NumFunds

  FOR Cnt = 1 TO NumGLAccts   'NumGLAccts
    GET AcctIdxFileNum, Cnt, AcctIdx
    GET AcctFileNum, AcctIdx.RecNum, Acct

    FundCode$ = LEFT$(Acct.Num, FundLen)
    ClosingThisFund = False
    FOR F = 1 TO NumFundstoClose
      IF FundCode$ = QPTrim$(FundCloseList$(F)) THEN
        ClosingThisFund = True
        EXIT FOR
      END IF
    NEXT

    QPrintRC Acct.Num, 25, 14, -1

    IF ClosingThisFund THEN

      CalcBal# = Round#(Acct.BegBal)            'get the beginning balance

      NextTr& = Acct.FrstTran   'get the first trans for this acct
      
      DO UNTIL NextTr& = 0      'keep going 'til we run out of trans

        GET TransFileNum, NextTr&, Trans

          IF Trans.TrDate <= FY1EndDate THEN

            SELECT CASE Acct.Typ
              CASE "A", "E"
                CalcBal# = CalcBal# + Round#(Trans.DrAmt) - Round#(Trans.CrAmt)
              CASE "L", "R"
                CalcBal# = CalcBal# + Round#(Trans.CrAmt) - Round#(Trans.DrAmt)
            END SELECT

          END IF
          
          NextTr& = Trans.NextTran                'Get the next transaction

      LOOP

      Acct.Bal = CalcBal#
      Acct.Work = CalcBal# 'hold this
      PUT AcctFileNum, AcctIdx.RecNum, Acct

    END IF   'test for account in fund range

  NEXT       'next account

  PrintHelp "Please wait..."

  TotDr# = 0: TotCr# = 0

  FOR Cnt = 1 TO NumFunds
    FundDr# = 0: FundCr# = 0
    GET FundIdxFileNum, Cnt, FundIdx

    FundNumber$ = QPTrim$(FundIdx.FundNum)
    ClosingThisFund = False
    FOR F = 1 TO NumFundstoClose
      IF FundNumber$ = QPTrim$(FundCloseList$(F)) THEN
        ClosingThisFund = True
        EXIT FOR
      END IF
    NEXT

    IF ClosingThisFund THEN

      FundRecNum = FindFund(FundNumber$)
      FundName$ = QPTrim$(GetFundTitle$(FundRecNum))

      GOSUB PrintPageHeader

      FOR RecNo = 1 TO NumGLAccts             'Active Accts

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

      PRINT #PrnFile,
      PRINT #PrnFile, FF$

    ELSE
      'were not closing this fund
      FundRecNum = FindFund(FundNumber$)
      FundName$ = QPTrim$(GetFundTitle$(FundRecNum))
      GOSUB PrintPageHeader
      PRINT #PrnFile, "Fund " + FundNumber$ + " " + FundName$ + " is not being closed."
      PRINT #PrnFile, FF$

    END IF    'if fund is in range test
  NEXT        'next fund

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
   PRINT #PrnFile, User$
   PRINT #PrnFile, FundName$ + " " + Header$
   PRINT #PrnFile, "Period Ending: " + Num2Date(FY1EndDate)
   PRINT #PrnFile,
   PRINT #PrnFile, Desc$(1)
   PRINT #PrnFile, DivLine$
   LineCnt = 6
RETURN


END SUB

SUB comments

'last update 08/09/96

END SUB

SUB GetFundCloseList (FundCloseList$(), NumFundstoClose)
  
  DIM FundsToClose AS GLFundCloseRecType
  CloseListFileName$ = "FCLOSE.LST"
  RecLen = LEN(FundsToClose)
  FundCloseListFile = FREEFILE
  OPEN CloseListFileName$ FOR RANDOM AS FundCloseListFile LEN = RecLen
  NumFundstoClose = LOF(FundCloseListFile) \ RecLen
  
  IF NumFundstoClose = 0 THEN EXIT SUB

  REDIM FundCloseList$(1 TO NumFundstoClose)
  
  FOR R = 1 TO NumFundstoClose
    GET FundCloseListFile, R, FundsToClose
    FundCloseList$(R) = FundsToClose.FundNum
  NEXT

  CLOSE FundCloseListFile
  
END SUB

SUB GetFundsToClose

  'DIM FundList$(1)
  'GetFundList FundList$(), NumFunds
  'REDIM Funds2Close$(NumFunds)
  'FOR F = 1 TO NumFunds
  '  Funds2Close$(F) = FundList$(F)
  'NEXT
  'GLFundCloseRecType

  QPrintRC "  F10 To Continue, SpaceBar To Toggle, Alt-M Marks All, Alt-C Clear All Marks", 25, 1, -1

  DIM FundIdx AS GLFundIndexType
  DIM Fund AS GLFundRecType
  OpenFundIdx FundIdxFileNum, NumFIdxRecs
  OpenFundFile FundFileNum, NumFunds
  OpenFundIdx FundIdxFileNum, NumFIdxRecs
  REDIM FundCloseList(NumFIdxRecs) AS PickListArray
  REDIM Picked(NumFIdxRecs) 'AS PickListArray
  DIM FundsToClose AS GLFundCloseRecType
  
  FundName$ = SPACE$(64)

  FOR F = 1 TO NumFIdxRecs
    GET FundIdxFileNum, F, FundIdx
    GET FundFileNum, FundIdx.RecNum, Fund
    LSET FundName$ = QPTrim$(FundIdx.FundNum) + " " + QPTrim$(Fund.Title)
    FundCloseList(F).Item = FundName$
    LSET FundName$ = ""
  NEXT

  CLOSE
  
  PickList FundCloseList(), Picked(), NPicked, Cnf

  '--process the pick list into a file of only the selected choices
  CloseListFileName$ = "FCLOSE.LST"
  KillFile CloseListFileName$
  FundCloseListFile = FREEFILE
  OPEN CloseListFileName$ FOR RANDOM AS FundCloseListFile LEN = LEN(FundsToClose)

  FOR PCnt = 1 TO NPicked
    PF = Picked(PCnt)
    FundsToClose.FundNum = LEFT$(FundCloseList(PF).Item, FundLen)
    PUT FundCloseListFile, PCnt, FundsToClose
  NEXT
  
  CLOSE

END SUB

SUB MakeOpenEntries
  
  LibFile2Scrn "GL.QSL", "OPS", MonoCode, Attribute, ErrorCode
  ToPrint$ = "Creating Opening Entries"
  QPrintRC ToPrint$, 5, 10, 15
  
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  Fmt$ = "#########,.##"
  FF$ = CHR$(12)

  REDIM Accts(1) AS GLAcctRecType
  REDIM AcctIndex(1) AS GLAcctIndexType
  REDIM OpenEntry(1) AS GLTransRecType

  TransRecLen = LEN(OpenEntry(1))
  OEFileNum = FREEFILE
  OPEN "OPNENTRY.DAT" FOR RANDOM AS OEFileNum LEN = TransRecLen
  'NumTrans& = LOF(TransFileNum) \ TransRecLen
  
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs

  'REDIM FundList$(1)                     'List of all active Funds
  'GetFundList FundList$(), NumFunds

  REDIM FundCloseList$(1)                'List of Funds to close
  GetFundCloseList FundCloseList$(), NumFundstoClose

  'FundBalAdjFileName$ = "FBADJ.DAT"
  'IF Exist(FundBalAdjFileName$) THEN
  '   KILL FundBalAdjFileName$
  'END IF

  'DIM GLFBAdj AS GLFBAdjRecType
  'FBARecLen = LEN(GLFBAdj)
  'FundBalAdjFile = FREEFILE
  'OPEN FundBalAdjFileName$ FOR RANDOM AS FundBalAdjFile LEN = FBARecLen

  'REDIM FundTotRev#(1 TO NumFunds)       'List of total revenues by fund
  'REDIM FundTotExp#(1 TO NumFunds)       'list of tot exp by fund
  'REDIM FundBalAdj#(1 TO NumFunds)

  PrnFile = FREEFILE
  ReportFile$ = "OPNENTRY.PRN"
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  ToPrint$ = SPACE$(80)

  GOSUB PrnOEHeader

  OpenEntry(1).TrDate = FY2BegDate
  OpenEntry(1).Desc = "Opening Entry"
  OpenEntry(1).Ref = "Sys"
  OpenEntry(1).Src = "GJ" + LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 2)
  OpenEntry(1).NextTran = 0
  
  FirstTime = True

  FOR Cnt = 1 TO NumGLAccts   'NumGLAccts
    GET AcctIdxFileNum, Cnt, AcctIndex(1)
    GET AcctFileNum, AcctIndex(1).RecNum, Accts(1)

    FundCode$ = LEFT$(Accts(1).Num, FundLen)
    ClosingThisFund = False
    FOR F = 1 TO NumFundstoClose
      IF FundCode$ = QPTrim$(FundCloseList$(F)) THEN
        ClosingThisFund = True
        EXIT FOR
      END IF
    NEXT
    
    IF ClosingThisFund THEN
      IF FirstTime THEN
        FirstTime = False
        LastFund$ = FundCode$
      END IF

      IF LastFund$ <> FundCode$ THEN
        GOSUB SubTotalFund
        GOSUB PrnOEHeader
      END IF

      IF Round#(Accts(1).Bal) <> 0 THEN
        SELECT CASE Accts(1).Typ
          CASE "A"
            IF Accts(1).Bal > 0 THEN
              OpenEntry(1).DrAmt = Round#(Accts(1).Bal)
              OpenEntry(1).CrAmt = 0
            ELSE
              OpenEntry(1).DrAmt = 0
              OpenEntry(1).CrAmt = Round#(ABS(Accts(1).Bal))
            END IF
            RecCnt = RecCnt + 1
            OpenEntry(1).AcctNum = Accts(1).Num
            GOSUB WriteToInterfaceFile
            GOSUB PrintEntry
            LastFund$ = LEFT$(Accts(1).Num, FundLen)
          CASE "L"
            IF Accts(1).Bal > 0 THEN
              OpenEntry(1).DrAmt = 0
              OpenEntry(1).CrAmt = Round#(Accts(1).Bal)
            ELSE
              OpenEntry(1).DrAmt = Round#(ABS(Accts(1).Bal))
              OpenEntry(1).CrAmt = 0
            END IF
            RecCnt = RecCnt + 1
            OpenEntry(1).AcctNum = Accts(1).Num
            GOSUB WriteToInterfaceFile
            GOSUB PrintEntry
            LastFund$ = LEFT$(Accts(1).Num, FundLen)
          CASE ELSE
        END SELECT
      END IF    'Account has balance
    END IF  'Fund being closed
  NEXT

  GOSUB SubTotalFund
  
  CLOSE
  
  PrintRptFile "G/L Opening Entries", ReportFile$, LPTNo, RetCode%, EntryPoint
  
EXIT SUB

PrnOEHeader:
  PRINT #PrnFile, User$
  PRINT #PrnFile, "Opening Entries"
  PRINT #PrnFile, "Fiscal Year Beginning "; Num2Date(FY2BegDate)
  PRINT #PrnFile,
  PRINT #PrnFile, "Acct                                                       Debit         Credit"
  PRINT #PrnFile, "-------------------------------------------------------------------------------"
RETURN

PrintEntry:
  LSET ToPrint$ = Accts(1).Num
  MID$(ToPrint$, 17) = Accts(1).Title
  MID$(ToPrint$, 53) = FUsing$(STR$(OpenEntry(1).DrAmt), Fmt$)
  MID$(ToPrint$, 67) = FUsing$(STR$(OpenEntry(1).CrAmt), Fmt$)
  PRINT #PrnFile, ToPrint$
  LSET ToPrint$ = ""
  TotDr# = Round#(TotDr# + OpenEntry(1).DrAmt)
  TotCr# = Round#(TotCr# + OpenEntry(1).CrAmt)
RETURN

SubTotalFund:

  LSET ToPrint$ = ""
  PRINT #PrnFile, ToPrint$
  LSET ToPrint$ = "Fund Totals"
  MID$(ToPrint$, 53) = FUsing$(STR$(TotDr#), Fmt$)
  MID$(ToPrint$, 67) = FUsing$(STR$(TotCr#), Fmt$)
  PRINT #PrnFile, ToPrint$

  LSET ToPrint$ = ""
  PRINT #PrnFile, ToPrint$

  OutofBal# = Round(TotDr#) - Round#(TotCr#)
  IF OutofBal# <> 0 THEN
    PRINT #PrnFile, ToPrint$
    LSET ToPrint$ = "Fund is Out of Balance :" + FUsing$(STR$(OutofBal#), Fmt$)
    PRINT #PrnFile, ToPrint$
  END IF

  PRINT #PrnFile, FF$

  LSET ToPrint$ = ""
  FirstTime = True
  TotDr# = 0
  TotCr# = 0

RETURN

WriteToInterfaceFile:
  PUT OEFileNum, RecCnt, OpenEntry(1)
RETURN


END SUB

FUNCTION Num2Month% (Dt%)

    D$ = Num2Date(Dt%)
    M$ = RIGHT$(D$, 2) + LEFT$(D$, 2)
    Num2Month% = VAL(M$)
    
END FUNCTION

SUB PostAdj (Unposted)

  ON LOCAL ERROR RESUME NEXT
  MKDIR "CLOSE97"
  ON LOCAL ERROR GOTO 0

  'Posts the fund balance adjustments to the accts file.
  'Must run prior to creating the opening entries for new year
  'backup files first
  
  LibFile2Scrn "GL.QSL", "OPS", MonoCode, Attribute, ErrorCode


  LOCATE 6, 10, 0
  QPrintRC "Copying Files...", 5, 10, 15
  SHELL "copy gl*.dat CLOSE97 > NUL"
  SHELL "copy gl*.idx CLOSE97  > NUL"
  SHELL "copy bgt*.dat CLOSE97 > NUL"
  
  LibFile2Scrn "GL.QSL", "OPS", MonoCode, Attribute, ErrorCode
  QPrintRC "Posting Adjustments...", 5, 10, 15
  PrintHelp "Wait"

  REDIM Accts(1) AS GLAcctRecType
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  
  FundBalAdjFileName$ = "FBADJ.DAT"
  DIM GLFBAdj AS GLFBAdjRecType
  FBARecLen = LEN(GLFBAdj)
  FundBalAdjFile = FREEFILE
  OPEN FundBalAdjFileName$ FOR RANDOM AS FundBalAdjFile LEN = FBARecLen
  NumFBARecs = LOF(FundBalAdjFile) \ FBARecLen

  FOR Cnt = 1 TO NumFBARecs
    GET FundBalAdjFile, Cnt, GLFBAdj
      AcctRec = FindAcct(QPTrim$(GLFBAdj.AcctNum))
      IF AcctRec > 0 THEN
        GET AcctFileNum, AcctRec, Accts(1)
        Accts(1).Bal = Accts(1).Bal + Round#(GLFBAdj.AdjAmt)
        PUT AcctFileNum, AcctRec, Accts(1)
      ELSE
        Unposted = Unposted + 1
      END IF
  NEXT

HandleError:

END SUB

SUB RepostNewYearTrans

      LibFile2Scrn "GL.QSL", "OPS", MonoCode, Attribute, ErrorCode
      QPrintRC "Posting Opening Entries...", 5, 10, 15
      Post2GL "OPNENTRY.DAT", BadTrans
      QPrintRC "Updating Transactions...  ", 5, 10, 15
      Post2GL "gltrans.y2", BadTrans
      
END SUB

SUB ResetYears

   DIM GLSetup AS GLSetupRecType
   SetUpRecLen = LEN(GLSetup)

   SetupFile = FREEFILE
   OPEN "GLSETUP.DAT" FOR RANDOM AS SetupFile LEN = SetUpRecLen

   GET SetupFile, 1, GLSetup

   FY1BegDate = GLSetup.FYBeg
   FY1EndDate = GLSetup.FYEnd
   FY2BegDate = GLSetup.NYBeg
   FY2EndDate = GLSetup.NYEnd

   FY1B$ = Num2Date(FY1BegDate)
   FY1BYr$ = RIGHT$(FY1B$, 4)
   NewYear = VAL(FY1BYr$) + 1
   NF1YB$ = LEFT$(FY1B$, 6) + QPTrim$(STR$(NewYear))
   
   FY1E$ = Num2Date(FY1EndDate)
   FY1EYr$ = RIGHT$(FY1E$, 4)
   NewYear = VAL(FY1EYr$) + 1
   NF1YE$ = LEFT$(FY1E$, 6) + QPTrim$(STR$(NewYear))

   FY2B$ = Num2Date(FY2BegDate)
   FY2BYr$ = RIGHT$(FY2B$, 4)
   NewYear = VAL(FY2BYr$) + 1
   NF2YB$ = LEFT$(FY2B$, 6) + QPTrim$(STR$(NewYear))
   
   FY2E$ = Num2Date(FY2EndDate)
   FY2EYr$ = RIGHT$(FY2E$, 4)
   NewYear = VAL(FY2EYr$) + 1
   NF2YE$ = LEFT$(FY2E$, 6) + QPTrim$(STR$(NewYear))
   
   GLSetup.FYBeg = Date2Num(NF1YB$)
   GLSetup.FYEnd = Date2Num(NF1YE$)
   GLSetup.NYBeg = Date2Num(NF2YB$)
   GLSetup.NYEnd = Date2Num(NF2YE$)

   PUT SetupFile, 1, GLSetup

   CLOSE SetupFile
   

END SUB

SUB SplitTransFile
  
  LibFile2Scrn "GL.QSL", "OPS", MonoCode, Attribute, ErrorCode
  QPrintRC "Updating Transaction Files.", 5, 10, 15
  PrintHelp "Parsing Histories..."

  SHARED TotLen%, FundLen%, AcctLen%, DetLen%

  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  Yr$ = RIGHT$(Num2Date(FY1EndDate), 2)
  
  '--List of Funds to close
  REDIM FundCloseList$(1)
  GetFundCloseList FundCloseList$(), NumFundstoClose
  
  DIM Tr AS GLTransRecType
  DIM Y1Trans AS GLTransRecType
  DIM Y2Trans AS GLTransRecType
  TransRecLen = LEN(Trans)
  
  OpenTransFile GLTransFile, NumTrans&

  Y1TransFile = FREEFILE
  OPEN "GLTRANS.Y1" FOR RANDOM AS Y1TransFile LEN = TransRecLen

  Y2TransFile = FREEFILE
  OPEN "GLTRANS.Y2" FOR RANDOM AS Y2TransFile LEN = TransRecLen
  
  FOR TrRec& = 1 TO NumTrans&
    GET GLTransFile, TrRec&, Tr

    Complete# = (TrRec& / NumTrans&) * 100
    LOCATE 6, 10
    COLOR 15
    PRINT USING "Processing files.  ###% Complete.        "; Complete#

    FundCode$ = LEFT$(Tr.AcctNum, FundLen)
    ClosingThisFund = False
    FOR F = 1 TO NumFundstoClose
      IF FundCode$ = QPTrim$(FundCloseList$(F)) THEN
        ClosingThisFund = True
        EXIT FOR
      END IF
    NEXT
    
    IF Tr.TrDate <= FY1EndDate AND ClosingThisFund THEN
      '--copy Year 1 trans
      Y1Recs& = Y1Recs& + 1
      Y1Trans.AcctNum = Tr.AcctNum
      Y1Trans.TrDate = Tr.TrDate
      Y1Trans.Desc = Tr.Desc
      Y1Trans.CrAmt = Tr.CrAmt
      Y1Trans.DrAmt = Tr.DrAmt
      Y1Trans.Ref = Tr.Ref
      Y1Trans.Src = Tr.Src
      Y1Trans.NextTran = 0 'Tr.NextTran
      PUT Y1TransFile, Y1Recs&, Y1Trans
    ELSE
      '--copy Year 2 trans
      Y2Recs& = Y2Recs& + 1
      Y2Trans.AcctNum = Tr.AcctNum
      Y2Trans.TrDate = Tr.TrDate
      Y2Trans.Desc = Tr.Desc
      Y2Trans.CrAmt = Tr.CrAmt
      Y2Trans.DrAmt = Tr.DrAmt
      Y2Trans.Ref = Tr.Ref
      Y2Trans.Src = Tr.Src
      Y2Trans.NextTran = 0  'Will be reposted
      PUT Y2TransFile, Y2Recs&, Y2Trans
    END IF

  NEXT

  CLOSE

  '--save the original transaction file

  CALL KillFile("gltrans.d97")
  NAME "gltrans.dat" AS "gltrans.d97"

  '--rename the gltrans.y1 file to .dat the relink y1
  NAME "gltrans.y1" AS "gltrans.dat"
  
  'Year 1 data now has .dat file extensions. Call relink to tie
  'everything together so we can run reports accurately in the future.
  ReLinkTrans

  QPrintRC SPACE$(80), 1, 1, 112
  LibFile2Scrn "GL.QSL", "OPS", MonoCode, Attribute, ErrorCode
  
  HistDir$ = "GL" + Yr$
  ON LOCAL ERROR RESUME NEXT
  MKDIR HistDir$
  ON LOCAL ERROR GOTO 0
  LOCATE 6, 10, 0: COLOR 15

  '--after linking copy the *.dat to the archive directory
  SHELL "copy gl*.dat " + HistDir$ + " > NUL"
  SHELL "copy gl*.idx " + HistDir$ + " > NUL"
  SHELL "copy bgt*.dat " + HistDir$ + " > NUL"
  SHELL "copy glrpt.exe " + HistDir$ + " > NUL"
  SHELL "copy gl*.q?l " + HistDir$ + " > NUL"

  KillFile "BGTTRANS.DAT"
  KillFile "GLTRANS.DAT"

  QPrintRC SPACE$(80), 1, 1, 112
  LibFile2Scrn "GL.QSL", "OPS", MonoCode, Attribute, ErrorCode
  
  'NewBgtFileName$ = "BGTTRANS.D" + Yr$
  'IF Exist("BGTTRANS.DAT") THEN
  '  NAME "BGTTRANS.DAT" AS NewBgtFileName$
  'END IF

  '--rename y2 to current & new year transactions
  'NAME "gltrans.y2" AS "gltrans.dat"
  'ReLinkTrans
  
  '--opening balances and y2 transactions will be posted
  'this is a copy of y1 trans file so kill it.  New trans will
  'be created when opening entries and gltrans.y2 files are
  'posted
  
  'IF Exist("BGTTRANS.DAT") THEN
  '  KILL "BGTTRANS.DAT"
  'END IF

END SUB

SUB UpDateNewYear

'*******************************************************************
  'operations to update the system for the new year.
  'last operation to be called.  Year one files are archived.

  'before this operation glacct is a copy of the file which is
  'linked to last years transactions.
'*******************************************************************

  '--List of Funds to close
  REDIM FundCloseList$(1)
  GetFundCloseList FundCloseList$(), NumFundstoClose


  QPrintRC "Updating files for new year", 5, 10, 15
  GetFYDates FY1BegDate, FY1EndDate, FY2BegDate, FY2EndDate
  
  REDIM Accts(1) AS GLAcctRecType
  OpenAcctFile AcctFileNum, NumGLAcctRecs

  REDIM BgtTrn(1) AS GLTransRecType
  
  BgtTransFile = FREEFILE
  BgtTransRecLen = LEN(BgtTrn(1))
  OPEN "BgtTrans.DAT" FOR RANDOM AS BgtTransFile LEN = BgtTransRecLen
  
  'Set the vars for the budget transaction history
  BgtTrn(1).TrDate = FY2BegDate
  BgtTrn(1).Desc = "Appropriation"
  BgtTrn(1).Ref = "Sys"
  BgtTrn(1).Src = "GJ" + LEFT$(DATE$, 2) + MID$(DATE$, 4, 2) + RIGHT$(DATE$, 2)
  BgtTrn(1).NextTran = 0
  BgtTrn(1).Marked = 0
  
  FOR R = 1 TO NumGLAcctRecs
    GET AcctFileNum, R, Accts(1)
    'IF LEFT$(Accts(1).Num, 9) = "70-151-01" THEN STOP
    FundCode$ = LEFT$(Accts(1).Num, FundLen)
    ClosingThisFund = False
    FOR F = 1 TO NumFundstoClose
      IF FundCode$ = QPTrim$(FundCloseList$(F)) THEN
        ClosingThisFund = True
        EXIT FOR
      END IF
    NEXT

   IF ClosingThisFund THEN
    M$ = "Updating Account: " + Accts(1).Num
    QPrintRC M$, 6, 10, -1

    '--create budget history file using the acct n/y approved budget field
    IF Accts(1).Typ = "R" OR Accts(1).Typ = "E" THEN
     IF NOT Accts(1).Deleted THEN
      IF Round#(Accts(1).NYApp) > -99999999.99# AND Round#(Accts(1).NYApp) < 99999999.99# THEN
       IF Round#(Accts(1).NYApp) <> 0 THEN
        BRec = BRec + 1
        BgtTrn(1).AcctNum = Accts(1).Num
        SELECT CASE Accts(1).Typ
          CASE "R"
            IF Accts(1).NYApp >= 0 THEN
              BgtTrn(1).CrAmt = Accts(1).NYApp
              BgtTrn(1).DrAmt = 0
            ELSE
              BgtTrn(1).CrAmt = 0
              BgtTrn(1).DrAmt = ABS(Accts(1).NYApp)
            END IF
          CASE "E"
            IF Accts(1).NYApp >= 0 THEN
              BgtTrn(1).DrAmt = Accts(1).NYApp
              BgtTrn(1).CrAmt = 0
            ELSE
              BgtTrn(1).DrAmt = 0
              BgtTrn(1).CrAmt = ABS(Accts(1).NYApp)
           END IF
          CASE ELSE
        END SELECT
        PUT BgtTransFile, BRec, BgtTrn(1)
       END IF
      END IF
     END IF
    END IF
   
    Accts(1).FrstTran = 0
    Accts(1).LastTran = 0
    Accts(1).PYAct = Accts(1).Work
    Accts(1).BegBal = 0
    Accts(1).Bgt = 0
    Accts(1).Bal = 0
    Accts(1).Encumb = 0
    Accts(1).MTD = 0
    Accts(1).YTD = 0
    Accts(1).NYEst = 0
    Accts(1).NYReq = 0
    Accts(1).NYRec = 0
    Accts(1).NYApp = 0
    Accts(1).FrstBTran = 0
    Accts(1).LastBTran = 0
    Accts(1).FrstPTran = 0
    Accts(1).LastPTran = 0
    'Accts(1).Work  'copy to p/y then 0
    Accts(1).Marked = 0
    PUT AcctFileNum, R, Accts(1)

    'reset work field
    Accts(1).Work = 0
    PUT AcctFileNum, R, Accts(1)
   END IF
  NEXT

  CLOSE

'************************
'acct field names
'Deleted
'Num
'Title
'Typ
'FrstTran
'LastTran
'PYAct
'BegBal
'Bgt
'Bal
'Encumb
'MTD
'YTD
'NYEst
'NYReq
'NYRec
'NYApp
'FrstBTran
'LastBTran
'FrstPTran
'LastPTran
'Work
'Res
'Marked
  
END SUB

