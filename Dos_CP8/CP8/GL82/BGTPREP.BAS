DECLARE SUB ClrBG ()
'BGTPREP.BAS
DEFINT A-Z

'$INCLUDE: 'DefCnf.bi'                    'monitor and color info
'$INCLUDE: 'FieldInf.bi'                  'field information
'$INCLUDE: 'FormEdit.bi'                  'form editing information
'$INCLUDE: 'PRO7.BI'                      'QuickPak Pro Routines
'$INCLUDE: 'QSCR.BI'                      'QuickScreen Routines
'$INCLUDE: 'GLAUX.BI'                     'Common GL Routines
'$INCLUDE: 'GL.bi'                        'GL File Types
'$INCLUDE: 'SCRAUX.BI'

'Subs in this module
DECLARE FUNCTION GetDeptTitle$ (DeptRecNum%)
DECLARE SUB PrintRptFile (RptTitle$, FileName$, LPTPort%, RetCode%, EntryPoint%)
DECLARE SUB BgtPrepEditAcct ()
DECLARE SUB MoveBgtAmts (Operation%)
DECLARE SUB PrintBgtWorksht ()
DECLARE SUB ClearWorksheet ()
DECLARE SUB PrintWorksheet ()
'DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB MnuTitle (Title$, Row, Col, MaxLen, clr)

DIM SHARED TotLen%, FundLen%, AcctLen%, DetLen%
DIM SHARED Acct AS GLAcctRecType
DIM SHARED AcctIdx AS GLAcctIndexType
DIM SHARED FundIdx AS GLFundIndexType
DIM SHARED Trans AS GLTransRecType
DIM SHARED GLSetUp1 AS GLSumSetupRecType

'$INCLUDE: 'SetCnf.bi'                    'Set monitor and color information
'$INCLUDE: 'init.bi'                      'Get GL system info
''$INCLUDE: 'keycode.bi'                   'Define keycodes

REDIM SHARED Choice$(0, 0)                'Used to hold multiple choices

REDIM MChoice$(1 TO 8)

MChoice$(1) = "Enter/Edit Budget Prep Amounts"
MChoice$(2) = "Copy Current Bgt to Estimate"
MChoice$(3) = "Copy Estimated to Requested"
MChoice$(4) = "Copy Requested to Recommended"
MChoice$(5) = "Copy Recommended to Approved"
MChoice$(6) = "Print Budget Prep Worksheet"
MChoice$(7) = "Clear Worksheet"
MChoice$(8) = "Quit"

'MChoice$(4) = "Jim"
'MChoice$(5) = "Util"

MaxLen = 0                              'Set menu width to zero
BoxBot = 20                               'limit the box to line 20
Action = 0                                'Stay on menu til a choice is made
Choice = 1                                'Pre-load choice to highlight

CONST False = 0
CONST True = NOT False

FOR Cnt = 1 TO UBOUND(MChoice$)           'Find max menu width
   TLen = LEN(MChoice$(Cnt))
   IF TLen > MaxLen THEN
      MaxLen = TLen
   END IF
NEXT

Row = ((25 - (UBOUND(MChoice$))) \ 2)  'Center Menu within Screen
Col = ((80 - MaxLen) \ 2) - 2

TextCursor -2, -2                         'Define the Mouse cursor
ShowCursor                                'Show it

DO

   LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, Attribute, Errorcode
   'TitleBox 3, Col, MaxLen + 3, "Budget Prep Menu ", Cnf
   'TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

   Title$ = "Budget Prep Menu "
   MnuTitle Title$, Row, Col, MaxLen, 15
   PrintTitle User$

   'PrintHelp "Chart of Accounts"

   LOCATE Row, Col, 0
   VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf

   IF Ky$ = CHR$(27) THEN EXIT DO

   SELECT CASE Choice
      CASE 1
        ClrBG
        BgtPrepEditAcct
      CASE 2
        ClrBG
        MoveBgtAmts 1
      CASE 3
        ClrBG
        MoveBgtAmts 2
      CASE 4
        ClrBG
        MoveBgtAmts 3
      CASE 5
        ClrBG
        MoveBgtAmts 4
      CASE 6
        ClrBG
        PrintWorksheet
      CASE 7
        ClrBG
        ClearWorksheet
      CASE 8
        CLS : END
   END SELECT

LOOP

ClrBG
HideCursor                                'Be nice and turn off mouse cursor
DispLoad
RUN "BGTEDIT"

SUB BgtPrepEditAcct

   'PRINT FRE(-2)
   
   REDIM Form$(0, 0)                      'Form array holds data while editing
   REDIM Fld(1) AS FieldInfo              'Field info Array
   DIM Frm AS FormInfo                    'Form editing info
   LibName$ = "GL.QSL"                    'Screen library
   FormName$ = "BGTPREP"
   LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
   Action = 1

   DO
     EditForm Form$(), Fld(), Frm, Cnf, Action
      IF Frm.PrevFld <> Frm.FldNo THEN    'if we have left a field
         IF Frm.FldNo = 1 THEN            'Blank the form if back on acct fld
           GOSUB ClearForm
           GOSUB Clearlabels
           RecordNum = 0
         END IF
         SELECT CASE Frm.PrevFld          'user is moving around in form
            CASE 1
               IF Frm.FldEdited THEN      'If the field was edited
                  Form$(1, 0) = QPTRIM$(Form$(1, 0))   'trim the field &
                  IF LEN(Form$(1, 0)) <> TotLen THEN   'make sure len is ok
                     Form$(1, 0) = ""
                     GOSUB Clearlabels
                     RecordNum = 0
                     Frm.FldNo = 1
                  ELSE                    'format & look up acct
                     Form$(1, 0) = FmtAcct$(Form$(1, 0), FundLen%, AcctLen%, DetLen%)
                     RecordNum = FindAcct(Form$(1, 0))
                     IF RecordNum > 0 THEN
                       GOSUB GetAcct
                       Frm.FldNo = 2
                     ELSE
                       Form$(1, 0) = ""
                       GOSUB Clearlabels
                       RecordNum = 0
                       Frm.FldNo = 1
                     END IF
                  END IF
               END IF
            CASE ELSE
         END SELECT
         Action = 1
      END IF
      SELECT CASE Frm.KeyCode             'Check the user's key presses
         CASE F10Key                      'SaveButton
            IF RecordNum > 0 THEN
              GOSUB SaveRec
              GOSUB ClearForm
              GOSUB Clearlabels
            ELSE
              BEEP
            END IF
         CASE EscKey
            EXIT DO
            'Check to see if form was edited before ESC was pressed
      END SELECT
   LOOP UNTIL Frm.KeyCode = EscKey

EXIT SUB

GetAcct:
   OpenAcctFile AcctFileNum, NumAccts
      GOSUB Clearlabels
      GET AcctFileNum, RecordNum, Acct

      QPrintRC LEFT$(Acct.Title, 25), 8, 36, -1
      QPrintRC Acct.Typ, 9, 36, -1
      LOCATE 10, 42: COLOR 15: PRINT USING "########,.##"; Acct.Bgt
      Form$(2, 0) = STR$(Acct.PYAct)
      Form$(3, 0) = STR$(Acct.NYEst)
      Form$(4, 0) = STR$(Acct.NYReq)
      Form$(5, 0) = STR$(Acct.NYRec)
      Form$(6, 0) = STR$(Acct.NYApp)

   CLOSE AcctFileNum
   
RETURN


SaveRec:
   OpenAcctFile AcctFileNum, NumAccts
     Acct.PYAct = Value(Form$(2, 0), errcode)
     Acct.NYEst = Value(Form$(3, 0), errcode)
     Acct.NYReq = Value(Form$(4, 0), errcode)
     Acct.NYRec = Value(Form$(5, 0), errcode)
     Acct.NYApp = Value(Form$(6, 0), errcode)
     PUT AcctFileNum, RecordNum, Acct
   CLOSE

RETURN


ClearForm:                               'reset form for more editing
   LOCATE 1, 1, 0
   FOR Cnt = 1 TO 8
      Form$(Cnt, 0) = ""
   NEXT
   Frm.FldNo = 1
   Action = 1

RETURN

Clearlabels:
  FOR Cnt = 8 TO 10
    QPrintRC SPACE$(25), Cnt, 36, -1
  NEXT
RETURN


END SUB

SUB ClearWorksheet

  REDIM Form$(0, 0)                      'Form array holds data while editing
  REDIM Fld(1) AS FieldInfo              'Field info Array
  DIM Frm AS FormInfo                    'Form editing info
  LibName$ = "GL.QSL"                    'Screen library
  FormName$ = "BGTCLR"
  BGScr$ = "BG"
  LibFile2Scrn LibName$, BGScr$, MonoCode, Attribute, Errorcode
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()

  CancelFlag = 0
  Action = 1

  DO
     EditForm Form$(), Fld(), Frm, Cnf, Action

      SELECT CASE Frm.KeyCode             'Check the user's key presses
         CASE F10Key                      'SaveButton
            EXIT DO
         CASE EscKey
            CancelFlag = -1
            EXIT DO
      END SELECT
  LOOP

  IF CancelFlag THEN EXIT SUB
  
  REDIM Acct(1) AS GLAcctRecType
  OpenAcctFile AcctFileNum, NumAccts

  FOR Cnt = 1 TO NumAccts
    PctDone$ = STR$(INT((Cnt / NumAccts) * 100)) + "%"
    QPrintRC "Processing: " + PctDone$, 25, 2, -1

    GET AcctFileNum, Cnt, Acct(1)
    Acct(1).NYEst = 0
    Acct(1).NYReq = 0
    Acct(1).NYRec = 0
    Acct(1).NYApp = 0
    PUT AcctFileNum, Cnt, Acct(1)

  NEXT

  CLOSE AcctFileNum


END SUB

SUB MoveBgtAmts (Operation)

  'CLS
  'PRINT "FRE(-2)"; FRE(-2)
  'PRINT "Press any key."
  'K$ = INPUT$(1)
  'CLS

  REDIM Acct(1) AS GLAcctRecType
  LibName$ = "GL.QSL"                    'Screen library
  FormName$ = "BGTMOVE"
  ScrnName$ = "BG"
  LibFile2Scrn LibName$, ScrnName$, MonoCode, Attribute, Errorcode

  ok = MsgBox%("GL", "BGTMOVEV")
  IF ok = 1 THEN EXIT SUB
  
  SELECT CASE Operation
    CASE 1
      '--copy Bgt to Requested
      WT$ = "Copy C/Y Budget to Estimate "
    CASE 2
      '--copy Est to Requested
      WT$ = "Copy Estimated to Requested "
    CASE 3
      '--copy Requested to Recommeded
      WT$ = "Copy Requested to Recommended "
    CASE 4
      '--copy Recommeded to Approved
      WT$ = "Copy Recommended to Approved "
  END SELECT

  '--WT$ is title for window
  TRow = 8
  TCol = 23
  WTLen = LEN(WT$) '--Find its length
  BrkOffSet = TCol + WTLen '--calc where to print ending bracket

  REDIM Form$(0, 0)                      'Form array holds data while editing
  REDIM Fld(1) AS FieldInfo              'Field info Array

  SHARED Choice$()
  REDIM Choice$(0 TO 2, 1)               '--Pick List Choices
  Choice$(0, 0) = "2"
  Choice$(1, 0) = "Increase"
  Choice$(2, 0) = "Decrease"
  DIM Frm AS FormInfo                    'Form editing info
  
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  QPrintRC WT$, TRow, TCol, 14
  QPrintRC "]", TRow, BrkOffSet, 10
  
  Form$(1, 0) = "0.00"
  Form$(2, 0) = Choice$(1, 0)
  CancelFlag = 0
  Action = 1

  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
    SELECT CASE Frm.KeyCode
      CASE F10Key
        SELECT CASE QPTRIM$(Form$(2, 0))
          CASE "Increase"
            FactorIn$ = Form$(1, 0)
          CASE "Decrease"
            FactorIn$ = "-" + Form$(1, 0)
        END SELECT
        Factor# = Round(VAL(FactorIn$) * .01)
        EXIT DO
      CASE EscKey
        CancelFlag = -1
        EXIT DO
    END SELECT
  LOOP
   
  IF CancelFlag THEN EXIT SUB
  
  OpenAcctFile AcctFileNum, NumAccts
  
  FOR Cnt = 1 TO NumAccts
    PctDone$ = STR$(INT((Cnt / NumAccts) * 100)) + "%"
    QPrintRC "Processing: " + PctDone$, 25, 2, -1
    
    GET AcctFileNum, Cnt, Acct(1)

    'IF QPTrim$(Acct(1).Num) = "10-420-33" THEN STOP

    '--carthage fix
    'FundNum$ = LEFT$(Acct(1).Num, 2)
    'IF FundNum$ = "10" AND Acct(1).Typ = "R" THEN

    SELECT CASE Operation
      CASE 1
        '--copy Bgt to Est
        IF Factor# <> 0 THEN
          NewVal# = Round#(Acct(1).Bgt * Factor#) + Acct(1).Bgt
          Acct(1).NYEst = NewVal#
        ELSE
          Acct(1).NYEst = Acct(1).Bgt
        END IF
      CASE 2
        '--copy Est to Requested
        IF Factor# <> 0 THEN
          NewVal# = Round#(Acct(1).NYEst * Factor#) + Acct(1).NYEst
          Acct(1).NYReq = NewVal#
        ELSE
          Acct(1).NYReq = Acct(1).NYEst
        END IF
      CASE 3
        '--copy Requested to Recommeded
        IF Factor# <> 0 THEN
          NewVal# = Round#(Acct(1).NYReq * Factor#) + Acct(1).NYReq
          Acct(1).NYRec = NewVal#
        ELSE
          Acct(1).NYRec = Acct(1).NYReq
        END IF
      CASE 4
        '--copy Recommeded to Approved
        IF Factor# <> 0 THEN
          NewVal# = Round#(Acct(1).NYRec * Factor#) + Acct(1).NYRec
          Acct(1).NYApp = NewVal#
        ELSE
          Acct(1).NYApp = Acct(1).NYRec
        END IF
    END SELECT
    
    PUT AcctFileNum, Cnt, Acct(1)
    'END IF 'carthage fix
  NEXT

  CLOSE AcctFileNum

END SUB

SUB PrintBgtWorksht

  LOCATE , , 0
  SHARED FundLen
  SHARED Acct AS GLAcctRecType

  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  REDIM Choice$(0 TO 3, 1)
  DIM Frm AS FormInfo
  LibName$ = "GL.QSL"
  FormName$ = "BGTPRNOP"

  'Pick List Choices
  Choice$(0, 0) = "4"
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"

  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  'PrintHelp "Budget Prep Worksheet"
  Action = 1

  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action
      SELECT CASE Frm.KeyCode   'Check the user's key presses
        CASE F10Key            'SaveButton
          EndDate = Date2Num(Form$(1, 0))
          StartFund$ = Form$(2, 0)
          EndFund$ = Form$(3, 0)
          IF LEN(LTRIM$(RTRIM$(Form$(4, 0)))) = 0 THEN
             Dev$ = "S"
             LPTNo = 1
          ELSE
             Dev$ = LEFT$(Form$(4, 0), 1)
             LPTNo = VAL(RIGHT$(RTRIM$(Form$(4, 0)), 1))
          END IF
        CASE EscKey
          EXIT SUB
      END SELECT
  LOOP UNTIL Frm.KeyCode = F10Key
   
  LOCATE , , 0
  LibFile2Scrn "GL.QSL", "BG", MonoCode, Attribute, Errorcode
  PrintHelp "Processing report. Please wait."

  'Define strings used for printing
  Header$ = "Budget Prep Report"
  REDIM Desc$(1)
  Desc$(1) = " Acct Number     Title"
  DescColor = 3
  RptColor = 15
  CRLF$ = CHR$(13) + CHR$(10)
  FF$ = CHR$(12)
  ScrLine$ = STRING$(80, CHR$(205))
  PDivLine$ = STRING$(77, "-")
  'PrnLine$ = STRING$(77, "-")
  'Prn2Line$ = STRING$(77, "=")

  OpenAcctIdx AcctIdxFileNum, NumIdxRecs
  OpenAcctFile AcctFileNum, NumAccts
  PrnFile = FREEFILE
  ReportFile$ = "ACCTLIST.PRN"
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  
  'Now Print
  FOR Cnt = 1 TO NumIdxRecs
    GET AcctIdxFileNum, Cnt, AcctIdx
    FundCode$ = QPTRIM$(LEFT$(AcctIdx.AcctNum, FundLen))
    IF FundCode$ >= StartFund$ AND FundCode$ <= EndFund$ THEN
      GET AcctFileNum, AcctIdx.RecNum, Acct
      IF Acct.Typ = "R" OR Acct.Typ = "E" THEN
         HowMany = HowMany + 1
         ToPrint$ = SPACE$(80)
         MID$(ToPrint$, 2) = Acct.Num
         MID$(ToPrint$, 18) = Acct.Title
         MID$(ToPrint$, 63) = Acct.Typ
         'MID$(ToPrint$, 65) = FUsing(STR$(Acct.FrstTran), "#####")
         'MID$(ToPrint$, 72) = FUsing(STR$(Acct.LastTran), "#####")
         PRINT #PrnFile, ToPrint$
      END IF
    END IF
  NEXT

  PRINT #PrnFile,
  PRINT #PrnFile, HowMany; "Accounts listed."

  CLOSE
  HideCursor

  SELECT CASE Dev$
    CASE "S"
      PrintRptFile Header$, ReportFile$, LPTNo, RetCode, 2
    CASE "P"
      PrintRptFile Header$, ReportFile$, LPTNo, RetCode, 4
  END SELECT

  'KILL ReportFile$
  ShowCursor


END SUB

SUB PrintWorksheet

  LOCATE , , 0
  SHARED FundLen
  SHARED Acct AS GLAcctRecType

  REDIM Beg(100), BEnd(100), Desc$(100)
  REDIM Form$(0, 0)
  REDIM Fld(0) AS FieldInfo
  REDIM Choice$(0 TO 3, 1)
  DIM Frm AS FormInfo
  LibName$ = "GL.QSL"
  FormName$ = "BGTPRNOP"

  GetFundCodes FirstFund$, LastFund$

  'Pick List Choices
  Choice$(0, 0) = "6"
  Choice$(1, 0) = "Screen"
  Choice$(2, 0) = "Prn#1"
  Choice$(3, 0) = "Prn#2"
  
  LibLoadDisplayForm LibName$, FormName$, Form$(), Fld()
  'PrintHelp "Budget Prep Worksheet"

  Form$(1, 0) = DATE$
  Form$(2, 0) = FirstFund$
  Form$(3, 0) = LastFund$
  Form$(4, 0) = "R"
  Form$(5, 0) = "Y"
  
  Action = 1

  DO
    EditForm Form$(), Fld(), Frm, Cnf, Action

      SELECT CASE Frm.KeyCode
        '--SaveButton
        CASE F10Key
          EndDate = Date2Num(Form$(1, 0))
          StartFund$ = QPTRIM$(Form$(2, 0))
          EndFund$ = QPTRIM$(Form$(3, 0))

          '--print worksheet or completed report
          SELECT CASE Form$(4, 0)
            CASE "W"
              WorkSheet = True
            CASE "R", ""
              WorkSheet = False
          END SELECT
          '--print depts on separate pages
          SELECT CASE Form$(5, 0)
            CASE "Y"
               DeptOnNewPage = True
            CASE "N", ""
              DeptOnNewPage = False
          END SELECT
          '--Screen or printer
          IF LEN(LTRIM$(RTRIM$(Form$(6, 0)))) = 0 THEN
             Dev$ = "S"
             LPTNo = 1
          ELSE
             Dev$ = LEFT$(Form$(6, 0), 1)
             LPTNo = VAL(RIGHT$(RTRIM$(Form$(6, 0)), 1))
          END IF
        CASE EscKey
          EXIT SUB
      END SELECT

  LOOP UNTIL Frm.KeyCode = F10Key
  
  ReportFile$ = "BGTPREP.PRN"    'Report File Name
  CommaFmt$ = "#########,.##"    'format takes 13 chars
  BgtFmt$ = "########,"         'format takes 11 chars
  TotalFmt$ = "#########,.##"   'format takes 14 chars
  SumLine$ = STRING$(12, "-")   'column summary line
  BSumLine$ = STRING$(10, "-")  'summary line for budget columns
  PSumLine$ = "----"            'summary line for Pct columns
  DivLine$ = STRING$(79, "-")   'dashed line
  DivLine2$ = STRING$(79, "=")  'Double Line
  Underline$ = "  ___________"
  CRLF$ = CHR$(13) + CHR$(10)
  FF$ = CHR$(12)
  MaxLines = 55

  '--Column offsets for printing amounts
  Col1 = 38
  Col2 = 49
  Col3 = 63
  Col4 = 76
  Col5 = 89
  Col6 = 102
  Col7 = 115

  
  PrnFile = FREEFILE
  OPEN ReportFile$ FOR OUTPUT AS #PrnFile
  
  OpenAcctIdx AcctIdxFileNum, NumGLAccts
  OpenAcctFile AcctFileNum, NumGLAcctRecs
  OpenTransFile TransFileNum, NumTrans&
  OpenFundIdx FundIdxFile, NumFunds

  GLFile = FREEFILE
  GlSetUpLen = LEN(GLSetUp1)
  OPEN "GLSETSUM.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GLFile LEN = GlSetUpLen
  IF LOF(GLFile) > 0 THEN
   GET GLFile, 1, GLSetUp1
   Beg(1) = VAL(GLSetUp1.Beg1)
   BEnd(1) = VAL(GLSetUp1.End1)
   Desc$(1) = GLSetUp1.Desc1
   Beg(2) = VAL(GLSetUp1.Beg2)
   BEnd(2) = VAL(GLSetUp1.End2)
   Desc$(2) = GLSetUp1.Desc2
   Beg(3) = VAL(GLSetUp1.Beg3)
   BEnd(3) = VAL(GLSetUp1.End3)
   Desc$(3) = GLSetUp1.Desc3
   Beg(4) = VAL(GLSetUp1.Beg4)
   BEnd(4) = VAL(GLSetUp1.End4)
   Desc$(4) = GLSetUp1.Desc4
   Beg(5) = VAL(GLSetUp1.Beg5)
   BEnd(5) = VAL(GLSetUp1.End5)
   Desc$(5) = GLSetUp1.Desc5
   Beg(6) = VAL(GLSetUp1.Beg6)
   BEnd(6) = VAL(GLSetUp1.End6)
   Desc$(6) = GLSetUp1.Desc6
   Beg(7) = VAL(GLSetUp1.Beg7)
   BEnd(7) = VAL(GLSetUp1.End7)
   Desc$(7) = GLSetUp1.Desc7
   Beg(8) = VAL(GLSetUp1.Beg8)
   BEnd(8) = VAL(GLSetUp1.End8)
   Desc$(8) = GLSetUp1.Desc8
   Beg(9) = VAL(GLSetUp1.Beg9)
   BEnd(9) = VAL(GLSetUp1.End9)
   Desc$(9) = GLSetUp1.Desc9
   BreakFlag = True
   ELSE
   BreakFlag = False
   CLOSE GLFile
  END IF

  
  
  REDIM RevAccts%(1 TO NumGLAccts)      'Holds all rev acct record nums
  REDIM ExpAccts%(1 TO NumGLAccts)      'Holds all exp acct record nums
  REDIM FundList$(1 TO NumFunds)        'List of all active Funds

  '--Initialize FundList
  IF NumFunds = 0 THEN
    ok = MsgBox("GL", "NOFUNDS")
    CLOSE
    EXIT SUB
  END IF
  FOR Fund = 1 TO NumFunds
    GET FundIdxFile, Fund, FundIdx
    FundList$(Fund) = QPTRIM$(FundIdx.FundNum)
  NEXT
  CLOSE FundIdxFile

  PrintHelp "Processing:"

  '--Calculate balances thru ending date
  FOR GLAcct = 1 TO NumGLAccts

    '--Initialize
    BGTBal# = 0
    YTDBal# = 0
    
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
        YTDBal# = Round#(Acct.BegBal)     'get the beginning balance
        NextTr& = Acct.FrstTran           'get the first trans for this acct

        DO UNTIL NextTr& = 0              'keep going 'til we run out

          GET TransFileNum, NextTr&, Trans
          
          '--Get balance thru end date
          IF Trans.TrDate <= EndDate THEN
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
        Acct.YTD = Round#(YTDBal#)
        PUT AcctFileNum, AcctIdx.RecNum, Acct

      END IF    '--test for rev or exp accts
    END IF      '--End of acct in fund range test
  NEXT          'Process next account

  '--Now write the report to file.
  PrintHelp "Generating Report..."
  FirstTimeThru = -1

  '--Check for fund in range we want
  FOR Fund = 1 TO NumFunds
    ThisFund$ = FundList$(Fund)
    IF ThisFund$ >= StartFund$ AND ThisFund$ <= EndFund$ THEN
      UsingFund = True
      FundRec = FindFund(ThisFund$)  'Get the fund name
      FundName$ = QPTRIM$(GetFundTitle(FundRec))
      Dept$ = ""
      LastDept$ = ""
      LastDeptName$ = ""
    ELSE
      UsingFund = False
    END IF
    
    IF UsingFund THEN
      Category$ = " Revenues"
      DoingExp = 0
      '--Advance page for each new fund except first fund
      IF FirstTimeThru THEN
        FirstTimeThru = 0
      ELSE
        PRINT #PrnFile, FF$
      END IF
      GOSUB PrintBGTWksPageHeader

      '--Search thru list of revenue account record numbers
      '--to see if we find one in this fund.
      FOR Cnt = 1 TO RCnt
        Rec = RevAccts%(Cnt)
        GET AcctFileNum, Rec, Acct
        FundCode$ = LEFT$(Acct.Num, FundLen)
        
        '--Yup.. found one
        IF FundCode$ = ThisFund$ THEN
          
          '--Print Account detail
          ToPrint$ = SPACE$(132)
          Account$ = QPTRIM$(Acct.Num) + "  " + QPTRIM$(Acct.Title)
          LSET ToPrint$ = LEFT$(Account$, 36)
          MID$(ToPrint$, Col1) = FUsing$(STR$(Acct.Bgt), BgtFmt$)
          MID$(ToPrint$, Col2) = FUsing$(STR$(Acct.YTD), CommaFmt$)
          MID$(ToPrint$, Col3) = FUsing$(STR$(Acct.PYAct), CommaFmt$)
          IF WorkSheet = False THEN
            MID$(ToPrint$, Col4) = FUsing$(STR$(Acct.NYEst), CommaFmt$)
            MID$(ToPrint$, Col5) = FUsing$(STR$(Acct.NYReq), CommaFmt$)
            MID$(ToPrint$, Col6) = FUsing$(STR$(Acct.NYRec), CommaFmt$)
            MID$(ToPrint$, Col7) = FUsing$(STR$(Acct.NYApp), CommaFmt$)
          ELSE
            IF Acct.NYEst <> 0 THEN
              MID$(ToPrint$, Col4) = FUsing$(STR$(Acct.NYEst), CommaFmt$)
            ELSE
              MID$(ToPrint$, Col4) = Underline$
            END IF
            IF Acct.NYReq <> 0 THEN
              MID$(ToPrint$, Col5) = FUsing$(STR$(Acct.NYReq), CommaFmt$)
            ELSE
              MID$(ToPrint$, Col5) = Underline$
            END IF
            IF Acct.NYRec <> 0 THEN
              MID$(ToPrint$, Col6) = FUsing$(STR$(Acct.NYRec), CommaFmt$)
            ELSE
              MID$(ToPrint$, Col6) = Underline$
            END IF
            IF Acct.NYApp <> 0 THEN
              MID$(ToPrint$, Col7) = FUsing$(STR$(Acct.NYApp), CommaFmt$)
            ELSE
              MID$(ToPrint$, Col7) = Underline$
            END IF
          END IF
          PRINT #PrnFile, ToPrint$
          LineCnt = LineCnt + 1
          IF LineCnt > MaxLines THEN
            PRINT #PrnFile, FF$
            GOSUB PrintBGTWksPageHeader
          END IF
          
          '--Add account to running totals
          BgtSum# = BgtSum# + Acct.Bgt
          YTDSum# = YTDSum# + Acct.YTD
          PYActSum# = PYActSum# + Acct.PYAct
          NYEstSum# = NYEstSum# + Acct.NYEst
          NYReqSum# = NYReqSum# + Acct.NYReq
          NYRecSum# = NYRecSum# + Acct.NYRec
          NYAppSum# = NYAppSum# + Acct.NYApp

        END IF

      NEXT Cnt     'Revenue Acct

      '--Summarize Revenues
      GOSUB PrintSummaryLines
      
      ToPrint$ = SPACE$(132)
      LSET ToPrint$ = "  Total Revenues"
      MID$(ToPrint$, Col1) = FUsing$(STR$(BgtSum#), BgtFmt$)
      MID$(ToPrint$, Col2) = FUsing$(STR$(YTDSum#), TotalFmt$)
      MID$(ToPrint$, Col3) = FUsing$(STR$(PYActSum#), CommaFmt$)
      IF WorkSheet = False THEN
        MID$(ToPrint$, Col4) = FUsing$(STR$(NYEstSum#), CommaFmt$)
        MID$(ToPrint$, Col5) = FUsing$(STR$(NYReqSum#), CommaFmt$)
        MID$(ToPrint$, Col6) = FUsing$(STR$(NYRecSum#), CommaFmt$)
        MID$(ToPrint$, Col7) = FUsing$(STR$(NYAppSum#), CommaFmt$)

      ELSE
        IF NYEstSum# <> 0 THEN
          MID$(ToPrint$, Col4) = FUsing$(STR$(NYEstSum#), CommaFmt$)
        ELSE
          MID$(ToPrint$, Col4) = Underline$
        END IF
        IF NYReqSum# <> 0 THEN
          MID$(ToPrint$, Col5) = FUsing$(STR$(NYReqSum#), CommaFmt$)
        ELSE
          MID$(ToPrint$, Col5) = Underline$
        END IF
        IF NYRecSum# <> 0 THEN
          MID$(ToPrint$, Col6) = FUsing$(STR$(NYRecSum#), CommaFmt$)
        ELSE
          MID$(ToPrint$, Col6) = Underline$
        END IF
        IF NYAppSum# <> 0 THEN
          MID$(ToPrint$, Col7) = FUsing$(STR$(NYAppSum#), CommaFmt$)
        ELSE
          MID$(ToPrint$, Col7) = Underline$
        END IF
      END IF

      PRINT #PrnFile, ToPrint$
      LineCnt = LineCnt + 1
      IF LineCnt > MaxLines THEN
        PRINT #PrnFile, FF$
        GOSUB PrintBGTWksPageHeader
      END IF
      
      '--Assign Summary totals for revenues to New variables so that we
      '--can reuse summary vars for exp
      FundRevBgt# = BgtSum#
      FundRevYTD# = YTDSum#
      FundRevPYActSum# = PYActSum#
      FundRevNYEstSum# = NYEstSum#
      FundRevNYReqSum# = NYReqSum#
      FundRevNYRecSum# = NYRecSum#
      FundRevNYAppSum# = NYAppSum#
      
      BgtSum# = 0
      YTDSum# = 0
      PYActSum# = 0
      NYEstSum# = 0
      NYReqSum# = 0
      NYRecSum# = 0
      NYAppSum# = 0
      
      '--Now process expenditures ********************************
      DoingExp = True   '-Flag used in printing report header
      Category$ = " Expenditures"
      
      IF DeptOnNewPage = False THEN
        PRINT #PrnFile, FF$
        NewPage = True  '-Flag used in printing report header
        GOSUB PrintBGTWksPageHeader
      END IF
      
      'initialize dept variables
      Break = 0
      DeptBgtSum# = 0
      DeptYTDSum# = 0
      DeptPYActSum# = 0
      DeptNYEstSum# = 0
      DeptNYReqSum# = 0
      DeptNYRecSum# = 0
      DeptNYAppSum# = 0

      DeptBgtSum1# = 0
      DeptYTDSum1# = 0
      DeptPYActSum1# = 0
      DeptNYEstSum1# = 0
      DeptNYReqSum1# = 0
      DeptNYRecSum1# = 0
      DeptNYAppSum1# = 0

      LastDept$ = ""
      LastDeptName$ = ""


      '--Search exp account list for accounts in this fund
      FOR Cnt = 1 TO ECnt
        Rec = ExpAccts%(Cnt)
        GET AcctFileNum, Rec, Acct
        FundCode$ = LEFT$(Acct.Num, FundLen)
        ObjCode$ = MID$(Acct.Num, FundLen% + AcctLen% + 3, DetLen)

        IF FundCode$ = ThisFund$ THEN
          Account$ = QPTRIM$(Acct.Num) + "  " + QPTRIM$(Acct.Title)

          '--Extract the Dept$ from the G/L Acct
          Dept$ = MID$(Acct.Num, FundLen + 2, AcctLen)

          '--Get the Department name from the Department name file
          IF Dept$ <> LastDept$ THEN
            DeptRecNum = FindDept(Dept$)
            IF DeptRecNum > 0 THEN
              DeptName$ = QPTRIM$(GetDeptTitle$(DeptRecNum))
            ELSE
              DeptName$ = "Department " + Dept$
            END IF
          END IF

          '--Print Department Header first time thru
          IF LEN(LastDeptName$) = 0 THEN
            LastDeptName$ = DeptName$
            LastDept$ = Dept$
            GOSUB PrintDeptHeader
            IF LineCnt > MaxLines THEN
              PRINT #PrnFile, FF$
              GOSUB PrintBGTWksPageHeader
            END IF
          END IF

         '--Put Budget Summary Breaks Here
         IF BreakFlag THEN
           IF Break = 0 THEN
            Break = 1
            PRINT #PrnFile, Desc$(Break)
           END IF
          IF VAL(ObjCode$) > BEnd(Break) THEN
           GOSUB PrintBreakEnds
           'Check to See Where This One Falls
           FOR LL = 1 TO 9
           IF VAL(ObjCode$) > Beg(LL) AND VAL(ObjCode$) < BEnd(LL) THEN
           Break = LL
           EXIT FOR
           END IF
           NEXT LL
           PRINT #PrnFile, Desc$(Break)
          END IF
         END IF

          '--see if we need to subtotal dept
          IF LEN(LastDept$) > 0 THEN
            IF Dept$ <> LastDept$ THEN
                  IF BreakFlag THEN
                        GOSUB PrintBreakEnds
                        Break = 0
                   END IF
              GOSUB PrintDeptTotals
              GOSUB PrintDeptHeader
                  IF BreakFlag THEN
                        Break = 1
                        PRINT #PrnFile, Desc$(Break)
                  END IF

              IF LineCnt > MaxLines THEN
                PRINT #PrnFile, FF$
                GOSUB PrintBGTWksPageHeader
              END IF
            END IF
          END IF
          
          ToPrint$ = SPACE$(132)
          LSET ToPrint$ = LEFT$(Account$, 36)
          MID$(ToPrint$, Col1) = FUsing$(STR$(Acct.Bgt), BgtFmt$)
          MID$(ToPrint$, Col2) = FUsing$(STR$(Acct.YTD), CommaFmt$)
          MID$(ToPrint$, Col3) = FUsing$(STR$(Acct.PYAct), CommaFmt$)
          IF WorkSheet = False THEN
            MID$(ToPrint$, Col4) = FUsing$(STR$(Acct.NYEst), CommaFmt$)
            MID$(ToPrint$, Col5) = FUsing$(STR$(Acct.NYReq), CommaFmt$)
            MID$(ToPrint$, Col6) = FUsing$(STR$(Acct.NYRec), CommaFmt$)
            MID$(ToPrint$, Col7) = FUsing$(STR$(Acct.NYApp), CommaFmt$)
          ELSE
            IF Acct.NYEst <> 0 THEN
              MID$(ToPrint$, Col4) = FUsing$(STR$(Acct.NYEst), CommaFmt$)
            ELSE
              MID$(ToPrint$, Col4) = Underline$
            END IF
            IF Acct.NYReq <> 0 THEN
              MID$(ToPrint$, Col5) = FUsing$(STR$(Acct.NYReq), CommaFmt$)
            ELSE
              MID$(ToPrint$, Col5) = Underline$
            END IF
            IF Acct.NYRec <> 0 THEN
              MID$(ToPrint$, Col6) = FUsing$(STR$(Acct.NYRec), CommaFmt$)
            ELSE
              MID$(ToPrint$, Col6) = Underline$
            END IF
            IF Acct.NYApp <> 0 THEN
              MID$(ToPrint$, Col7) = FUsing$(STR$(Acct.NYApp), CommaFmt$)
            ELSE
              MID$(ToPrint$, Col7) = Underline$
            END IF

          END IF
          PRINT #PrnFile, ToPrint$
          LineCnt = LineCnt + 1
          IF LineCnt > MaxLines THEN
            PRINT #PrnFile, FF$
            GOSUB PrintBGTWksPageHeader
          END IF
          
          BgtSum# = BgtSum# + Acct.Bgt
          YTDSum# = YTDSum# + Acct.YTD
          PYActSum# = PYActSum# + Acct.PYAct
          NYEstSum# = NYEstSum# + Acct.NYEst
          NYReqSum# = NYReqSum# + Acct.NYReq
          NYRecSum# = NYRecSum# + Acct.NYRec
          NYAppSum# = NYAppSum# + Acct.NYApp
          
          DeptBgtSum# = DeptBgtSum# + Acct.Bgt
          DeptYTDSum# = DeptYTDSum# + Acct.YTD
          DeptPYActSum# = DeptPYActSum# + Acct.PYAct
          DeptNYEstSum# = DeptNYEstSum# + Acct.NYEst
          DeptNYReqSum# = DeptNYReqSum# + Acct.NYReq
          DeptNYRecSum# = DeptNYRecSum# + Acct.NYRec
          DeptNYAppSum# = DeptNYAppSum# + Acct.NYApp

          DeptBgtSum1# = DeptBgtSum1# + Acct.Bgt
          DeptYTDSum1# = DeptYTDSum1# + Acct.YTD
          DeptPYActSum1# = DeptPYActSum1# + Acct.PYAct
          DeptNYEstSum1# = DeptNYEstSum1# + Acct.NYEst
          DeptNYReqSum1# = DeptNYReqSum1# + Acct.NYReq
          DeptNYRecSum1# = DeptNYRecSum1# + Acct.NYRec
          DeptNYAppSum1# = DeptNYAppSum1# + Acct.NYApp
          
          LastDept$ = Dept$
          LastDeptName$ = DeptName$
        
        END IF
      NEXT Cnt
         IF BreakFlag THEN
           GOSUB PrintBreakEnds
         END IF


      '--Summarize last Dept after loop
      GOSUB PrintDeptTotals

      '--Now summarize all expenditures for fund
      GOSUB PrintSummaryLines   'Print dashed line after last
      ToPrint$ = SPACE$(132)
      LSET ToPrint$ = "Total Expenditures for Fund:"
      MID$(ToPrint$, Col1) = FUsing$(STR$(BgtSum#), BgtFmt$)
      MID$(ToPrint$, Col2) = FUsing$(STR$(YTDSum#), CommaFmt$)
      MID$(ToPrint$, Col3) = FUsing$(STR$(PYActSum#), CommaFmt$)
      IF WorkSheet = False THEN
        MID$(ToPrint$, Col4) = FUsing$(STR$(NYEstSum#), CommaFmt$)
        MID$(ToPrint$, Col5) = FUsing$(STR$(NYReqSum#), CommaFmt$)
        MID$(ToPrint$, Col6) = FUsing$(STR$(NYRecSum#), CommaFmt$)
        MID$(ToPrint$, Col7) = FUsing$(STR$(NYAppSum#), CommaFmt$)
      ELSE
            IF NYEstSum# <> 0 THEN
              MID$(ToPrint$, Col4) = FUsing$(STR$(NYEstSum#), CommaFmt$)
            ELSE
              MID$(ToPrint$, Col4) = Underline$
            END IF
            IF NYReqSum# <> 0 THEN
              MID$(ToPrint$, Col5) = FUsing$(STR$(NYReqSum#), CommaFmt$)
            ELSE
              MID$(ToPrint$, Col5) = Underline$
            END IF
            IF NYRecSum# <> 0 THEN
              MID$(ToPrint$, Col6) = FUsing$(STR$(NYRecSum#), CommaFmt$)
            ELSE
              MID$(ToPrint$, Col6) = Underline$
            END IF
            IF NYAppSum# <> 0 THEN
              MID$(ToPrint$, Col7) = FUsing$(STR$(NYAppSum#), CommaFmt$)
            ELSE
              MID$(ToPrint$, Col7) = Underline$
            END IF

      END IF
      PRINT #PrnFile, ToPrint$

      '--Summarize balances
      BGTBal# = Round#(FundRevBgt# - BgtSum#)
      YTDBal# = Round#(FundRevYTD# - YTDSum#)
      PYActBal# = FundRevPYActSum# - PYActSum#
      NYEstBal# = FundRevNYEstSum# - NYEstSum#
      NYReqBal# = FundRevNYReqSum# - NYReqSum#
      NYRecBal# = FundRevNYRecSum# - NYRecSum#
      NYAppBal# = FundRevNYAppSum# - NYAppSum#
      PRINT #PrnFile,

      ToPrint$ = SPACE$(132)
      LSET ToPrint$ = "Revenues Over/(Under) Expenditures"
      MID$(ToPrint$, Col1) = FUsing$(STR$(BGTBal#), BgtFmt$)
      MID$(ToPrint$, Col2) = FUsing$(STR$(YTDBal#), CommaFmt$)
      MID$(ToPrint$, Col3) = FUsing$(STR$(PYActBal#), CommaFmt$)
      IF WorkSheet = False THEN
        MID$(ToPrint$, Col4) = FUsing$(STR$(NYEstBal#), CommaFmt$)
        MID$(ToPrint$, Col5) = FUsing$(STR$(NYReqBal#), CommaFmt$)
        MID$(ToPrint$, Col6) = FUsing$(STR$(NYRecBal#), CommaFmt$)
        MID$(ToPrint$, Col7) = FUsing$(STR$(NYAppBal#), CommaFmt$)
      ELSE
            IF NYEstBal# <> 0 THEN
              MID$(ToPrint$, Col4) = FUsing$(STR$(NYEstBal#), CommaFmt$)
            ELSE
              MID$(ToPrint$, Col4) = Underline$
            END IF
            IF NYReqBal# <> 0 THEN
              MID$(ToPrint$, Col5) = FUsing$(STR$(NYReqBal#), CommaFmt$)
            ELSE
              MID$(ToPrint$, Col5) = Underline$
            END IF
            IF NYRecBal# <> 0 THEN
              MID$(ToPrint$, Col6) = FUsing$(STR$(NYRecBal#), CommaFmt$)
            ELSE
              MID$(ToPrint$, Col6) = Underline$
            END IF
            IF NYAppBal# <> 0 THEN
              MID$(ToPrint$, Col7) = FUsing$(STR$(NYAppBal#), CommaFmt$)
            ELSE
              MID$(ToPrint$, Col7) = Underline$
            END IF

      END IF
      PRINT #PrnFile, ToPrint$

      '--Reset variables for next fund
      FundRevBgt# = 0
      FundRevYTD# = 0
      FundRevPYSum# = 0
      FundRevNYEstSum# = 0
      FundRevNYReqSum# = 0
      FundRevNYRecSum# = 0
      FundRevNYAppSum# = 0
      BgtSum# = 0
      YTDSum# = 0
      PYActSum# = 0
      NYEstSum# = 0
      NYReqSum# = 0
      NYRecSum# = 0
      NYAppSum# = 0
      
      DeptBgtSum# = 0
      DeptYTDSum# = 0
      DeptPYSum# = 0
      DeptNYEstSum# = 0
      DeptNYReqSum# = 0
      DeptNYRecSum# = 0
      DeptNYAppSum# = 0

    END IF '--Using fund test
  NEXT '--fund

  PRINT #PrnFile, FF$

  CLOSE

  '--End Report Creation

  SELECT CASE Dev$
    CASE "S"
      EntryPoint = 2
    CASE "P"
      EntryPoint = 5
  END SELECT
  PrintRptFile Header$, ReportFile$, LPTNo, RetCode, EntryPoint

  CLOSE
  'KILL ReportFile$

EXIT SUB

PrintBGTWksPageHeader:
  PRINT #PrnFile, "Budget Preparation"
  PRINT #PrnFile, FundName$ + Category$
  PRINT #PrnFile, "Period Ending: " + Num2Date$(EndDate)
  PRINT #PrnFile,
  PRINT #PrnFile, "Description                             Budget         Actual     Prior Year         Est          Req          Rec      Apprv'd"
  PRINT #PrnFile, "-------------------------------------------------------------------------------------------------------------------------------"
  IF DoingExp THEN
    IF NewPage = False THEN
      PRINT #PrnFile, DeptName$ + " Continued"
      LineCnt = 7
    ELSE
      LineCnt = 6
    END IF
  ELSE
    LineCnt = 6
  END IF
RETURN


PrintSummaryLines:
  '--Print summary lines
  ToPrint$ = SPACE$(132)
  MID$(ToPrint$, Col1) = BSumLine$
  MID$(ToPrint$, Col2 + 1) = SumLine$
  MID$(ToPrint$, Col3 + 1) = SumLine$
  MID$(ToPrint$, Col4 + 1) = SumLine$
  MID$(ToPrint$, Col5 + 1) = SumLine$
  MID$(ToPrint$, Col6 + 1) = SumLine$
  MID$(ToPrint$, Col7 + 1) = SumLine$
  PRINT #PrnFile, ToPrint$
  LineCnt = LineCnt + 1
RETURN

PrintBreakEnds:
  GOSUB PrintSummaryLines
  DeptSummary$ = " Totals " + Desc$(Break)
  ToPrint$ = SPACE$(132)
  LSET ToPrint$ = QPTRIM$(DeptSummary$)
  MID$(ToPrint$, Col1) = FUsing$(STR$(DeptBgtSum1#), BgtFmt$)
  MID$(ToPrint$, Col2) = FUsing$(STR$(DeptYTDSum1#), CommaFmt$)
  MID$(ToPrint$, Col3) = FUsing$(STR$(DeptPYActSum1#), CommaFmt$)
  IF WorkSheet = False THEN
    MID$(ToPrint$, Col4) = FUsing$(STR$(DeptNYEstSum1#), CommaFmt$)
    MID$(ToPrint$, Col5) = FUsing$(STR$(DeptNYReqSum1#), CommaFmt$)
    MID$(ToPrint$, Col6) = FUsing$(STR$(DeptNYRecSum1#), CommaFmt$)
    MID$(ToPrint$, Col7) = FUsing$(STR$(DeptNYAppSum1#), CommaFmt$)
  ELSE

  END IF
  PRINT #PrnFile, ToPrint$
  PRINT #PrnFile,
  LineCnt = LineCnt + 2
  DeptBgtSum1# = 0
  DeptYTDSum1# = 0
  DeptPYActSum1# = 0
  DeptNYEstSum1# = 0
  DeptNYReqSum1# = 0
  DeptNYRecSum1# = 0
  DeptNYAppSum1# = 0

RETURN
PrintDeptTotals:

  GOSUB PrintSummaryLines
  DeptSummary$ = LastDeptName$ + " Totals"
  ToPrint$ = SPACE$(132)
  LSET ToPrint$ = DeptSummary$
  MID$(ToPrint$, Col1) = FUsing$(STR$(DeptBgtSum#), BgtFmt$)
  MID$(ToPrint$, Col2) = FUsing$(STR$(DeptYTDSum#), CommaFmt$)
  MID$(ToPrint$, Col3) = FUsing$(STR$(DeptPYActSum#), CommaFmt$)
  IF WorkSheet = False THEN
    MID$(ToPrint$, Col4) = FUsing$(STR$(DeptNYEstSum#), CommaFmt$)
    MID$(ToPrint$, Col5) = FUsing$(STR$(DeptNYReqSum#), CommaFmt$)
    MID$(ToPrint$, Col6) = FUsing$(STR$(DeptNYRecSum#), CommaFmt$)
    MID$(ToPrint$, Col7) = FUsing$(STR$(DeptNYAppSum#), CommaFmt$)
  ELSE
      IF DeptNYEstSum# <> 0 THEN
        MID$(ToPrint$, Col4) = FUsing$(STR$(DeptNYEstSum#), CommaFmt$)
      ELSE
        MID$(ToPrint$, Col4) = Underline$
      END IF
      IF DeptNYReqSum# <> 0 THEN
        MID$(ToPrint$, Col5) = FUsing$(STR$(DeptNYReqSum#), CommaFmt$)
      ELSE
        MID$(ToPrint$, Col5) = Underline$
      END IF
      IF DeptNYRecSum# <> 0 THEN
        MID$(ToPrint$, Col6) = FUsing$(STR$(DeptNYRecSum#), CommaFmt$)
      ELSE
        MID$(ToPrint$, Col6) = Underline$
      END IF
      IF DeptNYAppSum# <> 0 THEN
        MID$(ToPrint$, Col7) = FUsing$(STR$(DeptNYAppSum#), CommaFmt$)
      ELSE
        MID$(ToPrint$, Col7) = Underline$
      END IF

  END IF
  PRINT #PrnFile, ToPrint$
  PRINT #PrnFile,
  LineCnt = LineCnt + 2
  DeptBgtSum# = 0
  DeptYTDSum# = 0
  DeptPYActSum# = 0
  DeptNYEstSum# = 0
  DeptNYReqSum# = 0
  DeptNYRecSum# = 0
  DeptNYAppSum# = 0
RETURN


PrintDeptHeader:
  IF DeptOnNewPage THEN
    PRINT #PrnFile, FF$
    NewPage = True  '--flag not to print Dept continued
    GOSUB PrintBGTWksPageHeader
    ToPrint$ = SPACE$(132)
    LSET ToPrint$ = DeptName$
    PRINT #PrnFile, ToPrint$
    LineCnt = LineCnt + 1
  ELSE
    '--Don't print header if not room for at least one trans
    IF LineCnt < 53 THEN
      NewPage = False
      ToPrint$ = SPACE$(132)
      LSET ToPrint$ = DeptName$
      PRINT #PrnFile, ToPrint$
      LineCnt = LineCnt + 1
    ELSE
      PRINT #PrnFile, FF$
      NewPage = True  '--flag not to print Dept continued
      GOSUB PrintBGTWksPageHeader
      ToPrint$ = SPACE$(132)
      LSET ToPrint$ = DeptName$
      PRINT #PrnFile, ToPrint$
      LineCnt = LineCnt + 1
    END IF
  END IF
RETURN


GotErr:
  Errorcode$ = STR$(ERR)
  SELECT CASE ERR
    CASE 70
      CLS
      QPrintRC "Access Denied.  Try again later.", 12, 1, 12
      QPrintRC "Error Code:" + Errorcode$, 13, 1, 12
      QPrintRC "Press any key exit.", 14, 1, 11
    CASE ELSE
      CLS
      QPrintRC "An Error has halted the system, Error Code: " + Errorcode$, 12, 1, 14
      QPrintRC "Press any key exit.", 12, 1, 14
   END SELECT
   K$ = INPUT$(1)
   EXIT SUB
RETURN


END SUB

