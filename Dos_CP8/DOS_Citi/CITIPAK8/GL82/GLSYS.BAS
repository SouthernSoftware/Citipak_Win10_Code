DEFINT A-Z

'$INCLUDE: 'Fileview.bi'
'$INCLUDE: 'DefCnf.bi'
'$INCLUDE: 'FieldInf.bi'
'$INCLUDE: 'FormEdit.bi'
'$INCLUDE: 'PRO7.bi'
'$INCLUDE: 'QSCR.BI'
'$INCLUDE: 'GLAUX.BI'
'$INCLUDE: 'GL.BI'

DECLARE SUB UnderConst ()
DECLARE SUB OpenSetupFile (SetupFileNum%)
DECLARE SUB DefineAcctStruct ()
DECLARE SUB UserInfo ()
DECLARE SUB SetFYDates ()
'DECLARE SUB ReLink ()
'DECLARE SUB ExportFiles ()
DECLARE SUB DispLoad ()
DECLARE SUB LibLoadDisplayForm (LibName$, FormName$, Form$(), Fld() AS ANY)
'DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB MnuTitle (Title$, X%, y%, Wdth%, TitColor%)
DECLARE SUB ClrBG ()
DECLARE SUB ReLinkTrans ()
DECLARE SUB RelinkBgtTrans ()
DECLARE FUNCTION WEnvTest ()
DECLARE FUNCTION GetPassword$ (Row%, Col%, Max%, Abort%)

'$INCLUDE: 'SetCnf.BI'
''$INCLUDE: 'keycode.bi'

CONST False = 0, True = NOT False



   DIM SHARED GLSetup AS GLSetupRecType

   OpenSetupFile SetupFileNum
   GET SetupFileNum, 1, GLSetup
      User$ = GLSetup.UserName
   '   TotLen = GLSetup.TotAcctLen
   '   FundLen = GLSetup.FundLen
   '   AcctLen = GLSetup.AcctLen
   '   DetLen = GLSetup.DetLen
   CLOSE SetupFileNum


'***Master Menu

   DIM mChoice$(1 TO 5)

   mChoice$(1) = "User File Maintenance"
   mChoice$(2) = "Define G/L Account Structure"
   mChoice$(3) = "Set Fiscal Periods"
   'mChoice$(4) = "Set Allowable Posting Date Range"
   'mChoice$(4) = "Re-Link G/L Transaction Database"
   'mChoice$(5) = "Re-Link Budget Transaction Database"
   'mChoice$(6) = "Quit"
   mChoice$(4) = "G/L Utilities"
   mChoice$(5) = "Quit"


   MaxLen = 0   'Set menu width to zero
   BoxBot = 20    'limit the box length to go no lower than line 20
   Action = 0     '0 means stay in the menu until they select something
   Choice = 1     'Pre-load choice to highlight

   FOR Cnt = 1 TO UBOUND(mChoice$)   'Find max menu width
     TLen = LEN(mChoice$(Cnt))
       IF TLen > MaxLen THEN
         MaxLen = TLen
       END IF
   NEXT

   'Center Menu within Screen
   Row = ((25 - (UBOUND(mChoice$))) \ 2) - 1
   Col = ((80 - MaxLen) \ 2) - 2

   ShowCursor 'turned mouse off temporarily
   
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
            IF INSTR(COMMAND$, "TEST") > 0 THEN
              RUN "GLUTIL"
            ELSE
              IF NOT WEnvTest THEN
                LibFile2Scrn "GL.QSL", "GETPW", MonoCode, Attribute, ErrorCode
                PW$ = GetPassword$(12, 42, 10, Abort%)
                IF PW$ = "SOSOFT" THEN
                  RUN "GLUTIL"
                ELSE
                  BEEP
                END IF
              ELSE
                RUN "GLUTIL"
              END IF
            END IF
         CASE 5
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
   REDIM Form$(0, 0)                         'Holds data from form
   REDIM Fld(0) AS FieldInfo                 'Field editing array
   
   DIM Frm AS FormInfo                       'Dim the form info array
   LibLoadDisplayForm "GL.QSL", "SETUPGL", Form$(), Fld()
   PrintHelp "Define Account Structure"
   
   OpenSetupFile SetupFileNum

   GET SetupFileNum, 1, GLSetup
   
   Form$(1, 0) = STR$(GLSetup.FundLen)
   Form$(2, 0) = STR$(GLSetup.AcctLen)
   Form$(3, 0) = STR$(GLSetup.DetLen)
   Form$(4, 0) = STR$(GLSetup.TotAcctLen)

   CLOSE

   Action = 1                                'Init form edit

   DO                                        'Edit the form
     
      EditForm Form$(), Fld(), Frm, Cnf, Action
      
      SELECT CASE Frm.KeyCode                'Respond to the key presses
         CASE F10Key                            'F10=Save
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
   DIM Acct AS GLAcctRecType
   
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

   REDIM Form$(0, 0)                         'Holds data from form
   REDIM Fld(0) AS FieldInfo                 'Field editing array
   DIM Frm AS FormInfo                       'Dim the form info array
   LibLoadDisplayForm "GL.QSL", "POSTDATE", Form$(), Fld()
   PrintHelp "Set Allowable Posting Dates"

   OpenSetupFile SetupFileNum

   GET SetupFileNum, 1, GLSetup

   Form$(1, 0) = Num2Date$(GLSetup.LPDate)
   Form$(2, 0) = Num2Date$(GLSetup.HPDate)
   CLOSE

   Action = 1                                'Init form edit

   DO                                        'Edit the form

      EditForm Form$(), Fld(), Frm, Cnf, Action

      SELECT CASE Frm.KeyCode                'Respond to the key presses
         CASE F10Key                            'F10=Save

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

   SHARED Cnf AS Config                      'System config & colors
   REDIM Form$(0, 0)                         'Holds data from form
   REDIM Fld(0) AS FieldInfo                 'Field editing array

   DIM Frm AS FormInfo                       'Dim the form info array
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

   Action = 1                                'Init form edit

   DO                                        'Edit the form

      EditForm Form$(), Fld(), Frm, Cnf, Action

      SELECT CASE Frm.KeyCode                'Respond to the key presses
         CASE F10Key                            'F10=Save

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
ok$ = INPUT$(1)

END SUB

SUB UserInfo

   SHARED Cnf AS Config                      'System config & colors
   REDIM Form$(0, 0)                         'Holds data from form
   REDIM Fld(0) AS FieldInfo                 'Field editing array

   LibLoadDisplayForm "GL.QSL", "USERINF1", Form$(), Fld()   'userinfo

   DIM Frm AS FormInfo                       'Dim the form info array

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

   Action = 1                                'Init form edit

   DO                                        'Edit the form

      EditForm Form$(), Fld(), Frm, Cnf, Action

      SELECT CASE Frm.KeyCode                'Respond to the key presses
         CASE F10Key                            'F10=Save
            OpenSetupFile SetupFileNum
            GLSetup.UserName = Form$(1, 0)
            GLSetup.CashAcct = Form$(2, 0)
            GLSetup.APAcct = Form$(3, 0)
            GLSetup.EncAcct = Form$(4, 0)
            GLSetup.FBAcct = Form$(5, 0)
            GLSetup.CDCash = Form$(6, 0) 'new
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

