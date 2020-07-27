'GLMAINT.BAS
DEFINT A-Z

'$INCLUDE: 'GL.BI'
'$INCLUDE: 'PRO7.BI'
'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'SCRAUX.BI'
''$INCLUDE: 'QSCR.BI'

'DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB MnuTitle (Title$, X%, y%, Wdth%, TitColor%)
DECLARE SUB ClrBG ()
'DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintLoading ()
DECLARE SUB SetPostDates ()

DIM User$

'$INCLUDE: 'SetCnf.BI'
'$INCLUDE: 'init.BI'
CONST False = 0, True = NOT False

'Dim the choice array to the number of menu items
DIM MenuText$(1 TO 9)

MenuText$(1) = "Fund Maintenance"
MenuText$(2) = "Chart of Accounts"
MenuText$(3) = "Department Maintenance"
MenuText$(4) = "Bank Maintenance"
MenuText$(5) = "Post External Transactions"
MenuText$(6) = "Set Allowable Posting Dates"
MenuText$(7) = "G/L Closing Operations"
MenuText$(8) = "System Configuration & Utilities"
MenuText$(9) = "Quit"

MaxLen = 0     'Set menu width to zero
BoxBot = 20    'limit the box length to go no lower than line 20
Action = 0     '0 means stay in the menu until they select something
Choice = 1     'Pre-load choice to highlight

FOR cnt = 1 TO UBOUND(MenuText$)   'Find max menu width
  TLen = LEN(MenuText$(cnt))
  IF TLen > MaxLen THEN
    MaxLen = TLen
  END IF
NEXT

'Center Menu within Screen
Row = ((25 - (UBOUND(MenuText$))) \ 2) - 1
Col = ((80 - MaxLen) \ 2) - 2

LOCATE 1, 1, 0
Help$ = "General Ledger Maintenance"
ShowCursor

DO

   LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode%, Attribute%, ErrorCode%

   'TitleBox 3, Col, MaxLen + 3, "Setup Menu", Cnf
   'TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

   Title$ = "G/L Maintenance Menu"
   MnuTitle Title$, Row, Col, MaxLen, 15
   
   QPrintRC User$, 1, 2, -1
   QPrintRC DATE$, 1, 70, -1
   QPrintRC Help$, 25, 2, -1

   LOCATE Row, Col, 0     'set upper left corner of menu, turn off the cursor
   
   CALL VertMenu(MenuText$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf)

   IF Ky$ = CHR$(27) THEN EXIT DO 'choice = 0

   SELECT CASE Choice
      CASE 1
         ClrBG
         'PrintLoading
         RUN "GLFUND"
      CASE 2
         ClrBG
         'PrintLoading
         RUN "GLACCT"
      CASE 3
         ClrBG
         'PrintLoading
         RUN "GLDEPT"
      CASE 4
         ClrBG
         'ok = MsgBox("GL", "NA")
         'PrintLoading
         RUN "GLBANK"
      CASE 5
         ClrBG
         IF Exist("GLIFBA.EXE") THEN
          RUN "GLIFBA"
         ELSE
          ok = MsgBox("GL", "NA")
         END IF
      CASE 6
         ClrBG
         SetPostDates
      CASE 7
         ClrBG
         PrintLoading
         RUN "GLCLOSE"
      CASE 8
         ClrBG
         PrintLoading
         RUN "GLSYS"
      CASE 9
         CLS
         END
  END SELECT

LOOP

ClrBG
HideCursor
'END
PrintLoading
RUN "GLMENU"

SUB PrintLoading

   QPrintRC " ", 1, 1, 119
   QPrintRC SPACE$(80), 25, 2, 112
   QPrintRC "Loading", 25, 2, 112
   LOCATE 1, 1

END SUB

SUB SetPostDates
   
   REDIM Form$(0, 0)                         'Holds data from form
   REDIM Fld(0) AS FieldInfo                 'Field editing array
   DIM Frm AS FormInfo                       'Dim the form info array
   LibLoadDisplayForm "GL.QSL", "POSTDATE", Form$(), Fld()
   PrintHelp "Set Allowable Posting Dates"
   
   DIM GLSetup AS GLSetupRecType
   SetupRecLen = LEN(GLSetup)
   GOSUB OpenSetupFile
   GET SetupFile, 1, GLSetup
   Form$(1, 0) = Num2Date$(GLSetup.LPDate)
   Form$(2, 0) = Num2Date$(GLSetup.HPDate)
   CLOSE

   Action = 1                                'Init form edit
   DO                                        'Edit the form
      EditForm Form$(), Fld(), Frm, Cnf, Action
      SELECT CASE Frm.KeyCode                'Respond to the key presses
         CASE F10Key                            'F10=Save
            GOSUB OpenSetupFile
            GLSetup.LPDate = Date2Num(Form$(1, 0))
            GLSetup.HPDate = Date2Num(Form$(2, 0))
            PUT SetupFile, 1, GLSetup
            CLOSE
            ExitFlag = True
         CASE EscKey
            ExitFlag = True
      END SELECT
   LOOP UNTIL ExitFlag

EXIT SUB

OpenSetupFile:
   SetupFile = FREEFILE
   OPEN "GLSETUP.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS SetupFile LEN = SetupRecLen
RETURN


END SUB

