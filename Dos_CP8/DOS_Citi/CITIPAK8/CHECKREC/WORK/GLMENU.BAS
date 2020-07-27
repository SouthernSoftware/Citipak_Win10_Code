'GLMENU.BAS
DEFINT A-Z

'$INCLUDE: 'GL.BI'
'$INCLUDE: 'PRO7.BI'
'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'

DECLARE SUB CJMenu (JrnlType%)
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintLoading ()
DECLARE SUB FGetAH (FileName$, SEG Element AS ANY, ElSize%, NumEls%)
'DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS Config)
DECLARE SUB MnuTitle (Title$, X%, y%, Wdth%, TitColor%)
DECLARE SUB ClrBG ()
DECLARE FUNCTION WEnvTest ()

'$INCLUDE: 'init.BI'
'$INCLUDE: 'setCnf.BI'

'Dim the choice array to the number of menu items
REDIM MenuText$(1 TO 8)
MenuText$(1) = "Budget Maintenance"
MenuText$(2) = "General Journal"
MenuText$(3) = "Cash Receipts Journal"
MenuText$(4) = "Cash Disbursements Journal"
MenuText$(5) = "Reports"
MenuText$(6) = "Bank Reconciliation"
MenuText$(7) = "G/L Setup & Maintenance"
MenuText$(8) = "Quit"

MaxLen = 0     'Set menu width to zero
BoxBot = 17    'limit the box length to go no lower than line 20
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
Help$ = "Budgetary Accounting"

'IF WEnvTest THEN
'  FileExt$ = ".BAS"
'ELSE
  FileExt$ = ".EXE"
'END IF

ShowCursor

DO
   LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode%, Attribute%, ErrorCode%
   'TitleBox 3, Col, MaxLen + 3, "Budgetary Accounting", Cnf
   'TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
   Title$ = "Budgetary Accounting "
   MnuTitle Title$, Row, Col, MaxLen, 15
  
   QPrintRC User$, 1, 2, -1
   QPrintRC DATE$, 1, 70, -1
   QPrintRC Help$, 25, 2, -1

   LOCATE Row, Col, 0     'set upper left corner of menu, turn off the cursor
   'LOCATE 2, 1, 0     'set upper left corner of menu, turn off the cursor

   'Action = 3
   CALL VertMenu(MenuText$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf)

   IF ky$ = CHR$(27) THEN EXIT DO 'choice = 0

   SELECT CASE Choice
      CASE 1
         ClrBG
         'PrintLoading
         'IF Exist("BGTEDIT" + FileExt$) THEN
           RUN "BGTEDIT"
         'END IF
      CASE 2
         ClrBG
         'PrintLoading
         'IF Exist("GLGJED" + FileExt$) THEN
           RUN "GLGJED"
         'END IF
      CASE 3
         ClrBG
         PrintLoading
         CJMenu 1
      CASE 4
         ClrBG
         PrintLoading
         CJMenu 2
      CASE 5
         ClrBG
         'PrintLoading
         'IF Exist("GLRPT" + FileExt$) THEN
           RUN "GLRPT"
         'END IF
      CASE 6
         ClrBG
         'PrintLoading
         'IF Exist("CHECKREC" + FileExt$) THEN
           RUN "CHECKREC"
         'END IF
      CASE 7
         ClrBG
         'PrintLoading
         'IF Exist("GLMAINT" + FileExt$) THEN
           RUN "GLMAINT"
         'END IF
      CASE 8
         ClrBG
         CLS
         HideCursor
         END

  END SELECT

LOOP

ClrBG
HideCursor
PrintLoading
IF Exist("CPMENU" + FileExt$) THEN
  RUN "CPMENU"
END IF

SUB PrintLoading

   QPrintRC " ", 1, 1, 119
   QPrintRC SPACE$(80), 25, 2, 112
   QPrintRC "Loading", 25, 2, 112
   LOCATE 1, 1

END SUB

