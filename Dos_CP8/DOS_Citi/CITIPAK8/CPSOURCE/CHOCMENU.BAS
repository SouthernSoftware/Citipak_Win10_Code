DECLARE SUB BlockClearScrn ()
'CPMENU.BAS

DEFINT A-Z

'$INCLUDE: 'GL.BI'

DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$) 'DECLARE SUB FileView (FileName$, ky, Action, FVI AS ANY, SEG Array)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION WEnvTest ()
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (t$, r%, c%, clr%)
DECLARE SUB ShowCursor ()
DECLARE FUNCTION WaitKey% ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PrintLoading ()
DECLARE SUB CashRecpt ()
DECLARE SUB CashDisb ()

'$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)

'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'

'$INCLUDE: 'SetCnf.BI'
''$INCLUDE: 'init.BI'

'Dim the choice array to the number of menu items
DIM MenuText$(1 TO 11)

 MenuText$(1) = "Business License         "
 MenuText$(2) = "Accounts Payable         "
 MenuText$(3) = "Cash Management System   "
 MenuText$(4) = "Fixed Assets             "
 MenuText$(5) = "General Ledger           "
 MenuText$(6) = "Payroll                  "
 MenuText$(7) = "Property Taxes           "
 MenuText$(8) = "Utility Billing System   "
 MenuText$(9) = "Vehicle Decals           "
' MenuText$(9) = "ESC Magnetic Media Disk  "
' MenuText$(10) = "401k Magnetic Media Disk "
 MenuText$(10) = "Exit to OS               "


MaxLen = 0     'Set menu width to zero
BoxBot = 17    'limit the box length to go no lower than line 20
Action = 0     '0 means stay in the menu until they select something
Choice = 1     'Pre-load choice to highlight

FOR Cnt = 1 TO UBOUND(MenuText$)   'Find max menu width
  TLen = LEN(MenuText$(Cnt))
  IF TLen > MaxLen THEN
    MaxLen = TLen
  END IF
NEXT

Row = ((25 - (UBOUND(MenuText$))) \ 2) + 1
Col = ((80 - MaxLen) \ 2) - 2

IF WEnvTest THEN
  FileExt$ = ".BAS"
ELSE
  FileExt$ = ".EXE"
END IF

LOCATE 1, 1, 0
'Help$ = "Citipak Financial Management System v8.1"
ShowCursor

DO

   BlockClearScrn

   TitleBox 3, Col, MaxLen + 3, "Citi-Pak Main Menu", Cnf
   TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
   
   LOCATE Row, Col, 0     'set upper left corner of menu, turn off the cursor

   CALL VertMenu(MenuText$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf)

   IF ky$ = CHR$(27) THEN EXIT DO 'choice = 0

   SELECT CASE Choice
      CASE 1
         IF Exist("armenu" + FileExt$) THEN
           RUN "armenu"
         END IF
      CASE 2
         IF Exist("apmenu" + FileExt$) THEN
            RUN "apmenu"
         END IF
      CASE 3
         IF Exist("cmmenu" + FileExt$) THEN
            RUN "cmmenu"
         END IF
      CASE 4
         IF Exist("famenu" + FileExt$) THEN
           RUN "famenu"
         END IF
      CASE 5
         IF Exist("glmenu" + FileExt$) THEN
            RUN "GLCOPYRG"
         END IF
      CASE 6
         IF Exist("prcopyrg" + FileExt$) THEN
           RUN "prcopyrg"
         END IF
      CASE 7
         IF Exist("tbcopyrg" + FileExt$) THEN
           RUN "tbcopyrg"
         END IF
      CASE 8
         IF Exist("ubmenu" + FileExt$) THEN
           RUN "ubmenu"
         ELSEIF Exist("wbmenu" + FileExt$) THEN
           RUN "wbmenu"
         END IF
      CASE 9
         IF Exist("dCmenu" + FileExt$) THEN
            RUN "dCmenu"
         END IF
      'CASE 9
      '   IF Exist("ESC2DISK" + FileExt$) THEN
      '      RUN "ESC2DISK"
      '   END IF
      'CASE 10
      '   IF Exist("401K2DSK" + FileExt$) THEN
      '      RUN "401K2DSK"
      '   END IF
      CASE 10
         EXIT DO
      
  END SELECT

LOOP

HideCursor
CLS
END

SUB BlockClearScrn
Temp1$ = STRING$(80, "±")
Temp2$ = STRING$(80, " ")

FOR Cnt = 2 TO 24
  QPrintRC Temp1$, Cnt, 1, 3
NEXT

RSET Temp2$ = DATE$ + " "
QPrintRC Temp2$, 1, 1, 112
LSET Temp2$ = "                    Citipak Financial Management System v8.1"
QPrintRC Temp2$, 25, 1, 112

END SUB

