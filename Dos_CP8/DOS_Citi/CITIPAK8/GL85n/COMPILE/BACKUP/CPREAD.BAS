'CPMENU.BAS
DEFINT A-Z

'$INCLUDE: 'GL.BI'

DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
DECLARE SUB FileView (FileName$, ky, Action, FVI AS ANY, SEG Array)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
DECLARE FUNCTION Num2Date$ (Dat%)
DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE FUNCTION WEnvTest ()
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (t$, r%, c%, clr%)
DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
DECLARE SUB TextCursor (FG%, BG%)
DECLARE FUNCTION WaitKey% ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PrintLoading ()
DECLARE SUB CashRecpt ()
DECLARE SUB CashDisb ()
'DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB MnuTitle (Title$, X%, y%, Wdth%, TitColor%)
DECLARE SUB ClrBG ()

'$INCLUDE: 'DefCnf.BI'
'$INCLUDE: 'formedit.BI'
'$INCLUDE: 'fieldinf.BI'
'$INCLUDE: 'SetCnf.BI'
'$INCLUDE: 'init.BI'

'Dim the choice array to the number of menu items
DIM MenuText$(1 TO 10)

MenuText$(1) = "Accounts Receivable"
MenuText$(2) = "Accounts Payable"
MenuText$(3) = "Budgetary Accounting"
MenuText$(4) = "Cash Management System"
MenuText$(5) = "Decal Management System"
MenuText$(6) = "Fixed Assets"
MenuText$(7) = "Payroll"
MenuText$(8) = "Property Taxes"
MenuText$(9) = "Utility Billing System "
MenuText$(10) = "Quit"

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

'IF WEnvTest THEN
'  FileExt$ = ".BAS"
'ELSE
  FileExt$ = ".EXE"
'END IF

LOCATE 1, 1, 0
Help$ = "Citipak Financial Management System v8.0"
ShowCursor

DO

   LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode%, Attribute%, ErrorCode%

   'TitleBox 3, Col, MaxLen + 3, "Main Menu", Cnf
   'TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
   Title$ = "Main Menu "
   MnuTitle Title$, Row, Col, MaxLen, 15
   PrintTitle User$
   PrintHelp Help$

   LOCATE Row, Col, 0     'set upper left corner of menu, turn off the cursor
   'LOCATE 2, 1, 0     'set upper left corner of menu, turn off the cursor
   CALL VertMenu(MenuText$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf)

   IF ky$ = CHR$(27) THEN EXIT DO 'choice = 0

   SELECT CASE Choice
      CASE 1
         ClrBG
         PrintLoading
         IF Exist("armenu" + FileExt$) THEN
           RUN "armenu"
         END IF

      CASE 2
         ClrBG
         PrintLoading
         IF Exist("apmenu" + FileExt$) THEN
            RUN "apmenu"
         END IF

      CASE 3
         ClrBG
         PrintLoading
         IF Exist("GLRPTRD" + FileExt$) THEN
            RUN "GLRPTRD"
         END IF
      
      CASE 4
         ClrBG
         PrintLoading
         IF Exist("cmmenu" + FileExt$) THEN
            RUN "cmmenu"
         END IF

      CASE 5
         ClrBG
         PrintLoading
         IF Exist("dbcopyrg" + FileExt$) THEN
            RUN "dbcopyrg"
         END IF

      CASE 6
         ClrBG
         PrintLoading
         IF Exist("famenu" + FileExt$) THEN
           RUN "famenu"
         END IF
      
      CASE 7
         ClrBG
         PrintLoading
         IF Exist("prcopyrg" + FileExt$) THEN
           RUN "prcopyrg"
         END IF

      CASE 8
         ClrBG
         PrintLoading
         IF Exist("tbcopyrg" + FileExt$) THEN
           RUN "tbcopyrg"
         END IF

      CASE 9
         ClrBG
         PrintLoading
         'IF Exist("wbcopyrg" + FileExt$) THEN
           RUN "UBMENU"
         'END IF
      
      CASE 10
         EXIT DO
      
  END SELECT

LOOP

HideCursor
CLS
END

SUB PrintLoading

   QPrintRC " ", 1, 1, 119
   BG$ = STRING$(80, CHR$(177))
   FOR cnt = 2 TO 24
    QPrintRC BG$, cnt, 1, 3
   NEXT
   QPrintRC SPACE$(80), 25, 2, 112
   QPrintRC "Loading", 25, 2, 112
   LOCATE 1, 1

END SUB

