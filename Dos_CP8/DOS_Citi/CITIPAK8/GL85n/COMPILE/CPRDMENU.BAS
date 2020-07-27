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
DIM MenuText$(1 TO 6)

MenuText$(1) = "Accounts Payable"
MenuText$(2) = "Purchase Orders"
MenuText$(3) = "Budgetary Accounting"
MenuText$(4) = "Cash Management System"
MenuText$(5) = "Utility Billing System "
MenuText$(6) = "Quit"

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
  FileExt$ = ".BAS"
'ELSE
'  FileExt$ = ".EXE"
'END IF

LOCATE 1, 1, 0
Help$ = "Citipak System v8.2 Report Only Menu"
ShowCursor

DO

   LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode%, Attribute%, ErrorCode%

   Title$ = "Report Only Menu "
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
         IF Exist("APMENU" + FileExt$) THEN
            RUN "aprptrd"
         END IF

      CASE 2
         ClrBG
         PrintLoading
         IF Exist("APPOWB" + FileExt$) THEN
            RUN "appowb"
         END IF


      CASE 3
         ClrBG
         PrintLoading
         IF Exist("GLMENU" + FileExt$) THEN
            RUN "GLRPTRD"
         END IF
      
      CASE 4
         ClrBG
         PrintLoading
         IF Exist("CMMENU" + FileExt$) THEN
           'RUN "CMJOURRD"
         END IF
      CASE 5
         ClrBG
         PrintLoading
         IF Exist("UBMENU" + FileExt$) THEN
           'RUN "UBREPRD"
         END IF
      
      CASE 6
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

