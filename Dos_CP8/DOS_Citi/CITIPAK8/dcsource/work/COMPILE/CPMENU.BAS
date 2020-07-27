'CPMENU.BAS

DEFINT A-Z

'$INCLUDE: 'GL.BI'

'DECLARE SUB BCopy (FromSeg%, FromAddr%, ToSeg%, ToAddr%, NumBytes%, Dir%)
'DECLARE SUB BLPrint (LPTPort%, Text$, ErrCode%)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Date2Num% (Dat$)
'DECLARE SUB FileView (FileName$, ky, Action, FVI AS ANY, SEG Array)
DECLARE FUNCTION FUsing$ (Number$, Mask$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION MGetKey% (Row%, Col%)
DECLARE FUNCTION Monitor% ()
'DECLARE SUB MScrnSave (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
'DECLARE SUB MScrnRest (ULRow%, ULCol%, LRRow%, LRCol%, SEG Element%)
'DECLARE SUB MPaintBox (ULRow%, ULCol%, LRRow%, LRCol%, Colr%)
'DECLARE FUNCTION Num2Date$ (Dat%)
'DECLARE FUNCTION Peek1% (Segment%, Address%)
DECLARE FUNCTION WEnvTest ()
DECLARE SUB QPrint (X$, Colr%, page%)
DECLARE SUB QPrintRC (t$, r%, c%, clr%)
'DECLARE FUNCTION QPTrim$ (X$)
DECLARE SUB ShowCursor ()
'DECLARE SUB SortT2 (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
'DECLARE SUB SortT (SEG Element AS ANY, NumElements%, Direction%, StructSize%, MemberOff%, MemberSize%)
'DECLARE SUB TextCursor (FG%, BG%)
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
DIM MenuText$(1 TO 10)
MenuText$(1) = "Accounts Payable"
MenuText$(2) = "Business License"
MenuText$(3) = "Cash Management System"
MenuText$(4) = "Fixed Assets"
MenuText$(5) = "General Ledger"
MenuText$(6) = "Payroll"
MenuText$(7) = "Property Taxes"
MenuText$(8) = "Utility Billing System "
MenuText$(9) = "Vehicle Decals"
MenuText$(10) = "Exit"

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
Row = ((25 - (UBOUND(MenuText$))) \ 2) + 1
Col = ((80 - MaxLen) \ 2) - 2

IF WEnvTest THEN
  FileExt$ = ".BAS"
ELSE
  FileExt$ = ".EXE"
END IF

LOCATE 1, 1, 0
Help$ = "Citipak Financial Management System v8.0"
ShowCursor

DO

   LibFile2Scrn "CM.QSL", "MENUBAK", MonoCode%, Attribute%, ErrorCode%

   TitleBox 3, Col, MaxLen + 3, "Citi-Pak Main Menu", Cnf
   TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

   PrintTitle User$
   PrintHelp Help$

   LOCATE Row, Col, 0     'set upper left corner of menu, turn off the cursor
   'LOCATE 2, 1, 0     'set upper left corner of menu, turn off the cursor
   CALL VertMenu(MenuText$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf)

   IF ky$ = CHR$(27) THEN EXIT DO 'choice = 0

   SELECT CASE Choice
      CASE 1
         PrintLoading
         IF Exist("apmenu" + FileExt$) THEN
            RUN "apmenu"
         END IF

      CASE 2
         PrintLoading
         IF Exist("armenu" + FileExt$) THEN
           RUN "armenu"
         END IF

      CASE 3
         PrintLoading
         IF Exist("cmmenu" + FileExt$) THEN
            RUN "cmmenu"
         END IF
       CASE 4
         PrintLoading
         IF Exist("famenu" + FileExt$) THEN
           RUN "famenu"
         END IF
      
      CASE 5
         PrintLoading
         IF Exist("glmenu" + FileExt$) THEN
            RUN "GLMENU"
         END IF

      CASE 6
         PrintLoading
         IF Exist("prcopyrg" + FileExt$) THEN
           RUN "prcopyrg"
         END IF

      CASE 7
         PrintLoading
         IF Exist("tbcopyrg" + FileExt$) THEN
           RUN "tbcopyrg"
         END IF

      CASE 8
         PrintLoading
         IF Exist("wbcopyrg" + FileExt$) THEN
           RUN "wbcopyrg"
         END IF
         IF Exist("ubmenu" + FileExt$) THEN
          RUN "ubmenu"
         END IF
      CASE 9
         PrintLoading
         IF Exist("dcmenu" + FileExt$) THEN
            RUN "dcmenu"
         END IF
      
      CASE 10
         EXIT DO
      
  END SELECT

LOOP

HideCursor
CLS
END

SUB PrintLoading

   QPrintRC " ", 1, 1, 119
   QPrintRC SPACE$(80), 25, 2, 112
   QPrintRC "Loading", 25, 2, 112
   LOCATE 1, 1

END SUB

