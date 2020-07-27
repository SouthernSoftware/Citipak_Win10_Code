DEFINT A-Z

DECLARE SUB ClearBack ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB printhelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB HideCursor ()
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf AS ANY)

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'SetCnf.BI'

  CONST False = 0, True = NOT False

  STACK 8000

  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF
  RunFile$ = "CPMENU" + Ext$

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 8)

  MChoice$(1) = " Add/Edit Inventory Items "
  MChoice$(2) = " Inventory Item Transactions "
  MChoice$(3) = " Inventory Item Reports"
  MChoice$(4) = " Exit to OS"

  MaxLen = 0     'Set menu width to zero
  BoxBot = 17    'limit the box length to go no lower than line 20
  Action = 0     '0 means stay in the menu until they select something
  Choice = 1     'Pre-load choice to highlight

  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT

  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) + 1
  Col = ((80 - MaxLen) \ 2)
  Help$ = "Business License Main Menu"

  DO

     '--Set upper left corner of menu, turn off the cursor
     LOCATE Row, Col, 0

     ClearBack

     TitleBox 3, Col, MaxLen + 3, "Item Inventory Main Menu", Cnf
     TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

     ShowCursor

     VertMenu MChoice$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf

     IF ky$ = CHR$(27) THEN EXIT DO 'choice = 0

     SELECT CASE Choice
     CASE 1
       RUN "ICITEM"
     CASE 2
       RUN "ICTRANS"
     CASE 3
       RUN "ICREPORT"
     CASE 4
       HideCursor
       CLS
       END
     END SELECT
  LOOP


  IF Exist(RunFile$) THEN
    RUN RunFile$
  ELSE
    HideCursor
    CLS
    END
  END IF

