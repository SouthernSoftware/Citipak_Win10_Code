DEFINT A-Z

DECLARE SUB UBCustomerMenu ()
DECLARE SUB ClearScrn ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB BlockClear ()
DECLARE SUB PrintTitle (Title$)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION Round# (DblNum#)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB ShowCursor ()
DECLARE SUB HideCursor ()

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'ubtrans.BI'
  
  STACK 5000
  
  IF Exist("UBRUN1ST.EXE") THEN
    RUN "UBRUN1ST"
  END IF

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 9)
  
  MChoice$(1) = "Customer Maintenance"
  MChoice$(2) = "Enter Payments and/or Deposits"
  MChoice$(3) = "Utility Billing, Readings, Penalties"
  MChoice$(4) = "Final Bill Processing Menu"
  MChoice$(5) = "Customer Reports"
  MChoice$(6) = "Statistical Reports"
  MChoice$(7) = "Work Order Processing"
  MChoice$(8) = "Misc. (Setup, Rate Tables, etc)"
  MChoice$(9) = "Exit to OS"

  MaxLen = 0    'Set menu width to zero
  BoxBot = 18   'limit the box length to go no lower than line 18
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2)
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0
    BlockClear
    
    TitleBox 2, Col, MaxLen + 3, "Main Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf

    ShowCursor
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0

    SELECT CASE Choice
    CASE 1
      UBCustomerMenu              'unrem
    CASE 2
      RUN "UBPAYMNT"
    CASE 3
      RUN "UBBillin"
    CASE 4
      RUN "UBFINBIL"
    CASE 5
      RUN "ubcustre"
    CASE 6
      RUN "ubstatre"
    CASE 7
      RUN "UBWrkOrd"
    CASE 8
      RUN "UBMisc"
    CASE 9
      HideCursor
      ClearScrn
      END
    END SELECT
  LOOP
  
  IF WEnvTest THEN
    Ext$ = ".bas"
  ELSE
    Ext$ = ".exe"
  END IF
  IF Exist("cpmenu" + Ext$) THEN
    RUN "cpmenu"
  ELSE
    HideCursor
    ClearScrn
  END IF

END
  
  

