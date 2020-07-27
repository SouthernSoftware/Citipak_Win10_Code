'APMENU.BAS
DEFINT A-Z

'$INCLUDE: 'DefCnf.BI'
  
DECLARE SUB OptionalMod ()
DECLARE SUB ClrBG ()
DECLARE SUB MnuTitle (Title$, X%, y%, Wdth%, TitColor%)
DECLARE SUB VendorMenu ()
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB QPrintRC (T$, r%, c%, clr%)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE FUNCTION Exist (Filename$)
  
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'gl.BI'
  '$INCLUDE: 'init.BI'
  
  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 6)
  
  MChoice$(1) = "Vendor Maintenance"
  MChoice$(2) = "Purchase Order Processing"
  MChoice$(3) = "Invoice Processing"
  MChoice$(4) = "Check Processing"
  MChoice$(5) = "Reports"
  MChoice$(6) = "Quit"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 20   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  '--Find max menu width
  FOR Cnt = 1 TO UBOUND(MChoice$)
    TLen = LEN(MChoice$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  '--Center Menu within Screen
  Row = ((25 - (UBOUND(MChoice$))) \ 2) - 1
  Col = ((80 - MaxLen) \ 2) - 2
  Help$ = "Accounts Payable Menu"
  
  DO
    LOCATE Row, Col, 0
    LibFile2Scrn "GL.QSL", "MENUBAK", MonoCode, -1, ErrorCode
    
    Title$ = "Accounts Payable Menu  "
    MnuTitle Title$, Row, Col, MaxLen, 15
    PrintTitle User$
    PrintHelp Help$
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf
    IF ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    SELECT CASE Choice
    CASE 1
      ClrBG
      RUN "APV"
    CASE 2
      ClrBG
      'OptionalMod
      RUN "APPO"
    CASE 3
      ClrBG
      RUN "APINV1"
    CASE 4
      ClrBG
      RUN "APCHECK"
    CASE 5
      ClrBG
      RUN "APRPT1"
    CASE 6
      ClrBG
      HideCursor
      CLS
      END
    END SELECT
  LOOP
  
  ClrBG
  HideCursor
  RUN "CPMENU"
  END

SUB OptionalMod
  
  CLS
  QPrintRC "Optional Module.  Contact Software Sales @ 1-800-842-8190 for more info.", 1, 1, -1
  QPrintRC "Press any key to return to menu.", 2, 1, -1
  K$ = INPUT$(1)
  
END SUB

