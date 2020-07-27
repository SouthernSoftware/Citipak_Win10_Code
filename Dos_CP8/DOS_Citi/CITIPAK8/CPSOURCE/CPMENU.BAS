DEFINT A-Z
DECLARE SUB Init ()
DECLARE SUB BlockClearScrn ()
DECLARE FUNCTION Exist% (FileName$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION WEnvTest ()
DECLARE SUB QPrintRC (t$, r%, c%, clr%)
DECLARE SUB ShowCursor ()
DECLARE SUB PrintLoading ()
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB MnuTitle (Title$, x%, Y%, Wdth%, TitColor%)
  
  '$INCLUDE: 'gl.BI'
  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'SetCnf.BI'

  'Dim the choice array to the number of menu items
  DIM MenuText$(1 TO 16)
  
  MenuText$(1) = "Business License"
  MenuText$(2) = "Accounts Payable"
  MenuText$(3) = "General Ledger"
  MenuText$(4) = "Payroll Processing"
  MenuText$(5) = "Fixed Assets"
  MenuText$(6) = "Property Taxes"
  MenuText$(7) = "Inventory Control"
  MenuText$(8) = "Cash Management System "
  MenuText$(9) = "Utility Billing System "
  MenuText$(10) = "Vehicle Decals"
  MenuText$(11) = "Tax Assessments"
  MenuText$(12) = "W2 Processing"
  MenuText$(13) = "Void Posted Payroll Check "
  MenuText$(14) = "401k To Disk Report"
  MenuText$(15) = "ESC To Disk Report"
  MenuText$(16) = "Exit to OS"
  
  MaxLen = 0    'Set menu width to zero
  BoxBot = 19   'limit the box length to go no lower than line 20
  Action = 0    '0 means stay in the menu until they select something
  Choice = 1    'Pre-load choice to highlight
  
  FOR Cnt = 1 TO UBOUND(MenuText$)              'Find max menu width
    TLen = LEN(MenuText$(Cnt))
    IF TLen > MaxLen THEN
      MaxLen = TLen
    END IF
  NEXT
  
  Row = 6' ((25 - (UBOUND(MenuText$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 1
  
'  IF WEnvTest THEN
'    FileExt$ = ".BAS"
'  ELSE
    FileExt$ = ".EXE"
'  END IF
  
  LOCATE 1, 1, 0
  'Help$ = "Citipak Financial Management System v8.1"

  Title$ = "Main Menu  "
  
  ShowCursor

  DO
    
    BlockClearScrn
    
'    TitleBox 3, Col, MaxLen + 3, "Citi-Pak Main Menu", Cnf
'    TitleBox 20, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    LOCATE Row, Col, 0          'set upper left corner of menu, turn off the cursor

    Init

    MnuTitle Title$, Row, Col, MaxLen, 15

    CALL VertMenu(MenuText$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf)
    
    IF ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      IF Exist("armenu.EXE") THEN
        RUN "armenu"
      END IF
    CASE 2
      IF Exist("apmenu.EXE") THEN
        RUN "apmenu"
      END IF
    CASE 3
      IF Exist("glcopyrg.EXE") THEN
        RUN "GLCOPYRG"
      ELSEIF Exist("glmenu.EXE") THEN
        RUN "GLMENU"
      END IF
    CASE 4
      IF Exist("prcopyrg.EXE") THEN
        RUN "prcopyrg"
      END IF
    CASE 5
      IF Exist("famenu.EXE") THEN
        RUN "famenu.exe"
      END IF
    CASE 6
      IF Exist("TAXCPYRG.EXE") THEN
        RUN "TAXCPYRG"
      ELSEIF Exist("taxmenu.EXE") THEN
        RUN "taxmenu.exe"
      ELSEIF Exist("tbmenu.EXE") THEN
        RUN "tbmenu.exe"
      END IF
    CASE 7
      IF Exist("icmenu.EXE") THEN
        RUN "icmenu.exe"
      END IF
    CASE 8
      IF Exist("cmmenu.EXE") THEN
        RUN "cmmenu"
      END IF
    CASE 9
      IF Exist("ubcopyrg.EXE") THEN
        RUN "ubcopyrg"
      ELSEIF Exist("ubmenu.EXE") THEN
        RUN "ubmenu"
      ELSEIF Exist("wbmenu.EXE") THEN
        RUN "wbmenu"
      END IF
    CASE 10
      IF Exist("dcmenu.EXE") THEN
        RUN "dcmenu"
      END IF
    CASE 11
      IF Exist("camenu.exe") THEN
        RUN "camenu"
      END IF
    CASE 12
      RUN "PR-W2"
    CASE 13
      IF Exist("prvoidck.exe") THEN
        RUN "prvoidck"
      END IF
    CASE 14
      RUN "401K2DSK"
    CASE 15
      RUN "ESC2DISK"
    CASE 16
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
  LSET Temp2$ = " Citipak Financial Management System v8.1   ³  " + CHR$(24) + CHR$(25) + " to Select  Enter=Go  Esc=Exit"
  QPrintRC Temp2$, 25, 1, 112
  
END SUB

SUB Init

  DIM GLSetup AS GLSetupRecType
  SetupRecLen = LEN(GLSetup)

  OPEN "GLSETUP.DAT" FOR RANDOM AS #1 LEN = SetupRecLen

  IF LOF(1) > 0 THEN
    GET #1, 1, GLSetup
    User$ = " " + QPTrim$(GLSetup.UserName)
  ELSE
    User$ = " Undefined"
  END IF
  CLOSE
  QPrintRC User$, 1, 1, -1

END SUB

