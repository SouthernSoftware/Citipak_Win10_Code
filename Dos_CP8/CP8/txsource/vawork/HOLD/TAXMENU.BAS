DEFINT A-Z
DECLARE FUNCTION CheckInterestInfo% ()
DECLARE SUB TAXCustomerMenu ()
DECLARE SUB ClearBack ()
DECLARE SUB SendDist2GL ()
DECLARE SUB UBMiscMenu ()
DECLARE SUB UBBillMenu ()
DECLARE SUB UBCustomerMenu ()
DECLARE SUB ClearScrn ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()
DECLARE FUNCTION OK2UPDateCust% ()
DECLARE SUB UpDateTicklerFile ()

'$INCLUDE: 'DefCnf.BI'
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION Monitor% ()
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS Config)
DECLARE SUB HideCursor ()
DECLARE SUB CursorOff ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)

  CONST False = 0, True = NOT False

  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'
  '$INCLUDE: 'TaxCust.BI'

  STACK 5000

  LOCATE 1, 1, 0

 ' IF CheckInterestInfo THEN
 '   RUN "TAXREINT"
 ' END IF

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 11)
  
  MChoice$(1) = "Customer Maintenance"
  MChoice$(2) = "Enter Tax Payments"
  MChoice$(3) = "Real Estate - Tax Billing Functions "
  MChoice$(4) = "Pers Property - Tax Billing Functions "
  MChoice$(5) = "Tax Reporting System"
  MChoice$(6) = "Tax Billing Abatements"
  MChoice$(7) = "Manual Tax Bill Entry - Real Estate"
  MChoice$(8) = "Manual Tax Bill Entry - Personal Property"
  MChoice$(9) = "Tax System Setup & Utilities "
  MChoice$(10) = "DMV Processing"
  MChoice$(11) = "Exit to OS"
  
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
  Row = ((25 - (UBOUND(MChoice$))) \ 2)
  Col = ((80 - MaxLen) \ 2) - 1
  
  DO
    
    '--Set upper left corner of menu, turn off the cursor
    LOCATE Row, Col, 0

    ClearBack
    
    TitleBox 2, Col, MaxLen + 3, "Virginia Tax Billing Main Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
      TAXCustomerMenu
    CASE 2
      RUN "TaxPaymn"
    CASE 3
      RUN "TaxREBil"
    CASE 4
      RUN "TaxPPBil"
    CASE 5
      RUN "TaxReprt"
    CASE 6
      RUN "TAXAJUST"
    CASE 7
     RUN "TaxMRBIL"
    CASE 8
     RUN "TaxMPBIL"
    CASE 9
     RUN "TaxSetup"
    CASE 10
     RUN "TaxDMV"

    CASE 11
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

FUNCTION CheckInterestInfo
  IF Exist("TAXINTCK.DAT") THEN
    TickFile = FREEFILE
    OPEN "TAXINTCK.DAT" FOR INPUT AS #TickFile
    LINE INPUT #TickFile, LastMonth$
    CLOSE TickFile
    ThisMonth$ = LEFT$(DATE$, 2)
    IF ThisMonth$ <> LastMonth$ THEN
      ClearBack
      Ok = MsgBox("TAX", "TAXINTCK")
      SELECT CASE Ok
      CASE 1   'no do it later
        CheckInterestInfo = False
      CASE 2   'yep set flag to load interest program
        CheckInterestInfo = True
        FromFile = FREEFILE
        OPEN "TAXFROMM.DAT" FOR OUTPUT AS #FromFile
        PRINT #FromFile, "FROM MAIN"
        CLOSE FromFile
      END SELECT
    END IF
  ELSE
    UpDateTicklerFile
    CheckInterestInfo = False
  END IF

END FUNCTION

