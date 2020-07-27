DEFINT A-Z
DECLARE FUNCTION CheckInterestInfo% ()
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION Monitor% ()
DECLARE FUNCTION MsgBox% (LibName$, FormName$)
DECLARE FUNCTION OK2UPDateCust% ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE FUNCTION WEnvTest% ()
DECLARE SUB ClearBack ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayUBScrn (ScrnName$)
DECLARE SUB HideCursor ()
DECLARE SUB PIProcessMenu (JrnlType%)
DECLARE SUB PrintHelp (H$)
DECLARE SUB PrintTitle (Title$)
DECLARE SUB SendDist2GL ()
DECLARE SUB ShowCursor ()
DECLARE SUB TAXCustomerMenu ()
DECLARE SUB TextCursor (MouseFg%, MouseBg%)
DECLARE SUB TitleBox (Row%, LeftCol%, BoxWidth%, Title$, Cnf AS ANY)
DECLARE SUB UBBillMenu ()
DECLARE SUB UBCustomerMenu ()
DECLARE SUB UBMiscMenu ()
DECLARE SUB UpDateTicklerFile ()
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf AS ANY)

  CONST False = 0, True = NOT False

  '$INCLUDE: 'DefCnf.BI'
  '$INCLUDE: 'formedit.BI'
  '$INCLUDE: 'fieldinf.BI'
  '$INCLUDE: 'qscr.BI'
  '$INCLUDE: 'SetCnf.BI'


  STACK 5000

  LOCATE 1, 1, 0

  IF INSTR(COMMAND$, "TEST") = 0 THEN
    IF CheckInterestInfo THEN
      RUN "CAINT"
    END IF
  END IF

  '--Dim the choice array to the number of menu items
  REDIM MChoice$(1 TO 8)
  
  MChoice$(1) = "Master Record Maintenance"
  MChoice$(2) = "Assessment Payments"
  MChoice$(3) = "Apply Interest"
  MChoice$(4) = "Assessment Reports"
  MChoice$(5) = "Print Invoices"
  MChoice$(6) = "Enter Beginning Transaction"
  MChoice$(7) = "Assessment Adjustments"
  MChoice$(8) = "Exit to OS"
  
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
    
    TitleBox 2, Col, MaxLen + 3, "Assessment Main Menu ", Cnf
    TitleBox 21, Col, MaxLen + 3, "Use " + CHR$(24) + "-" + CHR$(25) + " to select", Cnf
    
    ShowCursor
    
    VertMenu MChoice$(), Choice, MaxLen, BoxBot, Ky$, Action, Cnf
    
    IF Ky$ = CHR$(27) THEN EXIT DO              'choice = 0
    
    SELECT CASE Choice
    CASE 1
     RUN "CACUST"
    CASE 2
     RUN "CAPAY"
    CASE 3
     RUN "CAINT"
    CASE 4
     RUN "CAREPORT"
    CASE 5
     RUN "CAINV"
    CASE 6
     RUN "CABEG"
    CASE 7
     RUN "CAADJUST"
    CASE 8
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
  IF Exist("CAINTCK.DAT") THEN
    TickFile = FREEFILE
    OPEN "CAINTCK.DAT" FOR INPUT AS #TickFile
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

