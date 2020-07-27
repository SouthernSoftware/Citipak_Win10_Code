DEFINT A-Z

DECLARE FUNCTION ASCII% (Strng$)

DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()
DECLARE SUB HideCursor ()
DECLARE SUB PayMenu (BYVAL MenuNum%, Choice%, NumOfItems%)
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
DECLARE SUB ShowCursor ()
DECLARE SUB VertMenu (Item$(), Choice%, MaxLen%, BoxBot%, Ky$, Action%, Cnf AS ANY)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)

'$INCLUDE: 'DefCnf.bi'                  'Defines TYPE for monitor/color info.

'$INCLUDE: 'FieldInf.bi'                '        "        field information
'$INCLUDE: 'FormEdit.bi'                '        "        form editing info.
'$INCLUDE: 'PRFiles.bi'
'$INCLUDE: 'PREmpRec.biw'
'$INCLUDE: 'PageInfo.bi'    '        "          Page information
'$INCLUDE: 'Retire.Bi'

CONST False = 0
CONST True = NOT False

REDIM SHARED Menu$(0)

SUB PayMenu (BYVAL MenuNum, Choice, NumOfItems)

  REDIM PMenu$(0 TO 20)

  SELECT CASE MenuNum
  CASE MainMenu             'MainMenu
    PMenu$(0) = "Main Menu"
    PMenu$(1) = " Employee File Maintenance "
    PMenu$(2) = " Payroll Processing "
    PMenu$(3) = " Reports Processing "
    PMenu$(4) = " Control File Maintenance "
    PMenu$(5) = " W-2 Processing "
    PMenu$(6) = " NC ESC Mag Media Report "
    PMenu$(7) = " 401K Mag Media Report "
    PMenu$(8) = " Retirement Mag Media Report "


  CASE EmpMaintMenu       'EmployeeMaintance
    PMenu$(0) = "Employee Maintenance Menu"
    PMenu$(1) = " Add A New Employee Record "
    PMenu$(2) = " Edit / View Employee Record "
    PMenu$(3) = " Print Employee Data File "
    PMenu$(4) = " Print Employee List "
    PMenu$(5) = " Print Terminated Employee List "

  CASE PayProcessMenu  'PayrollProcessing
    PMenu$(0) = "Payroll Processing Menu"
    PMenu$(1) = " Accrue Leave Benefits "
    PMenu$(2) = " Set Pay Period Defaults "
    PMenu$(3) = " Enter / Edit Payroll Transactions "
    PMenu$(4) = " Print Register "
    PMenu$(5) = " Print Payroll Checks "
    PMenu$(6) = " ACH Bank Draft Processing "
    PMenu$(7) = " Post Payroll Transactions "
    PMenu$(8) = " Manual Transaction Entry "

    PMenu$(9) = " Void a Posted Transaction "
    PMenu$(10) = " Year End Operations "

  CASE ManMenu        'ManualTransactions
    PMenu$(0) = "Manual Transaction Menu"
    PMenu$(1) = " Enter Manual Transactions "
    PMenu$(2) = " Print Register "
    PMenu$(3) = " Post Manual Transactions "

  CASE YTDMenu           'YearEndOperations
    PMenu$(0) = "W-2 Processing Menu"
    PMenu$(1) = " Extract 941 W-2 Information "
    PMenu$(2) = " Extract Med Only W-2 Information "
    PMenu$(3) = " Edit/Review Employee's W-2 "
    PMenu$(4) = " Print W-2 Report "
    PMenu$(5) = " Print W-2 Forms "

  CASE ReportsMenu          'ReportsMenu
    PMenu$(0) = "Reports Menu"
    PMenu$(1) = " Print Employee Data File "
    PMenu$(2) = " Active Employee List "
    PMenu$(3) = " Terminated Employee List "
    PMenu$(4) = " Employee Earnings History "
    PMenu$(5) = " Gross Wage Report "
    PMenu$(6) = " Payroll Deductions Taken "
    PMenu$(7) = " ESC Report "
    PMenu$(8) = " Leave Benefit Report "
    PMenu$(9) = " YTD Wage Distributions "
   PMenu$(10) = " Checks Issued Report "
   PMenu$(11) = " Reprint Last/Prior Reports "
   PMenu$(12) = " Retirement Reports "
   PMenu$(13) = " Supplemental Retirement Report "
   PMenu$(14) = " LIFE & HOSPITAL Report "

  CASE ControlMenu         'ControlFileMaintenance
     PMenu$(0) = "Control Maintenance Menu"
     PMenu$(1) = " Employer File Maintenance "
     PMenu$(2) = " System File Maintenance "
     PMenu$(3) = " State Tax Table Maintenance "
     PMenu$(4) = " Federal Tax Table Maintenance "
     PMenu$(5) = " E I C Table Maintenance "
     PMenu$(6) = " Leave Benefit Table Maintenance "
     PMenu$(7) = " Deduction Code Maintenance "
     PMenu$(8) = " Earnings Code Maintenance "
     PMenu$(9) = " Retirement File Maintenance "
    PMenu$(10) = " ACH Bank Draft Information"
    PMenu$(11) = " Printer Setup & Configuration "
    PMenu$(12) = " INITIALIZE NEW YEAR PROCESSING! "

  CASE 8                 'SystemOperatorsMenu
    PMenu$(0) = "System Operator Menu"
    PMenu$(1) = " Password Maintenance "
    PMenu$(2) = " Activity Log "
    PMenu$(3) = " Check Printing Format "

  CASE 9               'CheckPrintMenu
    PMenu$(0) = "Check Printing Menu"
    PMenu$(1) = " Print Payroll Checks "
    PMenu$(2) = " Reprint Selected Checks "
    PMenu$(3) = " Print Check Register "

  CASE DraftMenu
    PMenu$(0) = "ACH Bank Draft Menu"
    PMenu$(1) = "Employees To Draft Report"
    PMenu$(2) = "Prepare Draft Transmission File"
    PMenu$(3) = "Prepare Draft Prenote File"
    PMenu$(4) = "Print Draft Employee Listing"

  'CASE 9               'CheckPrintMenu
  '  PMenu$(0) = " Check/Voucher Printing Menu "
  '  PMenu$(1) = " Print Payroll Checks "
  '  PMenu$(2) = " Print ACH Vouchers "
  '  PMenu$(3) = " Reprint Selected Checks"
  '  PMenu$(4) = " Reprint Selected Vouchers"
  '  PMenu$(5) = " Print Check/Vouchers Register"
  CASE ELSE
  END SELECT

  GOSUB SetMenu
  
  VAction = 1
  BoxBot = 17
  BoxTop = 9

  LeftCol = ((80 - MaxLen) \ 2) - 1
  TitleCol = ((80 - LEN(MenuTitle$)) \ 2) + 1

  CursorOff
  HideCursor
  LOCATE BoxTop, LeftCol
  BlockClear
  WazzWind 3, LeftCol, 5, LeftCol + MaxLen + 3, Cnf.MenBox, 1, True
  WazzWind 21, LeftCol, 23, LeftCol + MaxLen + 3, Cnf.MenBox, 1, True
  QPrintRC MenuTitle$, 4, TitleCol, Cnf.ActivCh
  QPrintRC " Use (" + CHR$(24) + "-" + CHR$(25) + ") to select", 22, LeftCol + (MaxLen \ 2) - 8, Cnf.ActivCh
  ShowCursor

  DO

'    call QPrintRC(STR$(FRE("")), 25, 1, 31)

    VertMenu Menu$(), Choice, MaxLen, BoxBot, Ky$, VAction, Cnf
    SELECT CASE ASCII(Ky$)
      CASE EscKey
        Choice = EscKey
        ExitFlag = True
      CASE 13
        ExitFlag = True
      CASE ELSE
    END SELECT
  LOOP UNTIL ExitFlag
'  LOCATE , , 0
  CursorOff

  ERASE Menu$
EXIT SUB

SetMenu:

  REDIM Menu$(NumOfItems)
  FOR Cnt = 1 TO NumOfItems
    Menu$(Cnt) = PMenu$(Cnt)
    ItemLen = LEN(Menu$(Cnt))
    IF MaxLen < ItemLen THEN
      MaxLen = ItemLen
    END IF
  NEXT
  MenuTitle$ = PMenu$(0)
  ERASE PMenu$
  RETURN

END SUB

