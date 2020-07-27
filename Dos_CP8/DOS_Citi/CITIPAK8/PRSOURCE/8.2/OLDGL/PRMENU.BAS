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
'$INCLUDE: 'PREmpRec.bi'
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

  CASE EmpMaintMenu       'EmployeeMaintance
    PMenu$(0) = "Employee Maintenance Menu"
    PMenu$(1) = " Add A New Employee Record "
    PMenu$(2) = " Edit / View Employee Record "
    PMenu$(3) = " Print Employee Data File "
    PMenu$(4) = " Print Employee List "
    PMenu$(5) = " Print Terminated Employee List "

'    PMenu$(3) = " Edit / View Employee Note File "
'    PMenu$(4) = " Terminate an Employee "
'    PMenu$(5) = " Reactivate Terminated Employee "

  CASE PayProcessMenu  'PayrollProcessing
    PMenu$(0) = "Payroll Processing Menu"
    PMenu$(1) = " Accrue Leave Benefits "
    PMenu$(2) = " Set Pay Period Defaults "
    PMenu$(3) = " Enter / Edit Payroll Transactions "
    PMenu$(4) = " Print Register "
    PMenu$(5) = " Print Payroll Checks "
    PMenu$(6) = " Post Payroll Transactions "
    PMenu$(7) = " Manual Transaction Entry "
    PMenu$(8) = " Void a Posted Transaction "
    PMenu$(9) = " Year End Operations "

  CASE ManMenu        'ManualTransactions
    PMenu$(0) = "Manual Transaction Menu"
    PMenu$(1) = " Enter Manual Transactions "
    PMenu$(2) = " Print Register "
    PMenu$(3) = " Post Manual Transactions "

  CASE YTDMenu           'YearEndOperations
    PMenu$(0) = "W-2 Processing Menu"

    PMenu$(1) = " Extract W-2 Information "
    PMenu$(2) = " Edit/Review Employee's W-2 "
    PMenu$(3) = " Print W-2 Report "
    PMenu$(4) = " Print W-2 Forms "

'    PMenu$(5) = " Electronic Filing "
'    PMenu$(4) = " Remove Terminated Employee Records "
   
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
   PMenu$(14) = " SEPP Contribution Report "

'    PMenu$(6) = " Employer Matching Exp by Emp "
'    PMenu$(7) = " Employer Matching Exp by Dept "
'    PMenu$(13) = " User Defined Query "
'    PMenu$(2) = " Employee Labor Analysis "
'    PMenu$(3) = " Departmental Labor Analysis "
'    PMenu$(3) = " Payroll Deduction Schedule "

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
    PMenu$(10) = " Printer Setup & Configuration "
    PMenu$(11) = " INITIALIZE NEW YEAR PROCESSING! "

    'PMenu$(10) = " Backup/Restore Payroll Files "
    'PMenu$(11) = " System Operator's Menu "
'    PMenu$(10) = " Matching Exp Code Maintenance "

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

