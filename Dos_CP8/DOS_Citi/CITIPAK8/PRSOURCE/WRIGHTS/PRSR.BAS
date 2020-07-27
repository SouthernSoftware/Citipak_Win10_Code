DECLARE SUB PRCheckMenu ()
DECLARE SUB PCPrintPayRegisterS ()
   DEFINT A-Z
  
  'Payroll v3.??
  
DECLARE FUNCTION NovellThere% ()
DECLARE FUNCTION ShareThere%
DECLARE FUNCTION Exist% (FileName$)
  
DECLARE SUB BlockClear ()
DECLARE SUB ClearScrn ()
DECLARE SUB CursorOff ()
DECLARE SUB DisplayMiscScrn (ScrnName$)
DECLARE SUB KillFile (FileName$)
DECLARE SUB HideCursor ()
DECLARE FUNCTION InitializeSystem% ()
  
DECLARE SUB PREmpMenu ()
DECLARE SUB PRCalcMenu ()
DECLARE SUB PRReportMenu ()
DECLARE SUB PRCtrlMenu ()
DECLARE SUB PRW2Menu ()
DECLARE SUB PayMenu (BYVAL MenuNum, Choice, NumOfItems%)
  
DECLARE SUB PRInitFiles ()
DECLARE SUB QPrintRC (Text$, Row, Col, FrameColor)
  
  '$INCLUDE: 'DefCnf.bi'      'Defines a TYPE for monitor and color info..
  '$INCLUDE: 'FieldInf.bi'    '        "          field information
  '$INCLUDE: 'FormEdit.bi'    '        "          form editing information
  '$INCLUDE: 'PRFiles.bi'
  
   CONST False = 0
   CONST True = NOT False
  
   STACK 8000
  
   CrLf$ = CHR$(13) + CHR$(10)
  
   FF$ = CHR$(12)
  
   IF NOT InitializeSystem THEN END
  
   COLOR 7, 1
  
   DO
     
      CursorOff
     
      Choice = 0
     
      PayMenu MainMenu, Choice, 5
     
      SELECT CASE Choice
      CASE 1
'         PREmpMenu              'unrem
      CASE 2
         PCPrintPayRegisterS
'         PRCalcMenu             'unrem
'         PRCheckMenu
      CASE 3
'         PRReportMenu           'unrem
       CASE 4
'         PRCtrlMenu             'unrem
      CASE 5
'         PRW2Menu               'unrem
      END SELECT
     
   LOOP UNTIL Choice = EscKey
  
   HideCursor   'Turn the moose cursor Off
  
PRExit:

   KillFile PRActiveFile
  
   IF Exist("CPMENU.EXE") THEN
     RUN "CPMENU"
   ELSE
     ClearScrn
   END IF

   END

