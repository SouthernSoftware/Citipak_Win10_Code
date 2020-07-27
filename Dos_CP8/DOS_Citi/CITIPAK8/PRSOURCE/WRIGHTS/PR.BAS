DECLARE SUB CheckIssuedRpt ()
DEFINT A-Z
DECLARE SUB BenefitReport ()
DECLARE SUB EMPrintEmpList ()
DECLARE SUB AccruLeave (SaveScrnFlag%)
DECLARE SUB SupRetReport ()
DECLARE SUB PCPrintManRegister ()
DECLARE SUB RetirementRpt ()
DECLARE SUB EmpEarnHistRpt ()
DECLARE SUB PRDraftMenu ()
DECLARE SUB PCPrintPayRegister ()
  
  'Payroll v3.??
  
DECLARE FUNCTION NovellThere% ()
DECLARE FUNCTION ShareThere%
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION WEnvTest% ()

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
  '$INCLUDE: 'prif.bi'

   CONST False = 0
   CONST True = NOT False
  
   STACK 8000
  
   CrLf$ = CHR$(13) + CHR$(10)
   FF$ = CHR$(12)

   'REDIM test(1) AS GLIFDataType14
   'testlen = LEN(test(1))

   IF NOT InitializeSystem THEN END

   IF Exist("PRFROM.CHK") THEN
     KILL "PRFROM.CHK"
     CalcFlag = True
   END IF

   COLOR 7, 1
  
   DO

      IF CalcFlag THEN
        CalcFlag = False
        Choice = 2
        GOTO FromChkEntry
      END IF

      CursorOff
      Choice = 0
      PayMenu MainMenu, Choice, 8

FromChkEntry:
      SELECT CASE Choice
      CASE 1
'        PREmpMenu              'unrem
'        EMPrintEmpList
      CASE 2
'         PRCalcMenu
'         PRDraftMenu
'         PCPrintPayRegister
      CASE 3
'EmpEarnHistRpt
'        CheckIssuedRpt
'         PRReportMenu           'unrem
         SupRetReport
'         BenefitReport
'         RetirementRpt
      CASE 4
'        PRCtrlMenu             'unrem
      CASE 5
'        PRW2Menu               'unrem
      CASE 6
        RUN "ESC2DISK"
      CASE 7
        RUN "401K2DSK"
      CASE 8
        RUN "RET2DISK"
      END SELECT
     
   LOOP UNTIL Choice = EscKey
  
   HideCursor   'Turn the moose cursor Off
  
PRExit:

   KillFile PRActiveFile

   IF WEnvTest THEN
     GOTO EnvExit
   END IF

   IF Exist("CPMENU.EXE") THEN
     RUN "CPMENU"
   ELSE
     ClearScrn
   END IF

EnvExit:
   ClearScrn
   END

