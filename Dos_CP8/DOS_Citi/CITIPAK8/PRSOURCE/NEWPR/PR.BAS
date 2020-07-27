DEFINT A-Z

'Payroll v3.??

DECLARE SUB AlphaSupRetReport ()
DECLARE SUB PCPrintManRegister ()
DECLARE SUB RetirementRpt ()
DECLARE SUB SCRetirementRpt ()
DECLARE SUB AccruLeave (SaveScrnFlag%)
DECLARE SUB YTDFundReport ()
DECLARE SUB PRDraftMenu ()
DECLARE SUB EmpEarnHistRpt ()
DECLARE SUB ESCQuartlyWageRpt ()
DECLARE SUB SupRetReport ()
DECLARE SUB PCPrintPayRegisterS ()
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
  '$INCLUDE: 'setcnf.bi'

   CONST False = 0
   CONST True = NOT False
  
   STACK 11000
  
   CrLf$ = CHR$(13) + CHR$(10)
  
   FF$ = CHR$(12)
  
   IF NOT InitializeSystem THEN END
  
   COLOR 7, 1

   IF Exist(PayrollStatusFile) THEN
     Choice = 2
     KillFile PayrollStatusFile
     GOTO FromChkEntry
   END IF
  
   DO
     
      CursorOff

      Choice = 0

      PayMenu MainMenu, Choice, 6
     
FromChkEntry:
      SELECT CASE Choice
      CASE 1
'        AccruLeave SaveScrnFlag
'        PREmpMenu              'unrem
      CASE 2
'PRDraftMenu
'        PRCalcMenu             'unrem
'        PCPrintManRegister
'       PCPrintPayRegisterS
'        RUN "prcheck"
      CASE 3
'        YTDFundReport
'       AlphaSupRetReport
'        SupRetReport
'        RetirementRpt
         EmpEarnHistRpt
'        PRReportMenu           'unrem
'        SupRetReport
'        SCRetirementRpt
'        ESCQuartlyWageRpt
'         YTDFundReport
      CASE 4
'        PRCtrlMenu             'unrem
      CASE 5
         RUN "pr-w2"
      CASE 6
'        RUN "401K2DSK"
'        RUN "ESC2DISK"
      CASE 7
'        RUN "401K2DSK"
      CASE 8
'        RUN "RET2DISK"
      END SELECT
     
   LOOP UNTIL Choice = EscKey
  
   HideCursor   'Turn the moose cursor Off
  
PRExit:

   KillFile PRActiveFile
  
   IF Exist("CPMENU.EXE") THEN
     'RUN "CPMENU"
   ELSE
     ClearScrn
   END IF

   END

