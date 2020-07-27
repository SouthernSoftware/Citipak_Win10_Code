DEFINT A-Z
DECLARE SUB SCRetirementRpt ()
DECLARE SUB YTDFundReport ()
DECLARE SUB EMPrintEmpData ()
DECLARE SUB CheckIssuedRpt ()
DECLARE SUB VARetirementRpt ()
DECLARE SUB Post2GL (FileName$, PSysRec() AS ANY, BadTrans%)
DECLARE SUB ManualMenu ()
DECLARE SUB EmpEarnHistRpt ()
DECLARE SUB EMPrintEmpList ()
DECLARE SUB SEPPContRpt ()
DECLARE SUB ESCQuartlyWageRpt ()
DECLARE SUB AccruLeave (SaveScrnFlag%)
DECLARE SUB BenefitReport ()
DECLARE SUB RetirementRpt ()
DECLARE SUB SupRetReport ()
DECLARE SUB GrossWageReport ()
DECLARE SUB PCPrintManRegister ()
DECLARE SUB PRCheckMenu ()
DECLARE SUB PCPrintPayRegisterS ()
  
'Payroll v3.??
  
DECLARE FUNCTION NovellThere% ()
DECLARE FUNCTION ShareThere%
DECLARE FUNCTION Exist% (FileName$)
DECLARE FUNCTION FUsing$ (Number$, Image$)

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
  
'   STACK 5000
  
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
       PREmpMenu              'unrem
'        EMPrintEmpList
'        EMPrintEmpData
      CASE 2
'        PCPrintPayRegister
'        PCPrintPayRegisterS
'        PCPrintManRegister
'        PRCalcMenu             'unrem
'        PRCheckMenu
      CASE 3
'        EmpEarnHistRpt
'        SupRetReport
'        BenefitReport
'        CheckIssuedRpt
'        ESCQuartlyWageRpt
'        SCRetirementRpt
'        RetirementRpt
'         GrossWageReport
       CASE 4
'        PRCtrlMenu             'unrem
      CASE 5
'         PRW2Menu               'unrem
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

'   END

'   Post2GL "TEMPIF.DAT", PSysRec(), BadAccts   'unrem

