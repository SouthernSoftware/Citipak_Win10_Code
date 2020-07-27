DEFINT A-Z

DECLARE SUB PRInitFiles ()
DECLARE SUB FCreate (FileName$)
DECLARE FUNCTION Exist% (FileName$)
DECLARE SUB InitializeSystem ()
DECLARE SUB PRInitFiles ()
DECLARE FUNCTION Monitor ()
DECLARE SUB HideCursor ()
DECLARE SUB TextCursor (Fg, Bg)
DECLARE SUB ShowCursor ()
DECLARE SUB InitMouse (MooseFlag%)

'$INCLUDE: 'DefCnf.bi'      'Defines a TYPE for monitor and color info..

'$INCLUDE: 'PRFILES.BI'

'$INCLUDE: 'FormEdit.BI'

SUB InitializeSystem

InitMouse Cnf.Mouse

Cnf.MonTyp = Monitor

ClrDisp = Cnf.MonTyp = 3 OR (Cnf.MonTyp >= 5 AND Cnf.MonTyp <= 11)
IF INSTR(COMMAND$, "/B") THEN ClrDisp = 0

IF ClrDisp THEN
   MonoCode = 0
   Cnf.PulBar = 31
   Cnf.MenBox = 23' 49
   Cnf.ActivCh = 31 '48
   Cnf.InActCh = 52
   Cnf.HiLite = 112'31
   Cnf.InActHiLt = 64
   Cnf.NonMen = 30
   Cnf.CsrSize = 7
ELSE 'monochrome
   MonoCode = 3
   Cnf.PulBar = 112
   Cnf.MenBox = 112
   Cnf.ActivCh = 112
   Cnf.InActCh = 112
   Cnf.HiLite = 15
   Cnf.InActHiLt = 80
   Cnf.NonMen = 7
   Cnf.CsrSize = 12
END IF
  
  TextCursor 14, 6             'define the Mouse cursor
'  TextCursor -2, -2             'define the Mouse cursor

  'TextCursor 7, 0                       'define the Mouse cursor
  ShowCursor    'Turn it on
  
END SUB

