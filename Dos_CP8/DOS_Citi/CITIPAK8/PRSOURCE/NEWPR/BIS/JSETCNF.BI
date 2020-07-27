'DIM Cnf AS Config

'CALL InitMouse(Cnf.Mouse)
'
'Cnf.MonTyp = Monitor
'
'ClrDisp = Cnf.MonTyp = 3 OR (Cnf.MonTyp >= 5 AND Cnf.MonTyp <= 11)
'IF INSTR(COMMAND$, "/B") THEN ClrDisp = 0
'
'IF ClrDisp THEN    'ncp  'pr
'   Cnf.PulBar = 3      '23
'   Cnf.MenBox = 10     '27
'   Cnf.ActivCh = 11   ' 31
'   Cnf.InActCh = 3     '31
'   Cnf.HiLite = 112    '112
'   Cnf.InActHiLt = 64 '112
'   Cnf.NonMen = 7    '30
'   Cnf.CurSize = 12     '7
'
' ELSE
'   Cnf.PulBar = 3
'   Cnf.MenBox = 10
'   Cnf.ActivCh = 11
'   Cnf.InActCh = 3
'   Cnf.HiLite = 112
'   Cnf.InActHiLt = 64
'   Cnf.NonMen = 7
'   Cnf.CurSize = 12
'END IF

'DIM SHARED Cnf AS Config

CALL InitMouse(Cnf.Mouse)

Cnf.MonTyp = Monitor

'CALL NoReTrace

ClrDisp = Cnf.MonTyp = 3 OR (Cnf.MonTyp >= 5 AND Cnf.MonTyp <= 11)
IF INSTR(COMMAND$, "/B") THEN ClrDisp = 0

IF ClrDisp THEN
   MonoCode = 0
'   Cnf.PulBar = 48
'   Cnf.MenBox = 49
'   Cnf.ActivCh = 48
'   Cnf.InActCh = 52
'   Cnf.HiLite = 31
'   Cnf.InActHiLt = 64
'   Cnf.NonMen = 30
'   Cnf.CsrSize = 7

   Cnf.PulBar = 3      '23
   Cnf.MenBox = 10     '27
   Cnf.ActivCh = 11   ' 31
   Cnf.InActCh = 3     '31
   Cnf.HiLite = 112    '112
   Cnf.InActHiLt = 64 '112
   Cnf.NonMen = 7    '30
   Cnf.CsrSize = 12


ELSE
   MonoCode = 3
'   Cnf.PulBar = 112
'   Cnf.MenBox = 112
'   Cnf.ActivCh = 112
'   Cnf.InActCh = 112
'   Cnf.HiLite = 15
'   Cnf.InActHiLt = 80
'   Cnf.NonMen = 7
'   Cnf.CsrSize = 12

   Cnf.PulBar = 3
   Cnf.MenBox = 10
   Cnf.ActivCh = 11
   Cnf.InActCh = 3
   Cnf.HiLite = 112
   Cnf.InActHiLt = 64
   Cnf.NonMen = 7
   Cnf.CsrSize = 12
   
END IF

