'COMMON SHARED Cnf as config

Cnf.MonTyp = Monitor

ClrDisp = Cnf.MonTyp = 3 OR (Cnf.MonTyp >= 5 AND Cnf.MonTyp <= 11)

IF INSTR(COMMAND$, "/NM") THEN
  Cnf.Mouse = 0
ELSE
  CALL InitMouse(Cnf.Mouse)
  CALL TextCursor(14, -2)
END IF

IF INSTR(COMMAND$, "/B") THEN ClrDisp = 0

IF ClrDisp THEN
   MonoCode = 0
   Cnf.PulBar = 3      '23
   Cnf.MenBox = 10     '27
   Cnf.ActivCh = 11    '31
   Cnf.InActCh = 3     '31
   Cnf.HiLite = 112    '112
   Cnf.InActHiLt = 64  '112
   Cnf.NonMen = 7      '30
   Cnf.CsrSize = 12
ELSE
   MonoCode = 3
   Cnf.PulBar = 15
   Cnf.MenBox = 7
   Cnf.ActivCh = 15
   Cnf.InActCh = 3
   Cnf.HiLite = 112
   Cnf.InActHiLt = 64
   Cnf.NonMen = 7
   Cnf.CsrSize = 12
END IF

