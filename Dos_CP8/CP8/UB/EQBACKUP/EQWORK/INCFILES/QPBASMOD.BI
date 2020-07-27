DEFINT A-Z

'Subs in other modules
DECLARE SUB TextIn (t$, Max, NumOnly, CapsOn, ExitCode, Colr)
DECLARE SUB DateIn (D$, ExitCode, Colr)
DECLARE SUB VertMenu (Item$(), Choice, MaxLen, BoxBot, ky$, Action, Cnf AS ANY)
DECLARE SUB YesNoB (YN$, ExitCode, Colr)
DECLARE SUB LibFile2Scrn (LibName$, ScrnName$, MonoCode%, Attribute%, ErrorCode%)
DECLARE FUNCTION InputBox$ (LibName$, ScrnName$, Row, Col, Max, NumOnly)
