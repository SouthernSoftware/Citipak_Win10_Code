DEFINT A-Z
DECLARE SUB QPrintRC (t$, r%, c%, clr%)
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
'$INCLUDE: 'DefCnf.BI'

SUB TitleBox (Row, LeftCol, BoxWidth, Title$, Cnf AS Config)

 TitleCol = (LeftCol + ((BoxWidth / 2) - (LEN(Title$) / 2)))
 IF NOT TitleCol MOD 2 THEN
   TitleCol = TitleCol + 1
 END IF

 WazzWind Row, LeftCol, Row + 2, LeftCol + BoxWidth, Cnf.MenBox, 1, 1

 QPrintRC Title$, Row + 1, TitleCol, Cnf.ActivCh

END SUB

