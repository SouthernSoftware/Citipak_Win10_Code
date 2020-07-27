'
'    1 = Mono
'    2 = Hercules mono
'    3 = CGA
'    4 = EGA (mono monitor)
'    5 = EGA (color monitor)
'    6 = VGA (mono monitor)
'    7 = VGA (color monitor)
'    8 = MCGA (mono monitor)
'    9 = MCGA (color monitor)
'   10 = EGA  (CGA monitor)
'   11 = IBM 8514/A
'
DECLARE FUNCTION WInitVideo (BYVAL FarStringFlag%)
'DECLARE SUB WInitVideo (BYVAL FarStringFlag%)
'WInitVideo CAN also be declared as a subroutine. (no monitor code returned)
DECLARE SUB WResetVideo ()
DECLARE SUB WCursorOn ()
DECLARE SUB WCursorOff ()
DECLARE SUB WLocate (BYVAL Row%, BYVAL Col%)
DECLARE SUB WCursorGetSize (TopScanLine%, BotScanLine%)
DECLARE SUB WCursorSetSize (BYVAL TopScanLine%, BYVAL BotScanLine%)
DECLARE SUB WCursorGetPos (Row%, Col%)
DECLARE SUB WCls ()
DECLARE SUB WColorSet (BYVAL ForeGround%, BYVAL BackGround%)
DECLARE SUB WColorGet (ForeGround%, BackGround%)
DECLARE SUB WPrint (AnyString$)
DECLARE SUB WScrollD ()
DECLARE SUB WScrollU ()
DECLARE SUB WScrollAreaD (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%)
DECLARE SUB WScrollAreaU (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%)
DECLARE SUB WScrnSave (SEG ScrnArray AS ANY)
DECLARE SUB WScrnRest (SEG ScrnArray AS ANY)
DECLARE SUB WPartRest (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, SEG Array AS ANY)
DECLARE SUB WPaintArea (BYVAL TRow%, BYVAL LCol%, BYVAL BRow%, BYVAL RCol%, BYVAL Attr%)

