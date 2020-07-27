'DECLARE SUB WColorSave ()
'DECLARE SUB WColorRest ()

DECLARE SUB WInitVideo ()
DECLARE SUB WResetVideo ()

DECLARE SUB WCls ()
DECLARE SUB WColor (BYVAL ForeGround%, BYVAL BackGround%)
DECLARE SUB WColorGet (ForeGround%, BackGround%)

DECLARE SUB WPrintRC (Text$, BYVAL Column%, BYVAL Row%)
'DECLARE SUB WPrintRC (BYVAL Column%, BYVAL Row%, Text$)
DECLARE SUB WPrint (Text$)
DECLARE SUB WPaintArea (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL Attr%)

DECLARE SUB WCursorOn ()
DECLARE SUB WCursorOff ()

DECLARE SUB WLocate (BYVAL Row%, BYVAL Column%)
DECLARE SUB WCursorGetPos (Row%, Column%)

DECLARE SUB WCursorGetSize (TopLine%, BotLine%)
DECLARE SUB WCursorSetSize (BYVAL TopLine%, BYVAL BotLine%)

DECLARE SUB WazzWindInit ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameType%, BYVAL Shadow%)

DECLARE SUB Blink (BYVAL BlinkFlag%)

'untested VVVVV
'Declare Sub WPRestore (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, SEG(Array%(X%))

