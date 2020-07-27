DEFINT A-Z
DECLARE SUB BlockClear ()
DECLARE SUB WaitForAction ()
DECLARE SUB CursorOff ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE SUB QPrintRC (Text$, Row, Col, Kolor)

CONST False = 0, True = NOT False



  '$INCLUDE: 'GL.BI'

SUB SendDist2GL

   CursorOff
   REDIM GJEditRec(1) AS TrEditRecType
   GJEdLen = LEN(GJEditRec(1))

   GJEditFileNum = FREEFILE
   OPEN "GLGJED.DAT" FOR RANDOM ACCESS READ WRITE SHARED AS GJEditFileNum LEN = GJEdLen
   NumEdTrans = LOF(GJEditFileNum) \ GJEdLen
   CLOSE
   IF NumEdTrans > 0 THEN
     BlockClear
     WazzWind 7, 20, 18, 60, 78, 1, True
     QPrintRC "[", 7, 22, -1
     QPrintRC "]", 7, 30, -1
     QPrintRC " ERROR ", 7, 23, 79
     QPrintRC "General Journal Edit file must be", 9, 22, 79
     QPrintRC "EMPTY before you can transfer Utility", 10, 22, 79
     QPrintRC "Billing Information.", 11, 22, 79
     QPrintRC "Ý        Þ", 16, 35, 127
     QPrintRC "Þ", 16, 44, 120
     QPrintRC "OK", 16, 39, 112
     WaitForAction
     GOTO SendExit
   END IF


SendExit:

END SUB

