DEFINT A-Z
DECLARE FUNCTION RatePassWordOK% ()
DECLARE SUB MakeBKGround ()
DECLARE SUB WazzWind (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
DECLARE FUNCTION nGetPassword$ (Row%, Col%, Max%, Abort%)
DECLARE SUB UBLog (Text$)
DECLARE SUB QPrintRC (Text$, Row%, Col%, Kolor)
DECLARE SUB WaitForAction ()
DECLARE SUB BlockClear ()
DECLARE SUB CursorOff ()

FUNCTION nGetPassword$ (Row%, Col%, Max%, Abort%)
  
Start:
  
  CurPos = Col
  
  DO
    
    LOCATE Row, CurPos, 1
    P$ = INPUT$(1)
    
    SELECT CASE P$
      
    CASE CHR$(8)                'Backspace

      IF CurPos > Col THEN
        CurPos = CurPos - 1
        IF CurPos = Col THEN     'Check for backspace key at
          PW$ = ""
        END IF
        LOCATE Row, CurPos ': PRINT " "             'Clear the last *
        QPrintRC " ", Row, CurPos, 14
      END IF

    CASE CHR$(13)               'Enter
      EXIT DO
      
    CASE CHR$(27)               'Esc
      Abort = True
      EXIT DO
      
    CASE CHR$(32) TO CHR$(126)  'ignore all other keys
      LOCATE Row, CurPos        ' except keyboard chars
      QPrintRC "*", Row, CurPos, 14
      PW$ = PW$ + P$
      CurPos = CurPos + 1
      
    END SELECT
    
  LOOP UNTIL LEN(PW$) = Max
  
  nGetPassword$ = UCASE$(PW$)

  CursorOff
END FUNCTION

