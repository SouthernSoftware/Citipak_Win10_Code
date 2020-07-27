DEFINT A-Z
DECLARE SUB WInput (Edit$, GoodKey$, Row%, Col%, ExitCode%)
DECLARE SUB BackLightToggle ()
DECLARE FUNCTION BiosKey% ()
DECLARE FUNCTION QPTrim$ (Text$)

SUB WInput (Edit$, GoodKey$, Row, Col, ExitCode)
'Return Codes:
' Enter Key Exit
'     0
' Escape Key Exit
'   -27 Escape Key was Pressed to exit sub
'       Edit string reset to original string
'

  STATIC InsertMode

  IF ExitCode = -1 THEN
    InsertMode = -1
  ELSE
    InsertMode = 0
  END IF

  GOSUB ShowCursor

  Original$ = Edit$
  ExitCode = 0

  MaxLength = LEN(Edit$)        'max length of input string
  Ptr = LEN(RTRIM$(Edit$))
  IF Ptr = MaxLength THEN Ptr = MaxLength - 1
  
  DO                            'main processing loop
    
    LOCATE Row, Col, 1           'update cursor position.
    PRINT Edit$;                 'and display the line
    LOCATE Row, Col + Ptr      'update cursor position.
    
    PressedKey = BiosKey        'wait for a key press

    SELECT CASE PressedKey      'process the key

'****  Exit keys
    CASE 13   'ENTER - Accept line and, exit
      Edit$ = QPTrim$(Edit$)
      IF LEN(Edit$) = 0 THEN
        ExitCode = -1
      ELSE
        ExitCode = 13
      END IF

      EXIT DO

    CASE 27   'ESCAPE - Abort operation exit
      ExitCode = -27
      Edit$ = Original$
      EXIT DO

    CASE -61             'F3
      BackLightToggle
    
    CASE -59, -60, -62, -72, -80   'F1 F2 F4 DnArrow UpArrow
      'UpArrow DnArrow F1 F2 F4
      ExitCode = ABS(PressedKey)
      EXIT DO

'add alt exits here
' i.e. downarrow
'End of Exit keys

'****  Cursor keys
    CASE -82   'INSERT - Toggle insert or, overtype mode
      InsertMode = NOT InsertMode
      GOSUB ShowCursor
    CASE -75   'LEFTARROW - Moves cursor left
      IF Ptr THEN Ptr = Ptr - 1
      IF Col + Ptr <= 0 THEN
        Ptr = Ptr + 1
      END IF
    CASE -77   'RIGHTARROW - Moves cursor right
      IF Ptr < MaxLength - 1 THEN Ptr = Ptr + 1
    CASE -71   'HOME - Moves cursor to first position
      Ptr = 0
    CASE -79   'END - Moves cursor to last position
      Ptr = LEN(RTRIM$(Edit$))
      IF Ptr = MaxLength THEN Ptr = MaxLength - 1
'End of Cursor keys

'****   Editing Keys & Normal Input
    CASE -117   'CONTROLEND - Erase from cursor to end of Edit$
      IF Ptr THEN
        LSET Edit$ = LEFT$(Edit$, Ptr)
      ELSE
        Edit$ = SPACE$(LEN(Edit$))
      END IF
    CASE -22    'ALT-U     - Restore to orginial string
      Edit$ = Original$
      Ptr = 0

    CASE 8      'BACKSPACE - Moves cursor left and erase characters
      IF Ptr THEN
        Edit$ = LEFT$(Edit$, Ptr - 1) + MID$(Edit$, Ptr + 1) + " "
        Ptr = Ptr - 1
      END IF

    CASE -83    'DELETE - Erases characters at cursor
      Edit$ = LEFT$(Edit$, Ptr) + MID$(Edit$, Ptr + 2) + " "
   
    CASE ELSE   'ALL OTHER NORMAL KEYS
      IF PressedKey > 1 AND PressedKey < 256 THEN
        K$ = CHR$(PressedKey)
        K$ = UCASE$(K$)
        IF INSTR(GoodKey$, K$) THEN       'if key is in the keylist
          IF InsertMode THEN              'Insert or Overtype the key
            Edit$ = LEFT$(Edit$, Ptr) + K$ + MID$(Edit$, Ptr + 1)
            Edit$ = LEFT$(Edit$, MaxLength)
          ELSE
            IF Ptr < MaxLength THEN MID$(Edit$, Ptr + 1, 1) = K$
          END IF
          IF Ptr < MaxLength - 1 THEN Ptr = Ptr + 1  'Are we at max length
        END IF
      END IF
'End of Editing & Input keys
    END SELECT
  LOOP
EXIT SUB

ShowCursor:

  IF InsertMode THEN
    LOCATE , , , 0, 6
  ELSE      'show appropriate cursor type
    LOCATE , , , 5, 6
  END IF
RETURN

END SUB

