DECLARE SUB MakeBKGround ()
DECLARE SUB WaitForKey ()
DECLARE SUB WazzWind2 (BYVAL TopRow%, BYVAL LeftCol%, BYVAL BotRow%, BYVAL RghtCol%, BYVAL FrameColor%, BYVAL FrameType%, BYVAL Shadow%)
GetPW:
  DO
'    MakeBKGround
    WazzWind2 8, 15, 16, 65, 15, 1, 1
    QPrintRC "Southern Software Systems, Inc.", 10, 22, 15
    QPrintRC "Budgetary Accounting System v8.20", 12, 22, 15
    QPrintRC "Enter the password to begin:", 14, 22, 15

    COLOR 14

    PW$ = nGetPassword(14, 51, 10, Abort)

    IF Abort THEN
      k = 27
      EXIT DO
      'GOTO GetOut
      'RUN "CPMENU"
    END IF

    IF PW$ = "GLMANG" THEN
      OkFlag = -1
      EXIT DO
    END IF

    Try = Try + 1

    IF Try = 4 THEN
      ExitFlag = -1
      'MakeBKGround
      WazzWind2 8, 15, 16, 65, 15, 1, 1
      QPrintRC "Get the password from the system operator.", 11, 20, 14
      QPrintRC "Press any key to return to the Main Menu.", 13, 20, 14
      WaitForKey
      EXIT DO
    END IF

    'LOCATE 1, 1, 0
    'QPSound 100, 3
    'QPSound 40, 3
    QPrintRC "Invalid Password. Press any key to try again.", 14, 18, 31


  LOOP

  IF k = 27 THEN
    GOTO GetOut
  END IF

  IF ExitFlag THEN
    RUN "CPMenu"
  ELSEIF OkFlag THEN
    MakeBKGround
    RUN "GLMENU"
  ELSE
    GOTO GetPW
  END IF

GetOut:
CLS
END

