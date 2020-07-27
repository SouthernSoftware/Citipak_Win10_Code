DEFINT A-Z
DECLARE SUB WaitForKey ()
DECLARE FUNCTION Exist% (FileName$)

CONST False = 0, True = NOT False

  CLS
  PRINT "Citipak Payroll W2 Update."
  LOCATE 3, 1


  IF Exist("PRUTIL1.EXE") THEN
    V1 = True
  ELSEIF Exist("PRUTIL9.EXE") THEN
    V2 = True
  ELSE
    NOGood = True
  END IF

  IF NOGood THEN
    PRINT "Unable to find Citipak Data files!"
    PRINT "Updated Aborted!!!"
    LOCATE 6, 1
    PRINT "Please call the software support staff for Assistance."
    LOCATE 10, 1
    PRINT "Press any key to exit. . .";
    WaitForKey
    GOTO AbortExit
  END IF
  PRINT "Installing W2 Update. . ."
  IF V1 THEN
    SHELL "a:prw2-1 -d -o"
  ELSEIF V2 THEN
    SHELL "a:prw2-2 -d -o"
  END IF
  PRINT
  PRINT
  PRINT
  PRINT "W2 Update Completed. . ."
  PRINT
  PRINT
  PRINT "Press any key to exit. . .";
  WaitForKey




AbortExit:

SUB WaitForKey
  a$ = INKEY$
  DO UNTIL LEN(INKEY$) > 0
  LOOP
END SUB

