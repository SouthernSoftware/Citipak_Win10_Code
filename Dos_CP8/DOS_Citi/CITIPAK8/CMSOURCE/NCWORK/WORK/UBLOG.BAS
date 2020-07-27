DEFINT A-Z
DECLARE FUNCTION QPTrim$ (Text$)

CONST False = 0, True = NOT False

SUB UBLog (Text$) STATIC

'  IF NOT BeenDone THEN
'    BeenDone = True
    Today$ = DATE$
    Today$ = LEFT$(Today$, 2) + MID$(Today$, 4, 2) + RIGHT$(Today$, 2)
'  END IF

  TheTime$ = TIME$
  IF LEFT$(TheTime$, 1) = "0" THEN
    Hour = VAL(MID$(TheTime$, 2, 1))
  ELSE
    Hour = VAL(MID$(TheTime$, 1, 2))
  END IF

  SELECT CASE Hour
  CASE IS > 11
    Hour = Hour - 12
    IF Hour = 0 THEN Hour = 12
    AmPm$ = "pm"
  CASE 1 TO 12
    AmPm$ = "am"
  CASE 0
    Hour = 12
    AmPm$ = "am"
  END SELECT
  SELECT CASE Hour
    CASE 1 TO 9
      Hour$ = "0" + QPTrim$(STR$(Hour))
    CASE ELSE
      Hour$ = QPTrim$(STR$(Hour))
  END SELECT
  TheTime$ = Hour$ + ":" + MID$(TheTime$, 4) + AmPm$
  LogFile = FREEFILE
  OPEN "UBLOG.DAT" FOR APPEND SHARED AS LogFile LEN = 255
  PRINT #LogFile, "UB: "; Today$; " @"; TheTime$
  PRINT #LogFile, "    "; Text$
  CLOSE #LogFile
  Text$ = "": TheTime$ = ""
END SUB

