DEFINT A-Z
DECLARE FUNCTION GetDEFPort% ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB OPENDrawer ()

SUB OPENDrawer
  ON LOCAL ERROR GOTO DrawerError
  RecPort = GetDEFPort%
  Port$ = "LPT" + QPTrim$(STR$(RecPort)) + ":"
  PortFile = FREEFILE
  OPEN Port$ FOR OUTPUT AS #PortFile
  PRINT #PortFile, CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250); CHR$(7);

DrawerError:
  RESUME DrawerExit

DrawerExit:
  CLOSE PortFile
END SUB

