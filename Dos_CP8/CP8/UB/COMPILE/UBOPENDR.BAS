DEFINT A-Z
DECLARE FUNCTION GetDEFPort% ()
DECLARE FUNCTION QPTrim$ (Text$)
DECLARE SUB OPENDrawer (RecpPort%)

SUB OPENDrawer (RecpPort)
  ON LOCAL ERROR RESUME NEXT

  'RecPort = GetDEFPort%
  Port$ = "LPT" + QPTrim$(STR$(RecpPort)) + ":"

  PortFile = FREEFILE
  OPEN Port$ FOR OUTPUT AS #PortFile
  PRINT #PortFile, CHR$(27); "p"; CHR$(0); CHR$(25); CHR$(250)
  PRINT #PortFile, CHR$(7)
  CLOSE PortFile

END SUB

