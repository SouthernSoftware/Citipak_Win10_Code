DEFINT A-Z
DECLARE FUNCTION GetZipEDigit2$ (Zip$)
DECLARE FUNCTION QPTrim$ (text$)
'Zip$ = ""
top:
PRINT
INPUT "Zip Code: ", Zip$
IF LEN(Zip$) = 0 THEN END
PRINT "ZipEDigit: "; GetZipEDigit2$(Zip$)
GOTO top

FUNCTION GetZipEDigit2$ (Zip$)
  
  ZipLen = LEN(Zip$)
  ZipVal = 0
  
  DashPos = INSTR(Zip$, "-")
  DO WHILE DashPos
    Zip$ = LEFT$(Zip$, DashPos - 1) + MID$(Zip$, DashPos + 1)
    DashPos = INSTR(Zip$, "-")
  LOOP
  
  FOR Cnt = 1 TO ZipLen
    ZipVal = ZipVal + VAL(MID$(Zip$, Cnt, 1))
  NEXT
  
  IF ZipVal MOD 10 > 0 THEN
    Dif = 10 - (ZipVal MOD 10)
  ELSE
    Dif = 0
  END IF
  
  GetZipEDigit2$ = QPTrim$(STR$(Dif))
  
END FUNCTION

